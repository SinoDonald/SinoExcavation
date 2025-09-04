using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;

namespace SinoExcavation_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class CreateFrame : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        public IList<bool> draw_dir
        {
            get;
            set;
        }

        public bool draw_channel_steel
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            foreach (string file_path in files_path)
            {
                //API CODE START

                try
                {
                    //讀取資料
                    ExReader dex = new ExReader();
                    dex.SetData(file_path, 1);
                    try
                    {
                        dex.PassFrameData();
                        dex.PassBeamData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message); }


                    ICollection<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                    Level levdeep = levels.Where(x => x.Name.Contains("斷面" + dex.section) && !(x.Name.Contains("擋土壁深度"))).OrderBy(x => x.Elevation).ToList()[0];

                    Room room = null;
                    foreach (SpatialElement spelement in new FilteredElementCollector(doc).OfClass(typeof(SpatialElement)).Cast<SpatialElement>().ToList())
                    {
                        Room r = spelement as Room;
                        if (r.Name.Split(' ')[0] == "斷面" + dex.section)
                        {
                            room = r;
                        }
                    }


                    IList<BoundarySegment> boundarySegments = null;
                    int wall_count = 0;
                    if (room.GetBoundarySegments(new SpatialElementBoundaryOptions()).Count != 0)
                    {
                        //撈選該房間之邊界資料
                        boundarySegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions())[0];
                        wall_count = boundarySegments.Count;
                    }
                    else
                    {
                        wall_count = dex.excaRange.Count() - 1;
                    }
                    //取得連續壁內座標點
                    XYZ[] innerwall_points = null;
                    try
                    {
                        innerwall_points = new XYZ[wall_count + 1];
                        for (int i = 0; i < wall_count; i++)
                        {
                            //inner
                            innerwall_points[i] = boundarySegments[i].GetCurve().Tessellate()[0];
                        }
                        innerwall_points[wall_count] = innerwall_points[0];
                    }
                    catch
                    {
                        innerwall_points = new XYZ[wall_count + 1];
                        for (int i = 0; i != dex.excaRange.Count(); i++)
                        {
                            innerwall_points[i] = new XYZ(dex.excaRange[i].Item1, dex.excaRange[i].Item2, 0) * 1000 / 304.8;
                        }
                    }

                    
                    //取得中間樁元件  (先篩選名字 因為會有其他instance沒有特定參數)
                    ICollection<FamilyInstance> columns_instance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                    where x.Name.Contains("中間樁")
                                                                    where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                                                    select x).ToList();


                    IList<XYZ> columns_xyz = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                              where x.Name.Contains("中間樁")
                                              where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                              select (x.Location as LocationPoint).Point).ToList();


                    //取得所有XY數值
                    List<double> Xs = new List<double>();
                    List<double> Ys = new List<double>();
                    double[] slope = new double[wall_count];
                    double[] bias = new double[wall_count];
                    for (int i = 0; i < columns_xyz.Count; i++)
                    {
                        Xs.Add(columns_xyz[i].X);
                        Ys.Add(columns_xyz[i].Y);
                    }
                    for (int i = 0; i < (innerwall_points.Count<XYZ>()); i++)
                    {
                        if (i < slope.Count())
                        {
                            if (innerwall_points[i + 1].X - innerwall_points[i].X == 0)
                            {
                                slope[i] = 20172018;
                                bias[i] = innerwall_points[i + 1].X;
                            }
                            else
                            {
                                slope[i] = (innerwall_points[i + 1].Y - innerwall_points[i].Y) / (innerwall_points[i + 1].X - innerwall_points[i].X);
                                if (slope[i] == 0 || Math.Abs(slope[i]) < 0.0000001)
                                {
                                    bias[i] = innerwall_points[i + 1].Y;
                                }
                                else
                                {
                                    bias[i] = innerwall_points[i + 1].Y - slope[i] * innerwall_points[i + 1].X;
                                }
                            }
                        }
                    }

                    double columns_dis = double.Parse(columns_instance.First().get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString()) * 1000 / 304.8;//透過標註參數讀取中間樁間距
                    double columns_H = double.Parse(columns_instance.First().Symbol.Name.Split('H')[1].Split('x')[0].ToString());//透過中間樁品類讀取中間樁H|||
                    double columns_B = double.Parse(columns_instance.First().Symbol.Name.Split('H')[1].Split('x')[1].ToString());//透過中間樁品類讀取中間樁B---

                    // 判斷中間樁方向 交換H&B
                    if ((columns_instance.First().Location as LocationPoint).Rotation != 0)
                    {
                        (columns_H, columns_B) = (columns_B, columns_H);
                    }

                    // 槽鋼元件 目前固定型號 C200X80
                    FamilySymbol channelSteel_symbol = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                        where x.Name == "C200x80"
                                                        select x).ToList().First();

                    
                    // 槽鋼自動放X向 
                    // 槽鋼端點點位 (不碰到圍囹,中間柱為端點)

                    List<Tuple<XYZ, XYZ>> channelSteel_columnPos = new List<Tuple<XYZ, XYZ>>();

                    if(draw_channel_steel == true)
                    {
                        for (int j = 0; j < Math.Abs(Ys.Max() - Ys.Min()) / columns_dis + 1; j++)
                        {

                            IList<XYZ> columns_y = columns_xyz.Where(pos => Math.Abs(pos.Y - (Ys.Min() + (j) * columns_dis)) < 0.01).OrderBy(pos => pos.X).ToList();
                            int start_idx = 0, end_idx = 0;

                            for (int k = 1; k < columns_y.Count; k++)
                            {
                                double dis = columns_y[k].X - columns_y[k - 1].X;
                                if (dis > columns_dis * 1.1)  // magic !! 用1.0會有小數點計算問題
                                {
                                    if (end_idx > start_idx)
                                    {
                                        var columns_pair = Tuple.Create(columns_y[start_idx], columns_y[end_idx]);
                                        channelSteel_columnPos.Add(columns_pair);
                                    }

                                    start_idx = k;

                                }
                                else
                                {
                                    end_idx = k;
                                    if (end_idx == columns_y.Count - 1)
                                    {
                                        var columns_pair = Tuple.Create(columns_y[start_idx], columns_y[end_idx]);
                                        channelSteel_columnPos.Add(columns_pair);
                                    }

                                }
                            }


                        }
                    }
                    




                    Transaction trans_2 = new Transaction(doc);
                    trans_2.Start("交易開始");
                    string erroemessage = "";
                    for (int lev = 0; lev != dex.supLevel.Count(); lev++)
                    {
                        int temp;
                        if (int.TryParse(dex.supLevel[lev].Item1, out temp) == true)
                        {
                            //開始建立支撐
                            //建立支撐
                            XYZ frame_startpoint = null;
                            XYZ frame_endpoint = null;

                            ICollection<FamilySymbol> frame_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();

                            foreach (FamilySymbol frame_type in frame_familyinstance)
                            {
                                if (frame_type.Name == dex.supLevel[lev].Item4)
                                {
                                    frame_type.Activate();
                                    //讀取圍囹HB
                                    double frame_H = double.Parse(dex.beamLevel[lev].Item3.Split('x')[0].Remove(0, 1));
                                    double frame_B = double.Parse(dex.beamLevel[lev].Item3.Split('x')[1]);
                                    //讀取支撐HB
                                    double Fframe_H = double.Parse(dex.supLevel[lev].Item4.Split('x')[0].Remove(0, 1));
                                    double Fframe_B = double.Parse(dex.supLevel[lev].Item4.Split('x')[1]);
                                    double dou_frame_H;
                                    //單雙向圍囹問題
                                    if (dex.beamLevel[lev].Item2 == 2)
                                    {
                                        dou_frame_H = 2 * frame_H;
                                    }
                                    else
                                    {
                                        dou_frame_H = frame_H;
                                    }
                                    //X向支撐----------
                                    if (draw_dir[0] == true && draw_channel_steel == false)
                                    {
                                        for (int j = 0; j < Math.Abs(Ys.Max() - Ys.Min()) / columns_dis + 1; j++)
                                        {
                                            for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis + 1); i++)
                                            {
                                                XYZ frame_location = new XYZ(Xs.Min() + (i) * columns_dis, Ys.Min() + (j) * columns_dis, 0);
                                                IList<XYZ> frame_intersections = intersection(frame_location, innerwall_points, slope, bias, true);
                                                for (int k = 0; k < (frame_intersections.Count / 2); k++)
                                                {
                                                    try
                                                    {
                                                        frame_startpoint = frame_intersections[k * 2];
                                                        frame_endpoint = frame_intersections[k * 2 + 1];
                                                        Line line = Line.CreateBound(frame_startpoint, frame_endpoint);
                                                        FamilyInstance frame_instance = doc.Create.NewFamilyInstance(line, frame_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                        frame_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-支撐");
                                                        //處理偏移與延伸問題
                                                        frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * -1000).ToString());
                                                        frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-Fframe_B - (columns_H - Fframe_B) / 2).ToString());
                                                        frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-dou_frame_H).ToString());
                                                        frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-dou_frame_H).ToString());
                                                        //取消接合
                                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                                        //若為雙向支撐，則鏡射支撐             
                                                        if (dex.supLevel[lev].Item3 == 2)
                                                        {
                                                            ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                                        }
                                                    }
                                                    catch (Exception e) { erroemessage = new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message; break; };
                                                }
                                                break;
                                            }
                                        }
                                    }
                                    //Y向支撐|||||||
                                    if (draw_dir[1] == true)
                                    {
                                        for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis + 1); i++)
                                        {
                                            for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis + 1); j++)
                                            {
                                                XYZ frame_location = new XYZ(Xs.Min() + (i) * columns_dis, Ys.Min() + (j) * columns_dis, 0);
                                                IList<XYZ> frame_intersections = intersection(frame_location, innerwall_points, slope, bias, false);
                                                for (int k = 0; k < (frame_intersections.Count / 2); k++)
                                                {
                                                    try
                                                    {
                                                        frame_startpoint = frame_intersections[k * 2];
                                                        frame_endpoint = frame_intersections[k * 2 + 1];
                                                        Line line = Line.CreateBound(frame_startpoint, frame_endpoint);
                                                        FamilyInstance frame_instance = doc.Create.NewFamilyInstance(line, frame_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                        frame_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-支撐");

                                                        //處理偏移與延伸問題
                                                        //double offset =double.Parse(frame_instance.LookupParameter("H").AsValueString());
                                                        frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * -1000 + (frame_B + Fframe_H) / 2).ToString());//2000為支撐階數深度，表1中
                                                        frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-Fframe_B - (columns_B - Fframe_B) / 2).ToString());
                                                        frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-dou_frame_H).ToString());
                                                        frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-dou_frame_H).ToString());
                                                        //取消接合
                                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                                        StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                                        //若為雙向支撐，則鏡射支撐
                                                        if (dex.supLevel[lev].Item3 == 2)
                                                        {
                                                            ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                                        }
                                                    }
                                                    catch (Exception e) { erroemessage = new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message; break; };
                                                }
                                                break;
                                            }
                                        }
                                    }


                                    //X向槽鋼

                                    if (draw_dir[0] == true && draw_channel_steel == true)
                                    {
                                        foreach (Tuple<XYZ, XYZ> column_pair in channelSteel_columnPos)
                                        {
                                            try
                                            {
                                                XYZ channel_steel_start = new XYZ(column_pair.Item1.X, column_pair.Item1.Y + columns_H / 2 / 304.8, 0);
                                                XYZ channel_steel_end = new XYZ(column_pair.Item2.X, column_pair.Item2.Y + columns_H / 2 / 304.8, 0);
                                                Line line = Line.CreateBound(channel_steel_start, channel_steel_end);
                                                FamilyInstance channelSteel_instance = doc.Create.NewFamilyInstance(line, channelSteel_symbol, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);

                                                channelSteel_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString(((dex.supLevel[lev].Item2) * -1000).ToString());

                                                //取消接合
                                                StructuralFramingUtils.DisallowJoinAtEnd(channelSteel_instance, 0);
                                                StructuralFramingUtils.DisallowJoinAtEnd(channelSteel_instance, 1);
                                                //鏡像
                                                ElementTransformUtils.MirrorElement(doc, channelSteel_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, column_pair.Item1));
                                            }
                                            catch (Exception e) { erroemessage = new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message; break; };

                                        }

                                    }


                                }
                            }
                        }
                    }
                    trans_2.Commit();
                }


                catch (Exception e) { TaskDialog.Show("error", e.ToString()); break; }
            }
            TaskDialog.Show("done", "支撐建置完畢");
        }
        public static bool IsInPolygon(XYZ point, XYZ[] polygon)//判斷點是否位於開挖範圍內
        {
            bool result = false;
            var a = polygon.Last();

            foreach (var b in polygon)
            {
                if ((b.X == point.X) && (b.Y == point.Y))
                {
                    return true;
                }
                if ((b.Y == a.Y) && (point.Y == a.Y) && (a.X <= point.X) && (point.X <= b.X))
                {
                    return true;
                }
                if ((b.X == a.X) && (point.X == a.X) && (a.Y <= point.Y) && (point.Y <= b.Y))
                {
                    return true;
                }
                if (((b.Y < point.Y) && (a.Y >= point.Y) || (a.Y < point.Y) && (b.Y >= point.Y)))
                {

                    bool check = false;
                    bool check2 = false;
                    double h = b.X + (point.Y - b.Y) / (a.Y - b.Y) * (a.X - b.X) - (point.X);

                    if (h > -0.1 && h < 0.1)
                    {
                        check = IsInPolygon(new XYZ(point.X + 0.5, point.Y, 0), polygon);
                        check2 = IsInPolygon(new XYZ(point.X - 0.5, point.Y, 0), polygon);
                    }
                    if (check && !check2)
                    {

                        if (((b.X + (point.Y - b.Y) / (a.Y - b.Y) * (a.X - b.X)) >= (point.X) + 0.1) && point.X != b.X && point.X != a.X)
                        {
                            result = !result;
                        }
                    }
                    else if (!check && !check2)
                    {
                        if (((b.X + (point.Y - b.Y) / (a.Y - b.Y) * (a.X - b.X)) >= (point.X)))
                        {
                            result = !result;
                        }
                    }
                    else
                    {
                        if (((b.X + (point.Y - b.Y) / (a.Y - b.Y) * (a.X - b.X)) >= (point.X) - 0.1) && point.X != b.X && point.X != a.X)
                        {
                            result = !result;
                        }
                    }
                }

                a = b;
            }
            return result;
        }
        public static IList<XYZ> intersection(XYZ checkPoint, XYZ[] polygonPoints, double[] slope, double[] bias, bool x_or_y)//計算支撐與連續壁交點，true為計算X向，false為計算Y向
        {
            List<XYZ> intersection_points = new List<XYZ>();
            List<XYZ> final_points = new List<XYZ>();
            List<double> x = new List<double>();
            List<double> y = new List<double>();
            if (x_or_y == true)
            {
                for (int i = 0; i < slope.Count<double>(); i++)
                {
                    if (slope[i] != 0 || Math.Abs(slope[i]) > 0.00001)
                    {
                        if (checkPoint.Y >= polygonPoints[i].Y && checkPoint.Y <= polygonPoints[i + 1].Y || checkPoint.Y <= polygonPoints[i].Y && checkPoint.Y >= polygonPoints[i + 1].Y)
                        {
                            if (slope[i] == 20172018)
                            {
                                intersection_points.Add(new XYZ(bias[i], checkPoint.Y, 0));
                            }
                            else
                            {
                                intersection_points.Add(new XYZ((checkPoint.Y - bias[i]) / slope[i], checkPoint.Y, 0));
                            }
                        }
                    }
                }

                foreach (XYZ point in intersection_points)
                {
                    if (IsInPolygon(point, polygonPoints))
                    {
                        final_points.Add(point);
                    }
                }
                final_points = final_points.OrderBy(p => p.X).ToList();
            }

            else//若為False則計算Y向
            {
                for (int i = 0; i < slope.Count<double>(); i++)
                {
                    if (slope[i] != 20172018)
                    {
                        if (checkPoint.X >= polygonPoints[i].X && checkPoint.X <= polygonPoints[i + 1].X || checkPoint.X <= polygonPoints[i].X && checkPoint.X >= polygonPoints[i + 1].X)
                        {
                            if (slope[i] == 0 || Math.Abs(slope[i]) < 0.00001)
                            {

                                intersection_points.Add(new XYZ(checkPoint.X, bias[i], 0));
                            }
                            else
                            {
                                intersection_points.Add(new XYZ(checkPoint.X, slope[i] * checkPoint.X + bias[i], 0));
                            }
                        }
                    }
                }

                foreach (XYZ point in intersection_points)
                {
                    //if (IsInPolygon(point, polygonPoints))
                    //{
                    final_points.Add(point);
                    //}
                }
                final_points = final_points.OrderBy(p => p.Y).ToList();

            }
            return final_points;

        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
