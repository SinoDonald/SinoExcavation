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

    class CreateSlope : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        public IList<bool> draw_direc
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
                    IList<Level> all_level_list = levels.Where(x => x.Name.Contains("斷面" + dex.section) && !(x.Name.Contains("擋土壁深度"))).ToList();


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
                    //取得中間樁元件
                    ICollection<FamilyInstance> columns_instance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList()
                                                                    where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                                                    select x).ToList();
                    IList<XYZ> columns_xyz = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList()
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
                            if (Math.Abs(innerwall_points[i + 1].X - innerwall_points[i].X) <= 0.1)
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

                    Transaction trans_2 = new Transaction(doc);
                    trans_2.Start("交易開始");
                    //string erroemessage = "";
                    for (int lev = 0; lev != dex.supLevel.Count(); lev++)
                    {
                        int temp;
                        if (int.TryParse(dex.supLevel[lev].Item1, out temp) == true)
                        {
                            XYZ frame_startpoint = null;
                            XYZ frame_endpoint = null;
                            string[] size_list = dex.supLevel[lev].Item4.Split('x');
                            size_list[0] = size_list[0].Remove(0, 1);

                            //建立斜撐
                            ICollection<FamilySymbol> slopframe_symbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                            foreach (FamilySymbol slopframe_type in slopframe_symbol)
                            {
                                if (slopframe_type.Name == "斜撐")
                                {
                                    double beam_h = double.Parse(dex.beamLevel[lev].Item3.Split('x')[0].Remove(0, 1));
                                    double cal_beam_h;
                                    if (dex.beamLevel[lev].Item2 == 1)
                                    {
                                        cal_beam_h = beam_h;
                                    }
                                    else
                                    {
                                        cal_beam_h = 2 * beam_h;
                                    }
                                    //X向斜撐
                                    if (draw_direc[0] == true)
                                    {
                                        for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis + 1); j++)
                                        {
                                            for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis + 1); i++)
                                            {
                                                XYZ frame_location = new XYZ(Xs.Min() + (i) * columns_dis, Ys.Min() + (j) * columns_dis, 0);
                                                IList<XYZ> frame_intersections = intersection(frame_location, innerwall_points, slope, bias, true);
                                                for (int k = 0; k < (frame_intersections.Count / 2); k++)
                                                {
                                                    frame_startpoint = frame_intersections[k * 2];
                                                    frame_endpoint = frame_intersections[k * 2 + 1];
                                                    FamilyInstance slopframe_1 = doc.Create.NewFamilyInstance(frame_startpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                    FamilyInstance slopframe_2 = doc.Create.NewFamilyInstance(frame_endpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                    slopframe_1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-斜撐");
                                                    slopframe_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-斜撐");
                                                    //寫入尺寸
                                                    slopframe_1.LookupParameter("H").SetValueString(size_list[0]);
                                                    slopframe_2.LookupParameter("H").SetValueString(size_list[0]);
                                                    slopframe_1.LookupParameter("B").SetValueString(size_list[1]);
                                                    slopframe_2.LookupParameter("B").SetValueString(size_list[1]);
                                                    slopframe_1.LookupParameter("tw").SetValueString(size_list[2]);
                                                    slopframe_2.LookupParameter("tw").SetValueString(size_list[2]);
                                                    slopframe_1.LookupParameter("tf").SetValueString(size_list[3]);
                                                    slopframe_2.LookupParameter("tf").SetValueString(size_list[3]);
                                                    slopframe_1.LookupParameter("支撐厚度").SetValueString(size_list[0]);
                                                    slopframe_2.LookupParameter("支撐厚度").SetValueString(size_list[0]);
                                                    slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - dex.supLevel[lev].Item2 * -1000 / 304.8));

                                                    slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - dex.supLevel[lev].Item2 * -1000 / 304.8));
                                                    
                                                    slopframe_1.LookupParameter("圍囹寬度").SetValueString((cal_beam_h).ToString());
                                                    slopframe_2.LookupParameter("圍囹寬度").SetValueString((cal_beam_h).ToString());
                                                    //旋轉斜撐元件
                                                    Line rotate_line_s = Line.CreateBound(frame_startpoint, frame_startpoint + new XYZ(0, 0, 1));
                                                    Line rotate_line_e = Line.CreateBound(frame_endpoint, frame_endpoint + new XYZ(0, 0, 1));
                                                    slopframe_1.Location.Rotate(rotate_line_s, 1.5 * Math.PI);
                                                    slopframe_2.Location.Rotate(rotate_line_e, 0.5 * Math.PI);

                                                    //鏡射斜撐元件
                                                    if (dex.supLevel[lev].Item3 == 2)//雙排
                                                    {
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_endpoint));
                                                    }
                                                    else//單排
                                                    {
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint.Add(new XYZ(0, -(slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2 + slopframe_1.LookupParameter("中間樁直徑").AsDouble() / 2), 0))));
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_endpoint));
                                                        slopframe_2.Location.Move((new XYZ(0, -(slopframe_1.LookupParameter("支撐厚度").AsDouble() + slopframe_1.LookupParameter("中間樁直徑").AsDouble()), 0)));

                                                    }
                                                }
                                                break;

                                            }
                                        }
                                    }
                                    //Y向斜撐
                                    if (draw_direc[1] == true)
                                    {
                                        for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis + 1); i++)
                                        {
                                            for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis + 1); j++)
                                            {
                                                XYZ frame_location = new XYZ(Xs.Min() + (i) * columns_dis, Ys.Min() + (j) * columns_dis, 0);
                                                IList<XYZ> frame_intersections = intersection(frame_location, innerwall_points, slope, bias, false);
                                                for (int k = 0; k < (frame_intersections.Count / 2); k++)
                                                {
                                                    frame_startpoint = frame_intersections[k * 2];
                                                    frame_endpoint = frame_intersections[k * 2 + 1];
                                                    FamilyInstance slopframe_1 = doc.Create.NewFamilyInstance(frame_startpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                    FamilyInstance slopframe_2 = doc.Create.NewFamilyInstance(frame_endpoint, slopframe_type, all_level_list[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                                    slopframe_1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-斜撐");
                                                    slopframe_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-斜撐");
                                                    //寫入尺寸
                                                    slopframe_1.LookupParameter("H").SetValueString(size_list[0]);
                                                    slopframe_2.LookupParameter("H").SetValueString(size_list[0]);
                                                    slopframe_1.LookupParameter("B").SetValueString(size_list[1]);
                                                    slopframe_2.LookupParameter("B").SetValueString(size_list[1]);
                                                    slopframe_1.LookupParameter("tw").SetValueString(size_list[2]);
                                                    slopframe_2.LookupParameter("tw").SetValueString(size_list[2]);
                                                    slopframe_1.LookupParameter("tf").SetValueString(size_list[3]);
                                                    slopframe_2.LookupParameter("tf").SetValueString(size_list[3]);
                                                    slopframe_1.LookupParameter("支撐厚度").SetValueString(size_list[0]);
                                                    slopframe_2.LookupParameter("支撐厚度").SetValueString(size_list[0]);
                                                    slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - dex.supLevel[lev].Item2 * (-1000) / 304.8 - double.Parse(size_list[0]) / 304.8));

                                                    slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(all_level_list[lev].Elevation - dex.supLevel[lev].Item2 * (-1000) / 304.8 - double.Parse(size_list[0]) / 304.8));
                                                    //slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(894  / 304.8);

                                                    //slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(894  / 304.8);
                                                    
                                                    slopframe_1.LookupParameter("圍囹寬度").SetValueString((cal_beam_h).ToString());
                                                    slopframe_2.LookupParameter("圍囹寬度").SetValueString((cal_beam_h).ToString());
                                                    //旋轉斜撐元件
                                                    Line rotate_line = Line.CreateBound(frame_endpoint, frame_endpoint + new XYZ(0, 0, 1));
                                                    slopframe_2.Location.Rotate(rotate_line, Math.PI);

                                                    //鏡射斜撐元件
                                                    if (dex.supLevel[lev].Item3 == 2)//雙排
                                                    {
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_endpoint));
                                                    }
                                                    else//單排
                                                    {
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint.Add(new XYZ(slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2 + slopframe_1.LookupParameter("中間樁直徑").AsDouble() / 2, 0, 0))));
                                                        ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_endpoint));
                                                        slopframe_2.Location.Move((new XYZ((slopframe_1.LookupParameter("支撐厚度").AsDouble() + slopframe_1.LookupParameter("中間樁直徑").AsDouble()), 0, 0)));

                                                    }
                                                }
                                                break;

                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    IList<FamilyInstance> familyInstances_slope = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Category.Name == "結構襯料").ToList();
                    List<XYZ> zero = new List<XYZ>();
                    zero.Add(new XYZ(-8000, -8000, 0));
                    foreach (FamilyInstance slopeelement in familyInstances_slope)
                    {
                        if (calculator_del((slopeelement.Location as LocationPoint).Point, zero))
                        {
                            doc.Delete(slopeelement.Id);
                        }
                    }
                    trans_2.Commit();
                }


                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); break; }
            }
            TaskDialog.Show("done", "斜撐建置完畢");
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
        public static bool calculator(XYZ checkPoint, List<XYZ> ranoutpoints)//判斷點與點是否為相同點，因為xyz類型無法做equals
        {
            bool a = false;
            foreach (XYZ ranoutpoint in ranoutpoints)
            {
                if (ranoutpoint.X - checkPoint.X < 1 && ranoutpoint.X - checkPoint.X > -1)
                {
                    if (ranoutpoint.Y - checkPoint.Y < 1 && ranoutpoint.Y - checkPoint.Y > -1)
                    {
                        a = true;
                    }
                }
            }
            return a;
        }
        public static bool calculator_del(XYZ checkPoint, List<XYZ> ranoutpoints)//判斷誤差值(刪除斜邊斜撐用)
        {
            bool a = false;
            foreach (XYZ ranoutpoint in ranoutpoints)
            {
                if (ranoutpoint.X - checkPoint.X < 10 && ranoutpoint.X - checkPoint.X > -10)
                {
                    if (ranoutpoint.Y - checkPoint.Y < 10 && ranoutpoint.Y - checkPoint.Y > -10)
                    {
                        a = true;
                    }
                }
            }
            return a;
        }
        public static IList<XYZ> intersection(XYZ checkPoint, XYZ[] polygonPoints, double[] slope, double[] bias, bool x_or_y)//計算支撐與連續壁交點，true為計算X向，false為計算Y向
        {
            List<XYZ> intersection_points = new List<XYZ>();
            List<XYZ> slope_non_points = new List<XYZ>();
            List<XYZ> final_points = new List<XYZ>();
            List<XYZ> checked_final_points = new List<XYZ>();
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
                                slope_non_points.Add(new XYZ((checkPoint.Y - bias[i]) / slope[i], checkPoint.Y, 0));
                            }
                        }
                    }
                }

                foreach (XYZ point in intersection_points)
                {/*
                    if (IsInPolygon(point, polygonPoints))
                    {*/
                        final_points.Add(point);
                    //}
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
                            if (slope[i] == 0 || Math.Abs(slope[i]) < 0.04)
                            {

                                intersection_points.Add(new XYZ(checkPoint.X, bias[i], 0));
                            }
                            else
                            {
                                intersection_points.Add(new XYZ(checkPoint.X, slope[i] * checkPoint.X + bias[i], 0));
                                slope_non_points.Add(new XYZ(checkPoint.X, slope[i] * checkPoint.X + bias[i], 0));
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

            //標記斜邊斜撐
            foreach (XYZ point in final_points)
            {
                if (calculator(point, slope_non_points))
                {
                    checked_final_points.Add(new XYZ(-8000, -8000, 0));
                }
                else
                {
                    checked_final_points.Add(point);
                }
            }
            return checked_final_points;

        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
