using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;

namespace SinoExcavation_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class CreateTriSupportofBeam : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        public IList<bool> elevation_decide
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
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); }


                    //取得三角托架元件
                    FamilySymbol big_trisupport = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                               where x.Name == "大三角架"
                                               select x).First();

                    FamilySymbol small_trisupport = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                   where x.Name == "三角架"
                                                   select x).First();

                    //FamilySymbol trisupport = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "TriSupport_test").Select(x => x).ToList().First();
                    //FamilySymbol trisupport = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "TriSupport_test").First();

                    ICollection<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                    Level levdeep = levels.Where(x => x.Name.Contains("斷面" + dex.section + "-" + "開挖階數1")).ToList().First();


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
                            innerwall_points[i] = new XYZ(innerwall_points[i].X, innerwall_points[i].Y, XYZ.Zero.Z);
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

                    Transaction trans_2 = new Transaction(doc);
                    trans_2.Start("交易開始");

                    ////啟用三角托架元件
                    //if (!trisupport.IsActive)
                    //{
                    //    trisupport.Activate();
                    //    //doc.Regenerate();
                    //}


                    for (int lev = 0; lev != dex.beamLevel.Count(); lev++)
                    {
                        XYZ beam_slope = XYZ.Zero;
                        XYZ new_beam_slope = XYZ.Zero;
                        bool turn_side = true; //會是Y向高程
                        if (elevation_decide[1] == false) //判斷是為X向 or Y向重疊
                        {
                            turn_side = false;  //會是X向高程
                        }

                        //建立圍囹
                        //開始建立圍囹
                        ICollection<FamilySymbol> beam_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                        foreach (FamilySymbol beam_type in beam_familyinstance)
                        {
                            if (beam_type.Name == dex.beamLevel[lev].Item3)
                            {
                                //beam_type.Activate();
                                double beam_H = double.Parse(beam_type.LookupParameter("H").AsValueString());
                                double beam_B = double.Parse(beam_type.LookupParameter("B").AsValueString());

                                FamilySymbol trisupport;

                                if (dex.beamLevel[lev].Item2 == 2)
                                {
                                    trisupport = big_trisupport;
                                }
                                else
                                {
                                    trisupport = small_trisupport;
                                }


                                //讀點位
                                    for (int i = 0; i < innerwall_points.Count<XYZ>() - 1; i++)
                                {

                                    Curve c = null;
                                    if (i == innerwall_points.Count<XYZ>())
                                    {
                                        try
                                        {
                                            int.Parse(dex.beamLevel[lev].Item1.ToString());
                                            c = Line.CreateBound(innerwall_points[i], innerwall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1));
                                        }
                                        catch { /*c = Line.CreateBound(sidewall_points[i], sidewall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1));*/ }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            int.Parse(dex.beamLevel[lev].Item1.ToString());
                                            c = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1));
                                        }

                                        catch { /*c = Line.CreateBound(sidewall_points[i], sidewall_points[i+1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1));*/ }
                                    }
                                    if (c != null)
                                    {
                                        //FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                        //beam.LookupParameter("斷面旋轉").SetValueString("90");
                                        //StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                                        //StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                                        //beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.beamLevel[lev].Item1).ToString() + "-圍囹");


                                        //判斷圍囹之垂直深度，斜率零為負，反之為正
                                        //new_beam_slope = ((beam.Location as LocationCurve).Curve as Line).Direction;
                                        new_beam_slope = (c as Line).Direction;
                                        double angle = beam_slope.AngleTo(new_beam_slope) * 180 / Math.PI;

                                        if (elevation_decide[0] == true) // 搭接
                                        {
                                            if (i == 0)
                                            {
                                                double angle_start = innerwall_points[i].AngleTo(innerwall_points[i + 1]) * 180 / Math.PI;
                                                if (Math.Abs(angle_start) > 10)
                                                {
                                                    turn_side = !turn_side;
                                                }
                                            }
                                            if (i != 0 && angle > 10)
                                            {
                                                turn_side = !turn_side;
                                            }
                                        }

                                        //if (turn_side)
                                        //{
                                        //    beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 - beam_B / 2).ToString());//2000為支撐階數深度，表1中
                                        //}
                                        //else
                                        //{
                                        //    beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 + beam_B / 2).ToString());
                                        //}

                                        //三角支撐

                                        double triSupport_depth;
                                        double triSupport_distance = c.Length;      //線段總長
                                        double triSupport_spacing = 3000 / 304.8;   //間距

                                        //double triSupport_start = 3000 / 304.8;     //起始距離
                                        //double truSupport_end = 100 / 304.8;  //末項擺放閾值
                                        //double triSupport_width = 75 / 304.8;

                                        double min_dis = (beam_H + (75 / 2)) / 304.8; //頭尾避免重疊最小距離  (75 其中一種size

                                        //線段總長大於 min 才擺放
                                        if (triSupport_distance > min_dis)
                                        {
                                            // 深度:從階層圍囹深度 + turn_side(搭街)計算
                                            if (turn_side) triSupport_depth = (dex.supLevel[lev].Item2 * 1000) / -304.8;
                                            else triSupport_depth = (dex.supLevel[lev].Item2 * 1000 + beam_B) / -304.8;

                                            //旋轉角度: 從圍囹方向向量計算 
                                            double triSupport_angle;
                                            if (new_beam_slope.Y >= 0) triSupport_angle = new XYZ(1, 0, 0).AngleTo(new_beam_slope);
                                            else triSupport_angle = new XYZ(-1, 0, 0).AngleTo(new_beam_slope) + Math.PI;


                                            for (int j = 0; j <= Math.Floor((triSupport_distance - min_dis) / triSupport_spacing); j++)
                                            {

                                                trisupport.Activate();
                                                XYZ pos = new XYZ(innerwall_points[i].X, innerwall_points[i].Y, triSupport_depth) + new_beam_slope * (min_dis + triSupport_spacing * j);

                                                FamilyInstance instance = doc.Create.NewFamilyInstance(pos, trisupport, StructuralType.NonStructural);

                                                XYZ vpos = pos + new XYZ(0, 0, 1);
                                                Line axis = Line.CreateBound(pos, vpos);

                                                ElementTransformUtils.RotateElement(doc, instance.Id, axis, triSupport_angle + Math.PI);
                                            }

                                            //最後一項

                                            XYZ last_pos = new XYZ(innerwall_points[i + 1].X, innerwall_points[i + 1].Y, triSupport_depth) - new_beam_slope * min_dis;

                                            FamilyInstance last_instance = doc.Create.NewFamilyInstance(last_pos, trisupport, StructuralType.NonStructural);

                                            XYZ vlast_pos = last_pos + new XYZ(0, 0, 1);
                                            Line last_axis = Line.CreateBound(last_pos, vlast_pos);

                                            ElementTransformUtils.RotateElement(doc, last_instance.Id, last_axis, triSupport_angle + Math.PI);

                                        }

                                        ////線段總長大於間距才擺放
                                        //if (triSupport_distance >= triSupport_spacing)
                                        //{
                                        //    // 深度:從階層圍囹深度 + turn_side(搭街)計算
                                        //    if (turn_side) triSupport_depth = (dex.supLevel[lev].Item2 * 1000) / -304.8;
                                        //    else triSupport_depth = (dex.supLevel[lev].Item2 * 1000 + beam_B) / -304.8;

                                        //    double triSupport_angle; //旋轉角度: 從圍囹方向向量計算 
                                        //    if (new_beam_slope.Y >= 0) triSupport_angle = new XYZ(1, 0, 0).AngleTo(new_beam_slope);
                                        //    else triSupport_angle = new XYZ(-1, 0, 0).AngleTo(new_beam_slope) + Math.PI;


                                        //    for (int j = 0; j <= Math.Floor((triSupport_distance - triSupport_start) / triSupport_spacing); j++)
                                        //    {

                                        //        trisupport.Activate();
                                        //        FamilyInstance instance = doc.Create.NewFamilyInstance(
                                        //            new XYZ(innerwall_points[i].X, innerwall_points[i].Y, triSupport_depth)
                                        //            + new_beam_slope * (triSupport_spacing * j + triSupport_start)
                                        //            , trisupport, StructuralType.NonStructural);

                                        //        XYZ point1 = innerwall_points[i] + new_beam_slope * (triSupport_spacing * j + triSupport_start);
                                        //        XYZ point2 = point1 + new XYZ(0, 0, 1);
                                        //        Line axis = Line.CreateBound(point1, point2);

                                        //        ElementTransformUtils.RotateElement(doc, instance.Id, axis, triSupport_angle + Math.PI);
                                        //    }

                                        //}


                                    }

                                    //if (dex.beamLevel[lev].Item2 == 2)  // 圍囹隻數
                                    //{
                                    //    c = null;
                                    //    if (i == innerwall_points.Count<XYZ>())
                                    //    {
                                    //        try
                                    //        {
                                    //            int.Parse(dex.beamLevel[lev].Item1.ToString());
                                    //            c = Line.CreateBound(innerwall_points[i], innerwall_points[0]).CreateOffset((2 * beam_H) / 304.8, new XYZ(0, 0, -1));  // offset H 再長一隻
                                    //        }
                                    //        catch { /*c = Line.CreateBound(sidewall_points[i], sidewall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1));*/ }

                                    //    }
                                    //    else
                                    //    {
                                    //        try
                                    //        {
                                    //            int.Parse(dex.beamLevel[lev].Item1.ToString());
                                    //            c = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]).CreateOffset((2 * beam_H) / 304.8, new XYZ(0, 0, -1));
                                    //        }
                                    //        catch { /*c = Line.CreateBound(sidewall_points[i], sidewall_points[i+1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1));*/ }
                                    //    }
                                    //    if (c != null)
                                    //    {
                                    //        FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    //        beam.LookupParameter("斷面旋轉").SetValueString("90");
                                    //        StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                                    //        StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                                    //        beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.beamLevel[lev].Item1).ToString() + "-圍囹");

                                    //        if (turn_side)
                                    //        {
                                    //            beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 - beam_B / 2).ToString());//2000為支撐階數深度，表1中
                                    //        }
                                    //        else
                                    //        {
                                    //            beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 + beam_B / 2).ToString());
                                    //        }

                                    //    }
                                    //}
                                    beam_slope = new_beam_slope;
                                }
                            }
                        }
                    }

                    trans_2.Commit();
                }


                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); break; }
            }
            TaskDialog.Show("done", "圍囹建置完畢");
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
