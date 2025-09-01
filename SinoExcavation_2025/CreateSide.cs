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
    class CreateSide : IExternalEventHandler
    {
        public IList<string> files_path
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
                        dex.PassSideData();
                        dex.PassBeamData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("error", e.StackTrace + e.Message + e.Source); }
                    //偏移量

                    Transaction trans = new Transaction(doc);
                    trans.Start("side_wall");
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

                    //取的側牆profile
                    IList<Curve> side_wall_profileloops = new List<Curve>();
                    for (int i = 0; i < innerwall_points.Count() - 1; i++)
                    {
                        Line line = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]);
                        side_wall_profileloops.Add(line);
                    }

                    //建立樓板回築樓層
                    Level[] re_levlist = new Level[dex.back.Count()];
                    for (int i = 0; i != dex.back.Count(); i++)
                    {
                        re_levlist[i] = Level.Create(doc, (dex.back[i].Item2 * -1 + dex.back[i].Item3 / 2) * 1000 / 304.8);
                        re_levlist[i].Name = ("斷面"+dex.section+"-"+dex.back[i].Item1).ToString();
                    }

                    //建立側牆
                    
                    List<Wall> side_wall = new List<Wall>();
                    ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
                    WallType wallType = null;
                   
                    for (int i = 0; i < dex.sideWall.Count; i++)
                    {
                        //檢查側牆
                        if (walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "側牆" && x.Name.Split('-')[1] == (dex.sideWall[i].Item2 * 1000).ToString() + "mm").ToList().Count != 0)
                            wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "側牆" && x.Name.Split('-')[1] == (dex.sideWall[i].Item2 * 1000).ToString() + "mm").ToList().First();
                        //建立側牆新類型
                        if (wallType == null)
                        {
                            wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "側牆").ToList().First();
                            WallType new_wallFamSym = wallType.Duplicate("側牆-" + dex.sideWall[i].Item2 * 1000 + "mm") as WallType;
                            CompoundStructure ly = new_wallFamSym.GetCompoundStructure();
                            ly.SetLayerWidth(0, dex.sideWall[i].Item2 * 1000 / 304.8);
                            new_wallFamSym.SetCompoundStructure(ly);
                            wallType = new_wallFamSym;
                        }

                        double floor_width = dex.back[i].Item3 * 1000 / 304.8;// set the width to a new value:
                        double floor_deep = -dex.back[i].Item2 * 1000 / 304.8;
                        double floor_offset = (floor_deep + floor_width / 2);
                        for (int j = 0; j < innerwall_points.Count() - 1; j++)
                        {
                            Curve side_c = side_wall_profileloops[j].CreateOffset((dex.sideWall[i].Item2 * 1000 / 2) / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線
                            Wall side_w = Wall.Create(doc, side_c, wallType.Id, re_levlist[i].Id, (floor_offset * (-1) - (dex.back[i + 1].Item2 + dex.back[i + 1].Item3 / 2) * 1000 / 304.8), 0, false, false);
                            WallUtils.DisallowWallJoinAtEnd(side_w, 0);
                            WallUtils.DisallowWallJoinAtEnd(side_w, 1);
                            side_w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + dex.sideWall[i].Item1 + "側壁");
                            side_w.get_Parameter(BuiltInParameter.DOOR_NUMBER).Set(dex.sideWall[i].Item3.ToString() );
                            side_wall.Add(side_w);
                        }
                    }
                    trans.Commit();

                    Transaction trans2 = new Transaction(doc);
                    
                    trans2.Start("Create Base");
                   
                    //取得側牆牆點座標
                    CurveArray profileloops_array = new CurveArray();
                    XYZ[] sidewall_points = new XYZ[innerwall_points.Count<XYZ>()];
                    for (int i = 0; i < (innerwall_points.Count<XYZ>() - 1); i++)
                    {
                        //side
                        XYZ side_wall_curve_point = (side_wall[i].Location as LocationCurve).Curve.Tessellate()[0];
                        side_wall_curve_point = new XYZ(side_wall_curve_point.X, side_wall_curve_point.Y, 0);
                        sidewall_points[i] = innerwall_points[i] - (innerwall_points[i] - side_wall_curve_point) * 2;
                        sidewall_points[i] = new XYZ(sidewall_points[i].X, sidewall_points[i].Y, 0);
                    }
                    sidewall_points[innerwall_points.Count<XYZ>() - 1] = sidewall_points[0];
                    
                    //建立底板範圍
                    for (int i = 0; i < innerwall_points.Count() - 1; i++)
                    {
                        Line line = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]);
                        profileloops_array.Append(line);
                    }
                    
                    //建立底板類型及實作元件
                    ICollection<FloorType> family = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().ToList();
                    FloorType floor_type = family.Where(x => x.Name == "通用 300mm").First();                    
                    int base_control = 0;
                    foreach (Tuple<string, double, double, string> base_tuple in dex.back)
                    {
                        FloorType newFamSym = null;
                        double floor_width = base_tuple.Item3 * 1000 / 304.8;// set the width to a new value:
                        double floor_deep = base_tuple.Item2 * -1 * 1000 / 304.8;
                        double floor_offset = (floor_deep + floor_width / 2) * 304.8;
                        //try { newFamSym = floor_type.Duplicate(base_tuple.Item1+"-"+ dex.section) as FloorType; } catch { newFamSym = family.Where(x => x.Name == base_tuple.Item1).First(); }
                        try { 
                            newFamSym = floor_type.Duplicate(base_tuple.Item1 + "-" + dex.section) as FloorType; 
                        } 
                        catch {
                            newFamSym = family.Where(x => x.Name.Contains(base_tuple.Item1)).First(); 
                        }

                        CompoundStructure ly = newFamSym.GetCompoundStructure();
                        ly.SetLayerWidth(0, floor_width);
                        newFamSym.SetCompoundStructure(ly);
                        //Floor floor = doc.Create.NewFloor(profileloops_array, newFamSym as FloorType, re_levlist[base_control], false);

                        // 培文改
                        IList<Curve> curves = new List<Curve>();
                        for (int i = 0; i < innerwall_points.Count() - 1; i++)
                        {
                            Line line = Line.CreateBound(innerwall_points[i], innerwall_points[i + 1]);
                            curves.Append(line);
                        }
                        CurveLoop curveLoop = CurveLoop.Create(curves);
                        IList<CurveLoop> curveLoops = new List<CurveLoop>();
                        curveLoops.Add(curveLoop);
                        Element floor = Floor.Create(doc, curveLoops, newFamSym.Id, Level.Create(doc, 0).Id);

                        floor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + base_tuple.Item1);
                        floor.get_Parameter(BuiltInParameter.DOOR_NUMBER).Set(base_tuple.Item4.ToString());
                        floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).SetValueString("0");
                        base_control += 1;
                    }
                    trans2.Commit();
                    
                    ICollection<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                    Level levdeep = levels.Where(x => x.Name.Contains("斷面" + dex.section + "-" + "開挖階數1")).ToList().First();

                    trans2.Start("建置回撐圍囹");
                    for (int lev = 0; lev != dex.supLevel.Count(); lev++)
                    {
                        //建立圍囹
                        //開始建立圍囹
                        ICollection<FamilySymbol> beam_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                        foreach (FamilySymbol beam_type in beam_familyinstance)
                        {
                            if (beam_type.Name == dex.supLevel[lev].Item4)
                            {
                                beam_type.Activate();
                                double beam_H = double.Parse(beam_type.LookupParameter("H").AsValueString());
                                double beam_B = double.Parse(beam_type.LookupParameter("B").AsValueString());
                                for (int i = 0; i < innerwall_points.Count<XYZ>() - 1; i++)
                                {
                                    Curve c = null;
                                    if (i == innerwall_points.Count<XYZ>())
                                    {
                                        try
                                        {
                                            int.Parse(dex.beamLevel[lev].Item1.ToString());
                                        }
                                        catch { c = Line.CreateBound(sidewall_points[i], sidewall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }

                                    }
                                    else
                                    {
                                        try
                                        {
                                            int.Parse(dex.beamLevel[lev].Item1.ToString());
                                        }
                                        catch { c = Line.CreateBound(sidewall_points[i], sidewall_points[i + 1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }
                                    }
                                    if (c != null)
                                    {
                                        FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                        beam.LookupParameter("斷面旋轉").SetValueString("90");
                                        StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                                        StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                                        beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.beamLevel[lev].Item1).ToString() + "-圍囹");


                                        //判斷圍囹之垂直深度，斜率零為負，反之為正
                                        if ((i % 2 == 0))
                                        {
                                            beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 - beam_B / 2).ToString());//2000為支撐階數深度，表1中
                                        }
                                        else
                                        {
                                            beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 + beam_B / 2).ToString());
                                        }

                                    }

                                    if (dex.beamLevel[lev].Item2 == 2)
                                    {
                                        c = null;
                                        if (i == innerwall_points.Count<XYZ>())
                                        {
                                            try
                                            {
                                                int.Parse(dex.beamLevel[lev].Item1.ToString());
                                            }
                                            catch { c = Line.CreateBound(sidewall_points[i], sidewall_points[0]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }

                                        }
                                        else
                                        {
                                            try
                                            {
                                                int.Parse(dex.beamLevel[lev].Item1.ToString());
                                            }
                                            catch { c = Line.CreateBound(sidewall_points[i], sidewall_points[i + 1]).CreateOffset((beam_H) / 304.8, new XYZ(0, 0, -1)); }
                                        }
                                        if (c != null)
                                        {
                                            FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                            beam.LookupParameter("斷面旋轉").SetValueString("90");
                                            StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                                            StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                                            beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.beamLevel[lev].Item1).ToString() + "-圍囹");


                                            //判斷圍囹之垂直深度，斜率零為負，反之為正
                                            if ((i % 2 == 0))
                                            {
                                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 - beam_B / 2).ToString());//2000為支撐階數深度，表1中
                                            }
                                            else
                                            {
                                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 + beam_B / 2).ToString());
                                            }

                                        }
                                    }
                                }
                            }
                        }
                    }
                    
                    trans2.Commit();

                    trans2.Start("建置回撐支撐");
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
                    //string erroemessage = "";
                    for (int lev = 0; lev != dex.supLevel.Count(); lev++)
                    {
                        try { int.Parse(dex.supLevel[lev].Item1.ToString());}
                        catch
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
                                                    frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-dou_frame_H - dex.sideWall[0].Item2 * 1000).ToString());
                                                    frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-dou_frame_H - dex.sideWall[0].Item2 * 1000).ToString());

                                                    //取消接合
                                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                                    //若為雙向支撐，則鏡射支撐             
                                                    if (dex.supLevel[lev].Item3 == 2)
                                                    {
                                                        ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                                    }
                                                }
                                                catch (Exception e) { TaskDialog.Show("error", e.StackTrace + e.Message + e.Source); break; };
                                            }
                                            break;
                                        }
                                    }
                                    //Y向支撐|||||||
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
                                                    frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * -1000 + frame_B).ToString());//2000為支撐階數深度，表1中
                                                    frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-Fframe_B - (columns_B - Fframe_B) / 2).ToString());
                                                    frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-dou_frame_H - dex.sideWall[0].Item2 * 1000).ToString());
                                                    frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-dou_frame_H - dex.sideWall[0].Item2 * 1000).ToString());

                                                    //取消接合
                                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                                    //若為雙向支撐，則鏡射支撐
                                                    if (dex.supLevel[lev].Item3 == 2)
                                                    {
                                                        ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                                    }
                                                }
                                                catch (Exception e) { TaskDialog.Show("error", e.StackTrace + e.Message + e.Source); break; };
                                            }
                                            break;
                                        }
                                    }

                                }
                            }
                        }
                    }

                    trans2.Commit();


                }

                catch (Exception e) { TaskDialog.Show("error", e.StackTrace + e.Message + e.Source); break; }
            }
            TaskDialog.Show("done", "回撐建置完畢");
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
