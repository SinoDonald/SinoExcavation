using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;

namespace SinoExcavation_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Circle_excavation : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        public IList<double> xy_shift
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            foreach (string file_path in files_path)
            {
                Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);
                Autodesk.Revit.DB.Document doc = uidoc.Document;
                Transaction trans = new Transaction(doc);
                trans.Start("交易開始");

                //讀取匯入檔案資料
                ExReader dex = new ExReader();
                dex.SetData(file_path, 1);
                try
                {
                    dex.PassCircle();
                    dex.CloseEx();
                }
                catch { dex.CloseEx(); }

                //偏移量
                double shift_x = xy_shift[0];
                double shift_y = xy_shift[0];

                //開挖各階之深度輸入
                List<double> height = new List<double>();
                foreach (var data in dex.excaLevel)
                    height.Add(data.Item2 * -1);

                //建立開挖階數            
                Level[] levlist = new Level[height.Count()];
                for (int i = 0; i != height.Count(); i++)
                {
                    levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                    levlist[i].Name = dex.section + "-" + "開挖階數" + (i + 1).ToString();
                }
                Level wall_level = Level.Create(doc, dex.wall_high * 1000 * -1 / 304.8);
                wall_level.Name = "擋土壁深度";

                //建立樓板回築樓層
                Level[] re_levlist = new Level[dex.circle_floor.Count()];
                for (int i = 0; i != dex.circle_floor.Count(); i++)
                {
                    re_levlist[i] = Level.Create(doc, (dex.circle_floor[i].Item2 * -1 + dex.circle_floor[i].Item3 / 2) * 1000 / 304.8);
                    re_levlist[i].Name = dex.section + "-" + (dex.circle_floor[i].Item1).ToString();
                }
                Level[] all_level_list = new Level[levlist.Count() + re_levlist.Count()];
                all_level_list = levlist.Concat(re_levlist).ToArray();
                //訂定開挖範圍
                IList<CurveLoop> profileloops = new List<CurveLoop>();

                //須回到原點
                XYZ[] points = new XYZ[dex.excaRange.Count()];

                for (int i = 0; i != dex.excaRange.Count(); i++)
                    points[i] = new XYZ(dex.excaRange[i].Item1 - shift_x, dex.excaRange[i].Item2 - shift_y, 0) * 1000 / 304.8;
                TaskDialog.Show("1", "5.1");
                dex.Diameter = 20;
                //建立環形輪廓
                IList<Curve> wall_profileloops = new List<Curve>();
                for (int i = 0; i < 2; i++)
                {
                    TaskDialog.Show("index", new XYZ(0 - shift_x, 0 - shift_y, 0).ToString() +"/"+ (dex.Diameter * 1000 / 304.8 / 2).ToString() +"/"+ (Math.PI * i).ToString() +"/"+ (Math.PI * (i + 1)).ToString()+"/"+ XYZ.BasisX+"/"+XYZ.BasisY);
                    wall_profileloops.Add(Arc.Create(new XYZ(0 - shift_x, 0 - shift_y, 0), dex.Diameter * 1000 / 304.8 / 2, Math.PI * i, Math.PI * (i + 1), XYZ.BasisX, XYZ.BasisY));
                }
                TaskDialog.Show("1", "5.2");

                CurveLoop profileloop = CurveLoop.Create(new List<Curve>(wall_profileloops));
                profileloops.Add(profileloop);
                Level levdeep = null;
                TaskDialog.Show("1", "6");

                //建立開挖深度
                ICollection<Level> level_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                TaskDialog.Show("1", "6.1");

                foreach (Level lev in level_familyinstance)
                {
                    TaskDialog.Show("1", lev.Name +"/"+ levlist[levlist.Count() - 1].Name);

                    if (lev.Name == levlist[levlist.Count() - 1].Name)
                    {
                        levdeep = lev;
                    }
                }
                TaskDialog.Show("1", "7");

                //建立連續壁
                IList<Curve> inner_wall_curves = new List<Curve>();
                double wall_W = dex.wall_width * 1000; //連續壁厚度
                WallType wallType = null;
                List<Wall> inner_wall = new List<Wall>();
                ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
                TaskDialog.Show("1", "8");

                //檢查擋土壁
                if (walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (dex.wall_width * 1000).ToString() + "mm").ToList().Count != 0)
                    wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (dex.wall_width * 1000).ToString() + "mm").ToList().First();
                //建立擋土壁新類型
                if (wallType == null)
                {
                    wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁").ToList().First();
                    WallType new_wallFamSym = wallType.Duplicate("連續壁-" + dex.wall_width * 1000 + "mm") as WallType;
                    CompoundStructure ly = new_wallFamSym.GetCompoundStructure();
                    ly.SetLayerWidth(0, dex.wall_width * 1000 / 304.8);
                    new_wallFamSym.SetCompoundStructure(ly);
                    wallType = new_wallFamSym;
                }
                //建立
                for (int i = 0; i < wall_profileloops.Count<Curve>(); i++)
                {
                    Curve c = wall_profileloops[i].CreateOffset(wall_W / 2 / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線
                    Wall w = Wall.Create(doc, c, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, false);
                    w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + "連續壁");
                    inner_wall.Add(w);
                }

                trans.Commit();


                //取得連續壁內座標點
                XYZ[] innerwall_points = new XYZ[points.Count<XYZ>()];
                XYZ[] for_check_innerwall_points = new XYZ[points.Count<XYZ>()];//用來計算支撐中間樁之擺放界線
                for (int i = 0; i < (inner_wall.Count<Wall>()); i++)
                {
                    //inner
                    XYZ wall_curve_point = (inner_wall[i].Location as LocationCurve).Curve.Tessellate()[0];
                    wall_curve_point = new XYZ(wall_curve_point.X, wall_curve_point.Y, 0);
                    innerwall_points[i] = points[i] - (points[i] - wall_curve_point) * 2;
                    for_check_innerwall_points[i] = points[i] - (points[i] - wall_curve_point) * 1.4;//用來計算支撐中間樁之擺放界線
                }
                innerwall_points[points.Count<XYZ>() - 1] = innerwall_points[0];
                for_check_innerwall_points[points.Count<XYZ>() - 1] = for_check_innerwall_points[0];




                //建立樓板回築範圍
                
                trans.Start("開始建置");

                //建立底板類型及實作元件
                
                ICollection<FloorType> family = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().ToList();
                FloorType floor_type = family.Where(x => x.Name == "通用 300mm").First();
                int base_control = 0;
                foreach (Tuple<string, double, double, double, double> base_tuple in dex.circle_floor)
                {
                    CurveArray profileloops_array = new CurveArray();
                    for (int i = 0; i < 2; i++)
                    {
                        profileloops_array.Append(Arc.Create(new XYZ(0 - shift_x, 0 - shift_y, 0), base_tuple.Item4 * 1000 / 304.8 / 2, Math.PI * i, Math.PI * (i + 1), XYZ.BasisX, XYZ.BasisY));
                    }
                    FloorType newFamSym = null;
                    try { newFamSym = floor_type.Duplicate(base_tuple.Item1) as FloorType; } catch { newFamSym = family.Where(x => x.Name == base_tuple.Item1).First(); }
                    double floor_width = base_tuple.Item3 * 1000 / 304.8;// set the width to a new value:
                    double floor_deep = base_tuple.Item2 * -1 * 1000 / 304.8;
                    double floor_offset = (floor_deep + floor_width / 2) * 304.8;
                    newFamSym.GetCompoundStructure().GetLayers()[0].Width = floor_width;
                    CompoundStructure ly = newFamSym.GetCompoundStructure();
                    ly.SetLayerWidth(0, floor_width);
                    newFamSym.SetCompoundStructure(ly);
                    //Floor floor = doc.Create.NewFloor(profileloops_array, newFamSym as FloorType, re_levlist[base_control], false);

                    // 培文改
                    IList<Curve> curves = new List<Curve>();
                    for (int i = 0; i < 2; i++)
                    {
                        curves.Append(Arc.Create(new XYZ(0 - shift_x, 0 - shift_y, 0), base_tuple.Item4 * 1000 / 304.8 / 2, Math.PI * i, Math.PI * (i + 1), XYZ.BasisX, XYZ.BasisY));
                    }
                    CurveLoop curveLoop = CurveLoop.Create(curves);
                    IList<CurveLoop> curveLoops = new List<CurveLoop>();
                    curveLoops.Add(curveLoop);
                    Element floor = Floor.Create(doc, curveLoops, newFamSym.Id, Level.Create(doc, 0).Id);

                    floor.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + base_tuple.Item1);
                    floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).SetValueString("0");
                    floor.get_Parameter(BuiltInParameter.DOOR_NUMBER).Set(base_tuple.Item5.ToString());
                    base_control += 1;
                }
                trans.Commit();
                

                
                
                for (int lev = 0; lev != dex.supLevel.Count(); lev++)
                {
                    //建立環形支撐
                    ICollection<FamilySymbol> beam_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                    foreach (FamilySymbol beam_type in beam_familyinstance)
                    {
                        if (beam_type.Name == "H150x150x7x10")
                        {

                            beam_type.Activate();
                            double beam_H = double.Parse(beam_type.LookupParameter("H").AsValueString());
                            double beam_B = double.Parse(beam_type.LookupParameter("B").AsValueString());
                            for (int i = 0; i < wall_profileloops.Count<Curve>(); i++)
                            {
                                Curve c = wall_profileloops[i].CreateOffset(wall_W * 2 / 304.8, XYZ.BasisZ * -1);

                                trans.Start("開始建置");
                                FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levlist[0], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                
                                StructuralFramingUtils.DisallowJoinAtEnd(beam, 0);
                                StructuralFramingUtils.DisallowJoinAtEnd(beam, 1);
                                
                                beam.LookupParameter("斷面旋轉").SetValueString("90");
                                beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + (dex.supLevel[lev].Item1).ToString() + "-環形鋼支保");


                                //修正環形鋼支保深度

                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000).ToString());
                                trans.Commit();

                            }
                        }
                    }
                }
                
            }
        }

        //判斷點與點是否為相同點，因為xyz類型無法做equals
        public static bool calculator(XYZ checkPoint, List<XYZ> ranoutpoints)
        {
            bool a = false;
            foreach (XYZ ranoutpoint in ranoutpoints)
            {
                if (ranoutpoint.X - checkPoint.X < 1 && ranoutpoint.X - checkPoint.X > -1)
                {
                    if (ranoutpoint.Y - checkPoint.Y < 1 && ranoutpoint.Y - checkPoint.Y > -1)
                    {
                        if (ranoutpoint.Z - checkPoint.Z < 1 && ranoutpoint.Z - checkPoint.Z > -1)
                        {
                            a = true;
                        }
                    }
                }
            }
            return a;
        }

        //判斷點是否位於開挖範圍內
        public static bool IsInPolygon(XYZ checkPoint, XYZ[] polygonPoints)
        {
            bool inside = false;
            int pointCount = polygonPoints.Count<XYZ>();
            XYZ p1, p2;
            for (int i = 0, j = pointCount - 1; i < pointCount; j = i, i++)
            {
                p1 = polygonPoints[i];
                p2 = polygonPoints[j];
                if (checkPoint.Y < p2.Y)
                {
                    if (p1.Y <= checkPoint.Y)
                    {
                        if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) > (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                        {

                            inside = (!inside);
                        }
                    }
                }
                else if (checkPoint.Y < p1.Y)
                {

                    if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) < (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                    {
                        inside = (!inside);
                    }
                }
            }
            return inside;
        }

        //計算支撐與連續壁交點，true為計算X向，false為計算Y向
        public static XYZ[] intersection(XYZ checkPoint, XYZ[] polygonPoints, double[] slope, double[] bias, bool x_or_y)
        {
            List<XYZ> intersection_points = new List<XYZ>();
            List<double> x = new List<double>();
            List<double> y = new List<double>();
            if (x_or_y == true)
            {
                for (int i = 0; i < slope.Count<double>();i++)
                {
                    if (slope[i] !=0||Math.Abs(slope[i])>0.00001)
                    {
                        if(slope[i] == 20172018)
                        {
                            intersection_points.Add(new XYZ(bias[i], checkPoint.Y, 0));
                        }
                        else
                        {
                            intersection_points.Add(new XYZ((checkPoint.Y - bias[i]) / slope[i], checkPoint.Y, 0));
                        }
                    }
                }
                
                foreach (XYZ point in intersection_points)
                {
                    if (IsInPolygon(point, polygonPoints))
                    {
                        x.Add(point.X);
                        y.Add(point.Y);
                    }                    
                }
            }

            else//若為False則計算Y向
            {
                for (int i = 0; i < slope.Count<double>(); i++)
                {
                    if (slope[i] != 20172018)
                    {
                        if (slope[i] == 0|| Math.Abs(slope[i]) < 0.00001)
                        {
                            
                            intersection_points.Add(new XYZ(checkPoint.X, bias[i], 0));
                        }
                        else
                        {
                            intersection_points.Add(new XYZ(checkPoint.X, slope[i]*checkPoint.X+bias[i],0));
                        }
                    }
                }

                foreach (XYZ point in intersection_points)
                {
                    if (IsInPolygon(point, polygonPoints))
                    {

                        x.Add(point.X);
                        y.Add(point.Y);
                    }
                }
            }

            XYZ[] intersection = new XYZ[2];
            try
            {
                intersection[0] = new XYZ(x.Min(), y.Min(), 0);
                intersection[1] = new XYZ(x.Max(), y.Max(), 0);
                return intersection;
            }
            catch { return intersection; }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}

