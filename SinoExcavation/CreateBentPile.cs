using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace SinoExcavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CreateBentPile : IExternalEventHandler
    {

        //讀取深開挖資料xlsx檔案
        public IList<string> files_path { get; set; }

        //從使用者介面讀取x,y值（平移量）
        public IList<double> xy_shift { get; set; } 

        public OpenFileDialog openFileDialog { get; set; }

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            Transaction tran = new Transaction(doc);

            string dwg_file_name = "";
            Transform project_transform = null;
            GeometryElement geoLines = null;

            try
            {
                dwg_file_name = openFileDialog.FileName;
            }
            catch { }

            if (dwg_file_name.Contains(".dwg"))
            {

                tran.Start("CAD");
                //插入CAD
                Autodesk.Revit.DB.View view = doc.ActiveView;
                DWGImportOptions dWGImportOptions = new DWGImportOptions();
                dWGImportOptions.ColorMode = ImportColorMode.Preserved;
                //dWGImportOptions.Placement = ImportPlacement.Origin;
                dWGImportOptions.Placement = ImportPlacement.Shared;
                LinkLoadResult linkLoadResult = new LinkLoadResult();
                ImportInstance toz = ImportInstance.Create(doc, view, dwg_file_name, dWGImportOptions, out linkLoadResult);
                toz.Pinned = false;
                ElementTransformUtils.MoveElement(doc, toz.Id, new XYZ(xy_shift[0], xy_shift[1], 0));

                tran.Commit();

                //取得CAD
                project_transform = toz.GetTotalTransform();
                GeometryElement geometryElement = toz.get_Geometry(new Options());
                geoLines = (geometryElement.First() as GeometryInstance).SymbolGeometry;
            }

            foreach (string file_path in files_path)
            {
                
                //讀取資料
                ExReader pileBent = new ExReader();
                pileBent.SetData(file_path, 1);
                try
                {
                    pileBent.PassWallData();
                    pileBent.PassPileBentData();
                    pileBent.CloseEx();
                }
                catch (Exception e) { pileBent.CloseEx(); TaskDialog.Show("Error", e.Message + e.StackTrace); }

                
                                
                tran.Start("放排樁");

                //開挖各階之深度輸入
                List<double> height = new List<double>();
                foreach (var data in pileBent.excaLevel)
                    height.Add(data.Item2 * -1);

                //偏移量
                double xshift = xy_shift[0];
                double yshift = xy_shift[1];

                //建立開挖階數            
                try
                {
                    Level[] levlist = new Level[height.Count()];
                    for (int i = 0; i != height.Count(); i++)
                    {
                        levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                        levlist[i].Name = String.Format("斷面{0}-開挖階數" + (i + 1).ToString(), pileBent.section);
                    }

                }
                catch { }
                Level wall_level = Level.Create(doc, pileBent.wall_high * 1000 * -1 / 304.8);
                try { wall_level.Name = String.Format("斷面{0}-擋土壁深度", pileBent.section); } catch { }

                // fake wall type
                WallType wallType = null;
                ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
                //檢查擋土壁
                if (walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (pileBent.pileBent_diameter * 1000).ToString() + "mm").ToList().Count != 0)
                    wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (pileBent.pileBent_diameter * 1000).ToString() + "mm").ToList().First();
                //建立擋土壁新類型
                if (wallType == null)
                {
                    wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁").ToList().First();
                    WallType new_wallFamSym = wallType.Duplicate("連續壁-" + pileBent.pileBent_diameter * 1000 + "mm") as WallType;
                    CompoundStructure ly = new_wallFamSym.GetCompoundStructure();
                    ly.SetLayerWidth(0, pileBent.pileBent_diameter * 1000 / 304.8);
                    new_wallFamSym.SetCompoundStructure(ly);
                    wallType = new_wallFamSym;
                }


                //temp
                List<Wall> inner_wall = new List<Wall>();
                List<Wall> inner_wall_for_arc = new List<Wall>();

                //須回到原點
                List<Line> lines = new List<Line>();
                if (dwg_file_name.Contains(".dwg"))
                {

                    IList<String> target_section = new List<String>();
                    target_section.Add(pileBent.section);
                    IList<XYZ> allXYZs = new List<XYZ>();

                    IList<XYZ> poly_wall_points = new List<XYZ>();

                    foreach (var v in geoLines)
                    {
                        allXYZs.Clear();
                        try//處理polyline
                        {
                            PolyLine pline = v as PolyLine;
                            GraphicsStyle check = doc.GetElement(pline.GraphicsStyleId) as GraphicsStyle;

                            //檢查是否為所要圖層
                            if (target_section.Contains(check.GraphicsStyleCategory.Name))
                            {

                                //撈取所有點位
                                foreach (XYZ p in pline.GetCoordinates())
                                {
                                    allXYZs.Add(project_transform.OfPoint(p - new XYZ(xshift, yshift, 0)));
                                }

                                bool Clockdirection = ClockwiseDirection(allXYZs);

                                //建置線段
                                if (Clockdirection == true)//若為順時針需要倒轉
                                {
                                    allXYZs = allXYZs.Reverse().ToList();
                                }
                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);
                                    lines.Add(line);;
                                }


                                // Create invisible walls
                                for (int i = 0; i < allXYZs.Count<XYZ>() - 1; i++)
                                {                                    
                                    //Curve c = lines[i].CreateOffset(pileBent.pileBent_diameter * 1000 / 2 / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線

                                    Wall w = Wall.Create(doc, lines[i], wallType.Id, wall_level.Id, pileBent.wall_high * 1000 / 304.8, 0, false, false);
                                                                      

                                    w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(pileBent.section + "-" + "連續壁");
                                    inner_wall.Add(w);
                                }

                                

                                foreach (Wall w in inner_wall)
                                {
                                    foreach (XYZ xyz in (w.Location as LocationCurve).Curve.Tessellate())
                                    {
                                        if (poly_wall_points.Select(x => x.IsAlmostEqualTo(xyz)).Contains(true))//過濾掉重複項
                                        {
                                        }
                                        else
                                        {
                                            poly_wall_points.Add(xyz);
                                        }
                                    }
                                }


                            }
                        }
                        catch (Exception) { }

                    }


                    //處理Arc順逆時針問題
                    foreach (var v in geoLines)
                    {
                        allXYZs.Clear();

                        

                        try
                        {                           
                            //處理ARC
                            Arc pline = v as Arc;


                            //檢查是否為所要圖層
                            if (target_section.Contains((doc.GetElement(pline.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory.Name))
                            {
                                //撈取所有點位
                                foreach (XYZ p in pline.Tessellate())
                                {
                                    allXYZs.Add(project_transform.OfPoint(p - new XYZ(xshift, yshift, 0)));
                                }

                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);
                                    lines.Add(line);
                                }


                                // Create invisible walls
                                for (int i = 0; i < allXYZs.Count<XYZ>() - 1; i++)
                                {
                                    Wall w = Wall.Create(doc, lines[i], wallType.Id, wall_level.Id, pileBent.wall_high * 1000 / 304.8, 0, false, false);

                                    w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("連續壁");
                                    inner_wall_for_arc.Add(w);

                                }

                                IList<XYZ> arc_wall_points = new List<XYZ>();

                                foreach (Wall w in inner_wall_for_arc)
                                {
                                    foreach (XYZ xyz in (w.Location as LocationCurve).Curve.Tessellate())
                                    {
                                        if (arc_wall_points.Select(x => x.IsAlmostEqualTo(xyz)).Contains(true))//過濾掉重複項
                                        {
                                        }
                                        else
                                        {
                                            arc_wall_points.Add(xyz);
                                        }
                                    }
                                }
                                XYZ arc_moveP = poly_wall_points.OrderBy(x => x.DistanceTo(arc_wall_points[0])).ToList()[0] - arc_wall_points[0];

                                ElementTransformUtils.MoveElements(doc, inner_wall_for_arc.Select(x => x.Id).ToList(), arc_moveP);
                            }                           


                        }
                        catch {

                        }
                       
                        
                    }

                }
                else
                {
                    //依照xlsx內給定的xy座標建立開挖範圍
                    //讀取座標建立points
                    for (int i = 0; i != pileBent.excaRange.Count() - 1; i++)
                    {
                        XYZ point1 = new XYZ(pileBent.excaRange[i].Item1 - xshift, pileBent.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;
                        XYZ point2 = new XYZ(pileBent.excaRange[i + 1].Item1 - xshift, pileBent.excaRange[i + 1].Item2 - yshift, 0) * 1000 / 304.8;

                        lines.Add(Line.CreateBound(point1, point2));
                    }
                }


                Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().First();

                
                //選出排樁族群
                FamilySymbol eleFamSym = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "排樁").ToList().First();
                //eleFamSym.Activate();

                //定義排樁實例參數
                double pileBent_D = pileBent.pileBent_diameter * 1000; //排樁直徑
                double pileBent_S = pileBent.pileBent_space * 1000 / 304.8; //排樁間距
                double pileBent_H = pileBent.wall_high * 1000; //排樁深度

                
                //建置排樁
                foreach (var line in lines)
                {

                    int pile_num = (int)(line.Length / pileBent_S);

                    if (line.Length - pile_num * pileBent_S > pileBent_D * 0.5 / 304.8) 
                        pile_num++;

                    IList<XYZ> column_points = new List<XYZ>();
                    
                    for (int j = 0; j < pile_num; j++)
                    {
                        XYZ column_point = line.GetEndPoint(0) + j * pileBent_S * line.Direction;// + pileBent_S * line.Direction;
                        column_points.Add(column_point);
                    }

                    for (int j = 0; j < pile_num; j++)
                    {
                        FamilyInstance pileBent_instance = doc.Create.NewFamilyInstance(column_points[j], eleFamSym, wall_level, Autodesk.Revit.DB.Structure.StructuralType.Column);
                        pileBent_instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).SetValueString("0");
                        pileBent_instance.LookupParameter("直徑").SetValueString((pileBent_D).ToString());//給定排樁直徑
                        pileBent_instance.LookupParameter("深度").SetValueString((pileBent_H).ToString());//給定排樁深度

                        // Don't be room boundary
                        pileBent_instance.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).Set(0);
                    }
                   
                }


                // create room
                IList<XYZ> wall_points = new List<XYZ>();

                foreach (Wall w in inner_wall.Concat(inner_wall_for_arc))
                {
                    foreach (XYZ xyz in (w.Location as LocationCurve).Curve.Tessellate())
                    {
                        if (wall_points.Select(x => x.IsAlmostEqualTo(xyz)).Contains(true))//過濾掉重複項
                        {
                        }
                        else
                        {
                            wall_points.Add(xyz);
                        }
                    }
                }

                Room room = doc.Create.NewRoom(wall_level, new UV(wall_points.Select(x => x.X).ToList().Average(), wall_points.Select(y => y.Y).ToList().Average()));
                room.Name = "斷面" + pileBent.section;
                room.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(pileBent.section);

                tran.Commit();
            }


        }

        private bool ClockwiseDirection(IList<XYZ> points)
        {
            //計算多邊形邊界線順時鐘/逆時鐘
            bool clock = true;
            int i, j, k;
            int count = 0;
            double z;
            int n = points.Count;
            for (i = 0; i < n; i++)
            {
                j = (i + 1) % n;
                k = (i + 2) % n;
                z = (points[j].X - points[i].X) * (points[k].Y - points[j].Y);
                z -= (points[j].Y - points[i].Y) * (points[k].X - points[j].X);
                if (z < 0)
                {
                    count--;
                }
                else if (z > 0)
                {
                    count++;
                }
            }
            if (count > 0)
            {
                clock = false;
            }

            return clock;

        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
