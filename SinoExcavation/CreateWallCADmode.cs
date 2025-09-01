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

namespace SinoExcavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class CreateWallCADmode : IExternalEventHandler
    {
        public delegate void ReturnDegree(double degree);
        public event ReturnDegree ReturnDegreeCallback;
        

        public OpenFileDialog openFileDialog
        {
            get;
            set;
        }
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
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            
            Transaction trans = new Transaction(doc);
            try
            {
                
                trans.Start("CAD");
                //插入CAD

                Autodesk.Revit.DB.View view = doc.ActiveView;
                DWGImportOptions dWGImportOptions = new DWGImportOptions();
                dWGImportOptions.Placement = ImportPlacement.Shared;
                //dWGImportOptions.Placement = ImportPlacement.Origin;

                dWGImportOptions.ColorMode = ImportColorMode.Preserved;
                LinkLoadResult linkLoadResult = new LinkLoadResult();
                ImportInstance toz = ImportInstance.Create(doc, view, openFileDialog.FileName, dWGImportOptions, out linkLoadResult);
                
                toz.Pinned = false;
                ElementTransformUtils.MoveElement(doc, toz.Id, new XYZ(xy_shift[0], xy_shift[1], 0));
                trans.Commit();

                //取得CAD
                Transform project_transform = toz.GetTotalTransform();
                GeometryElement geometryElement = toz.get_Geometry(new Options());
                GeometryElement geoLines = (geometryElement.First() as GeometryInstance).SymbolGeometry;

                IList<double> allThetas = new List<double>();
                List<Wall> all_walls = new List<Wall>();
                List<Room> all_rooms = new List<Room>();


                foreach (string file_path in files_path)
                {
                    //讀取資料
                    ExReader dex = new ExReader();
                    dex.SetData(file_path, 1);
                    try
                    {
                        dex.PassWallData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

                    trans.Start("交易開始");
                    //開挖各階之深度輸入
                    List<double> height = new List<double>();
                    foreach (var data in dex.excaLevel)
                        height.Add(data.Item2 * -1);

                    //偏移量
                    double xshift = xy_shift[0];
                    double yshift = xy_shift[1];

                    //建立開挖階數
                    Level[] levlist = new Level[height.Count()];
                    for (int i = 0; i != height.Count(); i++)
                    {
                        try
                        {
                            levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                            levlist[i].Name = String.Format("斷面{0}-開挖階數" + (i + 1).ToString(), dex.section);

                        }
                        catch { }
                    }
                    Level wall_level = Level.Create(doc, dex.wall_high * 1000 * -1 / 304.8);
                    try { 
                        wall_level.Name = String.Format("斷面{0}-擋土壁深度", dex.section);
                    }
                    catch { }

                    //建立連續壁
                    IList<Curve> inner_wall_curves = new List<Curve>();
                    double wall_W = dex.wall_width * 1000; //連續壁厚度

                    WallType wallType = null;
                    ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
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
                    trans.Commit();
                    Transaction transaction = new Transaction(doc);

                    transaction.Start("create wall");

                    IList<String> target_section = new List<String>();

                    target_section.Add(dex.section);

                    IList<XYZ> allXYZs = new List<XYZ>();
                    List<Wall> inner_wall = new List<Wall>();
                    List<Wall> inner_wall_for_arc = new List<Wall>();
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
                                IList<Curve> wall_profileloops = new List<Curve>();
                                if (Clockdirection == true)//若為順時針需要倒轉
                                {
                                    allXYZs = allXYZs.Reverse().ToList();
                                }
                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);

                                    allThetas.Add(Math.Atan((line.Direction.Y) / (line.Direction.X)) * 180 / Math.PI);

                                    wall_profileloops.Add(line);
                                }

                                //建立牆
                                for (int i = 0; i < allXYZs.Count<XYZ>() - 1; i++)
                                {
                                    Curve c = wall_profileloops[i].CreateOffset(wall_W / 2 / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線

                                    Wall w = Wall.Create(doc, c, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, false);

                                    w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + "連續壁");
                                    inner_wall.Add(w);
                                    all_walls.Add(w);

                                }
                            }
                        }
                        catch (Exception) { }
                    }

                    IList<XYZ> poly_wall_points = new List<XYZ>();

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
                                IList<Curve> wall_profileloops = new List<Curve>();

                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);

                                    allThetas.Add(Math.Atan((line.Direction.Y) / (line.Direction.X)) * 180 / Math.PI);

                                    wall_profileloops.Add(line);
                                }


                                //建立牆
                                for (int i = 0; i < allXYZs.Count<XYZ>() - 1; i++)
                                {
                                    Curve c = wall_profileloops[i].CreateOffset(wall_W / 2 / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線

                                    Wall w = Wall.Create(doc, c, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, false);

                                    w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("連續壁");
                                    inner_wall_for_arc.Add(w);

                                }
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
                        catch (Exception) { }

                    }

                    //蒐集所有連續壁點座標
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
                    
                    //建置房間
                    
                    Room room = doc.Create.NewRoom(wall_level, new UV(wall_points.Select(x => x.X).ToList().Average(), wall_points.Select(y => y.Y).ToList().Average()));
                    room.Name = "斷面" + dex.section;
                    room.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section);
                    all_rooms.Add(room);
                    
                    transaction.Commit();



                }
                //取角度眾數

                /*
                int modeValue = allThetas
                                  .GroupBy(x => (int)x)
                                  .OrderByDescending(x => x.Count()).ThenBy(x => x.Key)
                                  .Select(x => (int)x.Key)
                                  .FirstOrDefault();//1

                double rotate_angle = allThetas.Where(x => (int)x == modeValue).Average();

                ReturnDegreeCallback(rotate_angle);// 執行委派的方法到這個事件

                IList<ElementId> elementIds = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>().ToList().Select(x =>x.Id).ToList();

                elementIds.Concat(all_rooms.Select(x => x.Id).ToList());

                Transaction transaction_RC = new Transaction(doc);

                transaction_RC.Start("rotateWallRoom");

                ElementTransformUtils.RotateElements(doc, elementIds, Line.CreateBound(new XYZ(0, 0, 0), new XYZ(0, 0, 1)), (-rotate_angle) / 180 * Math.PI);

                transaction_RC.Commit();
                */
                
                TaskDialog.Show("omg", "done");
            }
            catch (Exception e){ TaskDialog.Show("error test!!",e.Message + e.StackTrace); }
        }

        public string GetName()
        {
            return "Event handler is working now!!";
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
                z = (points[j].X - points[i].X) * (points[k].Y - points[j].Y );
                z -= (points[j].Y  - points[i].Y) * (points[k].X - points[j].X);
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


    }
}
