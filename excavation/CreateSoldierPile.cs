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

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    
    class CreateSoldierPile : IExternalEventHandler
    {
        //讀取深開挖資料xlsx檔案
        public IList<string> files_path
        {
            get;
            set;
        }

        //從使用者介面讀取x,y值（平移量）
        public IList<double> xy_shift
        {
            get;
            set;
        }
        public string type
        {
            get;
            set;
        }
        public OpenFileDialog openFileDialog
        {
            get;
            set;
        }
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
                dWGImportOptions.Placement = ImportPlacement.Origin;
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
                ExReader sheet = new ExReader();
                try
                {
                    sheet.SetData(file_path, 1);
                    sheet.PassWallData();
                    sheet.PassColumnData();
                    sheet.CloseEx();

                    if (type == "型鋼樁")
                    {
                        sheet.SetData(file_path, 3);
                        sheet.PassSoldierPile();
                        sheet.CloseEx();
                    }else if(type == "鋼軌樁")
                    {
                        sheet.SetData(file_path, 6);
                        sheet.PassRailSoldierPile();
                        sheet.CloseEx();
                    }

                    sheet.SetData(file_path, 7);
                    sheet.PassTimberLagging();
                    sheet.CloseEx();
                }
                catch (Exception e) { sheet.CloseEx(); TaskDialog.Show("Error", e.Message + e.StackTrace); }

                SubTransaction subtran = new SubTransaction(doc);
                tran.Start("create soldier pile");

                //開挖各階之深度輸入
                List<double> height = new List<double>();
                foreach (var data in sheet.excaLevel)
                    height.Add(data.Item2 * -1);

                //偏移量
                double xshift = xy_shift[0];
                double yshift = xy_shift[1];

                //建立開挖階數            
                Level[] levlist = new Level[height.Count()];
                try
                {
                    for (int i = 0; i != height.Count(); i++)
                    {
                        levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                        levlist[i].Name = String.Format("斷面{0}-開挖階數" + (i + 1).ToString(), sheet.section);
                    }

                }
                catch { }
                Level wall_level = Level.Create(doc, sheet.wall_high * 1000 * -1 / 304.8);
                try 
                { 
                    wall_level.Name = String.Format("斷面{0}-擋土壁深度", sheet.section); 
                } 
                catch { }

                //須回到原點
                List<Line> lines = new List<Line>();
                if (dwg_file_name.Contains(".dwg"))
                {
                    IList<String> target_section = new List<String>();
                    target_section.Add(sheet.section);
                    IList<XYZ> allXYZs = new List<XYZ>();
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

                                //建置線段//
                                if (Clockdirection == true)//若為順時針需要倒轉
                                {
                                    allXYZs = allXYZs.Reverse().ToList();
                                }
                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);
                                    lines.Add(line);
                                }
                            }
                        }
                        catch (Exception e) { }

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
                            }
                        }
                        catch (Exception e) { }
                    }
                }
                else
                {
                    //依照xlsx內給定的xy座標建立開挖範圍
                    //讀取座標建立points
                    for (int i = 0; i != sheet.excaRange.Count() - 1; i++)
                    {
                        XYZ point1 = new XYZ(sheet.excaRange[i].Item1 - xshift, sheet.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;
                        XYZ point2 = new XYZ(sheet.excaRange[i + 1].Item1 - xshift, sheet.excaRange[i + 1].Item2 - yshift, 0) * 1000 / 304.8;

                        lines.Add(Line.CreateBound(point1, point2));
                    }
                }

                Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().First();

                double B = 0;
                double h = 0;
                double t1 = 0;
                double t2 = 0;
                double distance = 0;
                FamilySymbol fs = null;

                if (type == "型鋼樁")
                {
                    B = sheet.soldier_pile[0].Item1;
                    h = sheet.soldier_pile[0].Item2;
                    t1 = sheet.soldier_pile[0].Item3;
                    t2 = sheet.soldier_pile[0].Item4;
                    distance = B * 2 / 304.8;

                    // get h beam element
                    fs = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "型鋼樁").ToList().First();
                    fs.Activate();
                    fs.LookupParameter("高度").SetValueString((sheet.wall_high * 1000).ToString());
                    fs.LookupParameter("H").SetValueString(h.ToString());
                    fs.LookupParameter("B").SetValueString(B.ToString());
                    fs.LookupParameter("t1").SetValueString(t1.ToString());
                    fs.LookupParameter("t2").SetValueString(t2.ToString());
                }
                else if (type == "鋼軌樁")
                {
                    h = sheet.rail_soldier_pile[0].Item1; // A
                    B = sheet.rail_soldier_pile[0].Item2;
                    double C = sheet.rail_soldier_pile[0].Item3;
                    t1 = sheet.rail_soldier_pile[0].Item4; // G
                    double D = sheet.rail_soldier_pile[0].Item5;
                    double E = sheet.rail_soldier_pile[0].Item6;
                    t2 = h - D - E; // F

                    distance = B * 4 / 304.8;

                    // get h beam element
                    fs = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "鋼軌樁").ToList().First();
                    fs.Activate();
                    fs.LookupParameter("高度").SetValueString((sheet.wall_high * 1000).ToString());

                    fs.LookupParameter("A").SetValueString(h.ToString());
                    fs.LookupParameter("B").SetValueString(B.ToString());
                    fs.LookupParameter("C").SetValueString(C.ToString());
                    fs.LookupParameter("t").SetValueString(t1.ToString());
                    fs.LookupParameter("D").SetValueString(D.ToString());
                    fs.LookupParameter("E").SetValueString(E.ToString());
                    // fs.LookupParameter("F").SetValueString(t2.ToString());
                }

                // get timber lagging element
                double wall_W = sheet.timber_lagging[0].Item2 * 10; //cm in Excel
                WallType wallType = null;
                ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
                //檢查擋土壁
                if (walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == wall_W.ToString() + "mm").ToList().Count != 0)
                    wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == wall_W.ToString() + "mm").ToList().First();
                //建立擋土壁新類型
                if (wallType == null)
                {
                    wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁").ToList().First();
                    WallType new_wallFamSym = wallType.Duplicate("連續壁-" + wall_W + "mm") as WallType;
                    CompoundStructure ly = new_wallFamSym.GetCompoundStructure();
                    ly.SetLayerWidth(0, wall_W / 304.8);

                    ElementId material_id = Material.Create(doc, "timber lagging");
                    Material material = doc.GetElement(material_id) as Material;
                    double timber_lagging_width = sheet.timber_lagging[0].Item1 / 10 / 304.8;
                    FillPattern surface_pattern = new FillPattern("timber", FillPatternTarget.Drafting, FillPatternHostOrientation.ToView, 0, timber_lagging_width);
                    FillPatternElement fillPatternElement = FillPatternElement.Create(doc, surface_pattern);
                    material.SurfaceForegroundPatternId = fillPatternElement.Id;
                    material.SurfaceForegroundPatternColor = new Color(255, 255, 255);
                    material.Color = new Color(120, 120, 120);
                    ly.SetMaterialId(0, material.Id);

                    new_wallFamSym.SetCompoundStructure(ly);
                    wallType = new_wallFamSym;
                }

                for (int i = 0; i < lines.Count(); i++)
                {
                    subtran.Start();
                    Line edge_line = lines[i];
                    // doc.Create.NewDetailCurve(doc.ActiveView, edge_line);

                    double slope = edge_line.Direction.Y / edge_line.Direction.X;
                    int pile_num = (int)(edge_line.Length / distance) + 1;

                    List<ElementId> pre_nor_rotate_list = new List<ElementId>();
                    List<XYZ> pre_nor_rotate_point_list = new List<XYZ>();
                    List<ElementId> pre_rotate_list = new List<ElementId>();
                    List<XYZ> pre_rotate_point_list = new List<XYZ>();

                    XYZ nomal_vector = XYZ.BasisZ.CrossProduct(edge_line.Direction);
                    
                    // create corner h beam in each edge_line
                    XYZ corner_point = edge_line.GetEndPoint(1) + B / 2 / 304.8 * edge_line.Direction - (h / 2 / 304.8) * nomal_vector;
                    FamilyInstance instance3 = doc.Create.NewFamilyInstance(corner_point, fs, level, StructuralType.NonStructural);
                    ElementTransformUtils.RotateElement(doc, instance3.Id, Line.CreateBound(corner_point, (corner_point + XYZ.BasisZ)), Math.Atan(slope));

                    Line wood_line;
                    Curve c;
                    FamilyInstance instance2;
                    for (int j = 0; j < pile_num; j++)
                    {
                        XYZ point = edge_line.GetEndPoint(0) + (j * distance + B / 2 / 304.8) * edge_line.Direction - (h / 2 / 304.8) * nomal_vector;
                        XYZ next_point = point + distance * edge_line.Direction;
                        XYZ wall_point = point + (t1 / 2 / 304.8) * edge_line.Direction;
                        XYZ wall_next_point = point - (t1 / 2 / 304.8) * edge_line.Direction + distance * edge_line.Direction;

                        // if the next h beam collide with the corner h beam, create the last h beam
                        if (corner_point.DotProduct(edge_line.Direction) - next_point.DotProduct(edge_line.Direction) < B / 304.8)
                        {
                            // create h beam
                            instance2 = doc.Create.NewFamilyInstance(point, fs, level, StructuralType.NonStructural);
                            ElementTransformUtils.RotateElement(doc, instance2.Id, Line.CreateBound(point, (point + XYZ.BasisZ)), Math.Atan(slope));

                            // create timber lagging in all level
                            wood_line = Line.CreateBound(wall_point, corner_point - (t1 / 2 / 304.8) * edge_line.Direction);
                            c = wood_line.CreateOffset(((h - wall_W) / 2 - t2) / 304.8, new XYZ(0, 0, -1));
                            for (int k = 0; k < levlist.Length - 1; k++)
                            {
                                Wall wall = Wall.Create(doc, c, wallType.Id, level.Id, sheet.wall_high * 1000 / 304.8, 0, false, false);
                                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levlist[k].Id);
                                wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).Set(levlist[k + 1].Id);
                            }
                            break;
                        }

                        // create h beam
                        instance2 = doc.Create.NewFamilyInstance(point, fs, level, StructuralType.NonStructural);
                        ElementTransformUtils.RotateElement(doc, instance2.Id, Line.CreateBound(point, (point + XYZ.BasisZ)), Math.Atan(slope));
                        
                        // if the end point of the wall exceed the corner h beam, change end point to the last h beam
                        if (corner_point.DotProduct(edge_line.Direction) - wall_next_point.DotProduct(edge_line.Direction) < B / 2 / 304.8)
                        {
                            wall_next_point = corner_point - (t1 / 2 / 304.8) * edge_line.Direction;
                        }

                        // create timber lagging in all level
                        wood_line = Line.CreateBound(wall_point, wall_next_point);
                        c = wood_line.CreateOffset(((h - wall_W) / 2 - t2) / 304.8, new XYZ(0, 0, -1));
                        for (int k = 0; k < levlist.Length - 1; k++)
                        {
                            Wall wall = Wall.Create(doc, c, wallType.Id, level.Id, sheet.wall_high * 1000 / 304.8, 0, false, false);
                            wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levlist[k].Id);
                            wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).Set(levlist[k + 1].Id);
                        }

                    }

                    subtran.Commit();
                }
                tran.Commit();
            }
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

    }
}
