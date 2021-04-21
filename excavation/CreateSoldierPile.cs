using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;

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
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;

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

                Transaction tran = new Transaction(doc);
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
                XYZ[] points = new XYZ[sheet.excaRange.Count()];

                //依照xlsx內給定的xy座標建立開挖範圍
                //讀取座標建立points
                for (int i = 0; i != sheet.excaRange.Count(); i++)
                {
                    points[i] = new XYZ(sheet.excaRange[i].Item1 - xshift, sheet.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;
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
                    new_wallFamSym.SetCompoundStructure(ly);
                    wallType = new_wallFamSym;
                }

                for (int i = 0; i < points.Count() - 1; i++)
                {
                    subtran.Start();
                    Line edge_line = Line.CreateBound(points[i], points[i + 1]);
                    // doc.Create.NewDetailCurve(doc.ActiveView, edge_line);

                    double slope = edge_line.Direction.Y / edge_line.Direction.X;
                    int pile_num = (int)(edge_line.Length / distance) + 1;

                    List<ElementId> pre_nor_rotate_list = new List<ElementId>();
                    List<XYZ> pre_nor_rotate_point_list = new List<XYZ>();
                    List<ElementId> pre_rotate_list = new List<ElementId>();
                    List<XYZ> pre_rotate_point_list = new List<XYZ>();

                    XYZ nomal_vector = XYZ.BasisZ.CrossProduct(edge_line.Direction);

                    // create corner h beam in each edge_line
                    XYZ corner_point = points[i + 1] + B / 2 / 304.8 * edge_line.Direction - (h / 2 / 304.8) * nomal_vector;
                    FamilyInstance instance3 = doc.Create.NewFamilyInstance(corner_point, fs, level, StructuralType.NonStructural);
                    ElementTransformUtils.RotateElement(doc, instance3.Id, Line.CreateBound(corner_point, (corner_point + XYZ.BasisZ)), Math.Atan(slope));

                    Line wood_line;
                    Curve c;
                    FamilyInstance instance2;
                    for (int j = 0; j < pile_num; j++)
                    {
                        XYZ point = points[i] + (j * distance + B / 2 / 304.8) * edge_line.Direction - (h / 2 / 304.8) * nomal_vector;
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

    }
}
