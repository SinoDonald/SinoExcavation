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

    class CreateUnit : IExternalEventHandler
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
        public string unit_dwg_file_name
        {
            get;
            set;
        }
        public string unit_id_position_excel_path
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
            Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Document doc = uidoc.Document;
            ProjectPosition projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);

            Transaction trans = new Transaction(doc);
            try
            {
                //讀取資料
                ExReader dex = new ExReader();
                dex.SetData(files_path[0], 1);
                try
                {
                    dex.PassWallData();
                    //dex.SetData(@"\\Mac\Home\Desktop\excavation\test.xlsx", 1);
                    dex.SetData(unit_id_position_excel_path, 1);
                    dex.PassUnitText();
                    dex.CloseEx();
                }
                catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }
                                
                //插入CAD
                trans.Start("CAD");
                //string unit_dwg_file_name = @"\\Mac\Home\Desktop\excavation\LG05連續壁單元分割圖.dwg";
                //string diaphragm_dwg_file_name = @"\\Mac\Home\Desktop\excavation\LG05測試\LG05-20181129.dwg";

                Autodesk.Revit.DB.View view = doc.ActiveView;
                DWGImportOptions dWGImportOptions = new DWGImportOptions();
                dWGImportOptions.Placement = ImportPlacement.Shared;
                dWGImportOptions.ColorMode = ImportColorMode.Preserved;
                dWGImportOptions.Unit = ImportUnit.Meter;

                LinkLoadResult linkLoadResult = new LinkLoadResult();
                ImportInstance unit_dwg = ImportInstance.Create(doc, view, unit_dwg_file_name, dWGImportOptions, out linkLoadResult);
                //ImportInstance diaphragm_dwg = ImportInstance.Create(doc, view, diaphragm_dwg_file_name, dWGImportOptions, out linkLoadResult);

                unit_dwg.Pinned = false;
                //diaphragm_dwg.Pinned = false;

                ElementTransformUtils.MoveElement(doc, unit_dwg.Id, new XYZ(xy_shift[0], xy_shift[1], 0));
                //ElementTransformUtils.MoveElement(doc, diaphragm_dwg.Id, new XYZ(xy_shift[0], xy_shift[1], 0));

                TextNoteOptions textNoteOptions = new TextNoteOptions();
                TextNoteType tType = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();
                textNoteOptions.HorizontalAlignment = HorizontalTextAlignment.Left;
                textNoteOptions.TypeId = tType.Id;

                
                Level wall_level = Level.Create(doc, dex.wall_high * 1000 * -1 / 304.8);

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

                List<XYZ> unit_text_point_list = new List<XYZ>();
                List<string> unit_text_list = new List<string>();
                foreach (Tuple<string, double, double> unit_text in dex.unit_text)
                {
                    XYZ unit_text_point = new XYZ(unit_text.Item2 * 1000 / 304.8 - projectPosition.EastWest, unit_text.Item3 * 1000 / 304.8 - projectPosition.NorthSouth, 0);
                    Transform rotate = Transform.CreateRotationAtPoint(new XYZ(0, 0, 10), 48.6 / 180 * Math.PI, new XYZ(0, 0, 0));
                    //unit_text_point = rotate.OfPoint(unit_text_point);

                    unit_text_point_list.Add(unit_text_point);
                    unit_text_list.Add(unit_text.Item1);
                    TextNote note = TextNote.Create(doc, view.Id, unit_text_point, unit_text.Item1, textNoteOptions);
                }

                trans.Commit();

                // get unit point
                trans.Start("get point");
                Transform project_transform = unit_dwg.GetTotalTransform();
                GeometryElement geometryElement = unit_dwg.get_Geometry(new Options());
                GeometryElement geoLines = (geometryElement.First() as GeometryInstance).SymbolGeometry;
                string target_section = "端版";
                List<XYZ> unit_point = new List<XYZ>();

                foreach (var v in geoLines)
                {
                    try
                    {
                        PolyLine pline = v as PolyLine;
                        GraphicsStyle check = doc.GetElement(pline.GraphicsStyleId) as GraphicsStyle;
                        if (check.GraphicsStyleCategory.Name.Contains(target_section))
                        {
                            IList<XYZ> edge_panel = new List<XYZ>();
                            XYZ middle_point = new XYZ(0, 0, 0);
                            int count = 0;

                            foreach (XYZ p in pline.GetCoordinates())
                            {
                                count++;
                                middle_point += project_transform.OfPoint(p);
                            }

                            middle_point = middle_point / count;
                            unit_point.Add(middle_point);
                        }
                    }
                    catch { }
                }

                // get wall coordinate from CreateWallCADmode
                IList<Element> walls = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Walls).ToElements();
                int k = 0;
                foreach (Element e in walls)
                {
                    Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, 0));
                    SketchPlane sp = SketchPlane.Create(doc, plane);

                    Wall wall = e as Wall;
                    LocationCurve lc = wall.Location as LocationCurve;

                    XYZ start_point = lc.Curve.GetEndPoint(0);
                    start_point = new XYZ(start_point.X, start_point.Y, 0);
                    //Arc arc = Arc.Create(project_transform.OfPoint(start_point), 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                    //ModelCurve mc = doc.Create.NewModelCurve(arc, sp);

                    XYZ end_point = lc.Curve.GetEndPoint(1);
                    end_point = new XYZ(end_point.X, end_point.Y, 0);
                    //Arc arc2 = Arc.Create(project_transform.OfPoint(end_point), 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                    //ModelCurve mc2 = doc.Create.NewModelCurve(arc2, sp);

                    List<XYZ> wall_point_list = new List<XYZ>();
                    wall_point_list.Add(start_point);
                    wall_point_list.Add(end_point);
                    foreach (XYZ point in unit_point)
                    {
                        double wall_slope = Math.Atan((start_point.X - end_point.X) / (start_point.Y - end_point.Y));
                        double unit_slope = Math.Atan((start_point.X - point.X) / (start_point.Y - point.Y));
                        
                        // check unit point and slope whether between wall
                        if ( 
                             (((start_point.X - point.X) * (end_point.X - point.X) < 0 ) || ((start_point.Y - point.Y) * (end_point.Y - point.Y) < 0)) &&
                             Math.Abs(wall_slope - unit_slope) < 0.01 
                             )
                        {
                            wall_point_list.Add(point);
                        }
                    }

                    // sort by x direction
                    wall_point_list.Sort((x, y) => x.X.CompareTo(y.X));
                    //TextNote note = TextNote.Create(doc, view.Id, wall_point_list[0], k.ToString() + "start", textNoteOptions);
                    //TextNote note2 = TextNote.Create(doc, view.Id, wall_point_list.Last(), k.ToString() + "end", textNoteOptions);
                    k++; 
                    // check sorting result
                    //int k = 0;

                    /*
                    foreach (XYZ point in wall_point_list)
                    {
                        //Arc arc3 = Arc.Create(project_transform.OfPoint(point), 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                        //ModelCurve mc3 = doc.Create.NewModelCurve(arc3, sp);
                        //TextNote note = TextNote.Create(doc, view.Id, point, k.ToString(), textNoteOptions);
                        //k++;
                    }*/

                    // create unit wall
                    doc.Delete(wall.Id);
                    // 7000 mm
                    double wall_length_limit = 7000 / 304.8;
                    double wall_length = wall_point_list[0].DistanceTo(wall_point_list.Last());
                    if (wall_point_list.Count() > 2)
                    {
                        for (int j = 0 ; j < wall_point_list.Count() - 1 ; j++)
                        {
                            Line line = Line.CreateBound(wall_point_list[j], wall_point_list[j + 1]);
                            XYZ middle_point = (wall_point_list[j] + wall_point_list[j + 1]) / 2;
                            Wall w = Wall.Create(doc, line, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, false);
                            int nearest_index = FindNearest(middle_point, unit_text_point_list);
                            TextNote note3 = TextNote.Create(doc, view.Id, unit_text_point_list[nearest_index], unit_text_list[nearest_index], textNoteOptions);
                            w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(unit_text_list[nearest_index]);

                        }

                    }
                    else if(wall_length < wall_length_limit)
                    {
                        for (int j = 0; j < wall_point_list.Count() - 1; j++)
                        {
                            Line line = Line.CreateBound(wall_point_list[j], wall_point_list[j + 1]);
                            Wall w = Wall.Create(doc, line, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, false);
                        }
                    }
                    
                    wall_point_list.Clear();
                }

                

                //doc.Delete(unit_dwg.Id);
                //doc.Delete(diaphragm_dwg.Id);
                trans.Commit();
            }
            catch (Exception e) { TaskDialog.Show("error test!!", e.Message + e.StackTrace); }
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }

        public int FindNearest(XYZ element_point, List<XYZ> point_list)
        {
            double distance = element_point.DistanceTo(point_list[0]);
            int nearest_index = 0;
            int i = 0;

            foreach (XYZ point in point_list)
            {
                if (element_point.DistanceTo(point) < distance)
                {
                    distance = element_point.DistanceTo(point);
                    nearest_index = i;
                }
                i++;
            }

            return nearest_index;

        }

    }
}
