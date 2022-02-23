using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using ExReaderConsole;
using Excel = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB.Analysis;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class CreateMonitor : IExternalEventHandler
    {
        public IList<string> files_path
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
            UIDocument uiDoc = app.ActiveUIDocument;

            Transaction tran = new Transaction(doc);
            int i = 0;

            List<double> coordinate_n = new List<double>();
            List<double> coordinate_e = new List<double>();
            List<string> equipment_id = new List<string>(); 
            List<double> settlement = new List<double>();
            List<double> alert = new List<double>();
            ProjectPosition projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);

            View view = doc.ActiveView;

            TextNoteOptions textNoteOptions = new TextNoteOptions();
            TextNoteType tType = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();
            textNoteOptions.HorizontalAlignment = HorizontalTextAlignment.Left;
            textNoteOptions.TypeId = tType.Id;


            //List<string> monitor_id_list = new List<string>{ "SB0022", "SB0017", "SB0018", "SB0032", "SB0033", "SB0034", "SB0035", "SB0036", "SB0037", "SB0038" };

            //讀取資料
            ExReader dex = new ExReader();
            dex.SetData(@"\\Mac\Home\Desktop\excavation\20210615給台大範例\CQ852_20210415.XLS", 1);
            //dex.SetData(@"\\Mac\Home\Desktop\excavation\20211209_制式化表單-監測儀器data.xlsx", 4);
            try
            {
                coordinate_n = dex.PassMonitorDouble("N座標", 3).ToList();
                coordinate_e = dex.PassMonitorDouble("E座標", 3).ToList();
                equipment_id = dex.PassMonitortString("儀器編號", 3).ToList();
                settlement = dex.PassMonitorDouble("沉陷量", 3).ToList();
                alert = dex.PassFirstRowDouble("高行動值", 3).ToList();

                dex.CloseEx();
            }
            catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }
            //List<string> monitor_id_list = equipment_id;

            List<XYZ> monitor_point_list = new List<XYZ>();
            for (i = 0; i != coordinate_n.Count(); i++)
            {
                coordinate_e[i] = coordinate_e[i] * 1000 / 304.8 - projectPosition.EastWest;
                coordinate_n[i] = coordinate_n[i] * 1000 / 304.8 - projectPosition.NorthSouth;
                XYZ monitor_point = new XYZ(coordinate_e[i], coordinate_n[i], 0);
                monitor_point_list.Add(monitor_point);
            }
            
            
            tran.Start("start");
            if (files_path[0].Contains("dwg"))
            {
                DWGImportOptions dWGImportOptions = new DWGImportOptions();
                dWGImportOptions.ColorMode = ImportColorMode.Preserved;
                dWGImportOptions.Placement = ImportPlacement.Shared;
                dWGImportOptions.Unit = ImportUnit.Meter;
                LinkLoadResult linkLoadResult = new LinkLoadResult();
                ImportInstance dwg = ImportInstance.Create(doc, view, files_path[0], dWGImportOptions, out linkLoadResult);

                dwg.Pinned = false;

                //取得CAD
                Transform project_transform = dwg.GetTotalTransform();
                GeometryElement geometryElement = dwg.get_Geometry(new Options());
                GeometryElement geoLines = (geometryElement.First() as GeometryInstance).SymbolGeometry;

                string target_section = "監測儀器";

                Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, 0));
                SketchPlane sp = SketchPlane.Create(doc, plane);
                
                i = 0;
                foreach (var v in geoLines)
                {
                    try
                    {
                        PolyLine pline = v as PolyLine;
                        GraphicsStyle check = doc.GetElement(pline.GraphicsStyleId) as GraphicsStyle;
                        if (check.GraphicsStyleCategory.Name.Contains(target_section) && check.GraphicsStyleCategory.Name.Contains(type))
                        {
                            IList<XYZ> edge_panel = new List<XYZ>();
                            XYZ middle_point = new XYZ(0, 0, 0);
                            int count = 0;
                            foreach (XYZ p in pline.GetCoordinates())
                            {
                                count++;
                                middle_point += project_transform.OfPoint(p);
                            }
                            i++;
                            middle_point = middle_point / count;

                            FamilySymbol familySymbol;
                            try
                            {
                                familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name.Contains(type)).First();
                            }
                            catch
                            {
                                familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "監測_連續壁沉陷觀測點").First();

                            }

                            FamilyInstance monitor = doc.Create.NewFamilyInstance(middle_point, familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            int nearest_index = FindNearest(middle_point, monitor_point_list);
                            string monitor_id = equipment_id[nearest_index];

                            /*if (monitor_id_list.Contains(monitor_id))
                            {
                                TextNote note3 = TextNote.Create(doc, view.Id, monitor_point_list[nearest_index], equipment_id[nearest_index], textNoteOptions);

                            }*/
                        }
                    }
                    catch { }
                }
                //doc.Delete(dwg.Id);

            } else if (files_path[0].Contains("xls") || files_path[0].Contains("XLS"))
            {
                string detectObject = "監測_土中傾度管";
                //List<string> elementIds = new List<string>();
                //List<double> settlement_list = new List<double>();
                //List<XYZ> visual_monitor_point_list = new List<XYZ>();

                //List<Tuple<string, double, XYZ>> visual_data = new List<Tuple<string, double, XYZ>>();

                Wall target_wall = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>()
                            .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == "FM50").First();

                IList<Wall> wall_list = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>().ToList();

                LocationCurve target_lc = target_wall.Location as LocationCurve;
                XYZ target_start = target_lc.Curve.GetEndPoint(0);
                XYZ target_end = target_lc.Curve.GetEndPoint(1);
                XYZ target_mid = (target_start + target_end) / 2;

                double target_slope = (target_start.Y - target_end.Y) / (target_start.X - target_end.X);
                double target_normal = -1 / target_slope;
                double target_normal_b = target_mid.Y - target_normal * target_mid.X;

                double except_distance = 50;
                double previous_distance = 10000;
                Wall source_wall = target_wall;
                foreach (Wall w in wall_list)
                {
                    LocationCurve source_lc = w.Location as LocationCurve;
                    XYZ source_start = source_lc.Curve.GetEndPoint(0);
                    XYZ source_end = source_lc.Curve.GetEndPoint(1);
                    XYZ source_mid = (source_start + source_end) / 2;

                    double source_slope = (source_start.Y - source_end.Y) / (source_start.X - source_end.X);

                    double mid_distance = target_mid.DistanceTo(source_mid);
                    double project_distance = Math.Abs(target_normal * source_mid.X + target_normal_b - source_mid.Y);
                    if (project_distance < previous_distance && mid_distance > except_distance && Math.Abs(target_slope - source_slope) < 0.1)
                    {
                        previous_distance = project_distance;
                        source_wall = w;
                    }
                }
                TaskDialog.Show("1", source_wall.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString());

                double width = 10;
                LocationCurve select_source_lc = source_wall.Location as LocationCurve;
                XYZ select_source_start = select_source_lc.Curve.GetEndPoint(0);
                XYZ select_source_end = select_source_lc.Curve.GetEndPoint(1);
                XYZ select_source_mid = (select_source_start + select_source_end) / 2;

                ViewFamilyType viewFamilyType = (from v in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                                                 where v.ViewFamily == ViewFamily.ThreeDimensional
                                                 select v).First();

                View3D view3d = View3D.CreateIsometric(doc, viewFamilyType.Id);
                view3d.DisplayStyle = DisplayStyle.ShadingWithEdges;
                ViewDisplayModel viewDisplayModel = view3d.GetViewDisplayModel();

                view3d.Name = "section box";

                XYZ center = (target_mid + select_source_mid) / 2;
                double center_distance = target_mid.DistanceTo(center);


                BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                boundingBoxXYZ.Max = new XYZ(center.X + center_distance + 30 * width, target_mid.Y + 1 * width, 0);
                boundingBoxXYZ.Min = new XYZ(center.X - center_distance - 30 * width, select_source_mid.Y - 1 * width, center.Z);
                view3d.SetSectionBox(boundingBoxXYZ);

                XYZ axis = (boundingBoxXYZ.Max + boundingBoxXYZ.Min) / 2;
                Line axis_z = Line.CreateBound(axis, axis + new XYZ(0, 0, 10));


                Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, 0));
                SketchPlane sp = SketchPlane.Create(doc, plane);

                Arc arc3 = Arc.Create(new XYZ(axis.X, axis.Y, 0), 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                ModelCurve mc3 = doc.Create.NewModelCurve(arc3, sp);



                FamilySymbol familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == detectObject).First();
                Transform rotate = Transform.CreateRotationAtPoint(new XYZ(0, 0, 10), Math.Atan(target_normal), new XYZ(axis.X, axis.Y, 0));
                Transform reverse_rotate = Transform.CreateRotationAtPoint(new XYZ(0, 0, 10), -Math.Atan(target_normal), new XYZ(axis.X, axis.Y, 0));
                //Transform rotate = Transform.CreateRotationAtPoint(new XYZ(0, 0, 1), Math.PI, new XYZ(axis.X, axis.Y, 0));
                XYZ point = new XYZ(center.X + center_distance + 30 * width, target_mid.Y + 1 * width, 0);

                Arc arc = Arc.Create(point, 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                ModelCurve mc = doc.Create.NewModelCurve(arc, sp);

                Arc arc2 = Arc.Create(point, 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                ModelCurve mc2 = doc.Create.NewModelCurve(arc2, sp);
                

                boundingBoxXYZ.Transform = boundingBoxXYZ.Transform.Multiply(rotate);
                view3d.SetSectionBox(boundingBoxXYZ);


                List<string> visual_id = new List<string>();
                List<double> visual_settlement = new List<double>();
                List<XYZ> visual_coord = new List<XYZ>();
                i = 0;

                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets.Add();

                int j = 1;
                worksheet.Cells[j, 1] = "儀器編號";
                worksheet.Cells[j, 2] = "沉陷量";
                for (i = 0; i != coordinate_n.Count(); i++)
                {
                    string monitor_id = equipment_id[i];
                    
                    XYZ point_xyz = monitor_point_list[i];
                    XYZ rotate_point_xyz = reverse_rotate.OfPoint(point_xyz);

                    if (rotate_point_xyz.X <= boundingBoxXYZ.Max.X && rotate_point_xyz.Y <= boundingBoxXYZ.Max.Y &&
                        rotate_point_xyz.X >= boundingBoxXYZ.Min.X && rotate_point_xyz.Y >= boundingBoxXYZ.Min.Y)
                    {
                        //monitor_point_list[i] = reverse_rotate.OfPoint(monitor_point_list[i]);
                        FamilyInstance monitor3 = doc.Create.NewFamilyInstance(monitor_point_list[i], familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        monitor3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(equipment_id[i] + "/" + settlement[i]);
                        
                        visual_id.Add(equipment_id[i]);
                        visual_settlement.Add(settlement[i]);
                        visual_coord.Add(monitor_point_list[i]);

                        j++;
                        worksheet.Cells[j, 1] = equipment_id[i];
                        worksheet.Cells[j, 2] = settlement[i];
                    }


                    /*
                    TextNote note = TextNote.Create(doc, view.Id, monitor_point_list[i], "*" + equipment_id[i], textNoteOptions);
                    note.TextNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(0.1);
                    note.TextNoteType.get_Parameter(BuiltInParameter.LINE_COLOR).Set(170000);                       

                    var data = Tuple.Create(equipment_id[i], settlement[i], monitor_point_list[i]);
                    visual_data.Add(data);
                    */
                }

                application.ActiveWorkbook.SaveAs(@"\\Mac\Home\Desktop\excavation\test.xls", Excel.XlFileFormat.xlWorkbookNormal);
                workbook.Close();
                application.Quit();

                TaskDialog.Show("1", "export done");
                
                XYZ x1 = new XYZ(center.X + center_distance + 30 * width, target_mid.Y + 1 * width, 0);
                XYZ x2 = new XYZ(center.X + center_distance + 30 * width, select_source_mid.Y - 1 * width, 0);
                XYZ x3 = new XYZ(center.X - center_distance - 30 * width, select_source_mid.Y - 1 * width, 0);
                XYZ x4 = new XYZ(center.X - center_distance - 30 * width, target_mid.Y + 1 * width, 0);
                
                CurveArray profile = new CurveArray();
                profile.Append(Line.CreateBound(x1, x2));
                profile.Append(Line.CreateBound(x2, x3));
                profile.Append(Line.CreateBound(x3, x4));
                profile.Append(Line.CreateBound(x4, x1));

                FloorType floorType = new FilteredElementCollector(doc).OfClass(typeof(FloorType)).Cast<FloorType>().Where(x => x.Name == "通用 150mm").First();
                Element floor = doc.Create.NewFloor(profile, floorType, Level.Create(doc, 0), false, XYZ.BasisZ);
                ElementTransformUtils.RotateElement(doc, floor.Id, axis_z, Math.Atan(target_normal));

                tran.Commit();
                tran.Start("visualize");

                double previous_z = -100;
                PlanarFace target_face = null;
                GeometryElement geometryElement = floor.get_Geometry(new Options());

                foreach(GeometryObject geometryObject in geometryElement)
                {
                    Solid solid = geometryObject as Solid;
                            
                    if (solid != null)
                    {
                        FaceArray faceArray = solid.Faces;
                        foreach (PlanarFace planarFace in faceArray)
                        {
                            if (previous_z <= planarFace.Origin.Z)
                            {
                                previous_z = planarFace.Origin.Z;
                                target_face = planarFace;
                            }
                        }
                    }
                }

                AnalysisDisplayStyle analysisDisplayStyle = null;
                ICollection<Element> collection = new FilteredElementCollector(doc).OfClass(typeof(AnalysisDisplayStyle)).ToElements();

                var displayStyle = from element in collection
                                    where element.Name == "Display Style 2"
                                    select element;

                if (displayStyle.Count() == 0)
                {
                    AnalysisDisplayColoredSurfaceSettings coloredSurfaceSettings = new AnalysisDisplayColoredSurfaceSettings();
                    coloredSurfaceSettings.ShowGridLines = true;

                    AnalysisDisplayColorSettings colorSettings = new AnalysisDisplayColorSettings();
                    colorSettings.ColorSettingsType = AnalysisDisplayStyleColorSettingsType.SolidColorRanges;

                    Color deepRed = new Color(166, 0, 0);
                    Color red = new Color(255, 44, 01);
                    Color green = new Color(0, 253, 0);
                    Color lightGreen = new Color(128, 255, 12);
                    Color orange = new Color(255, 205, 0);
                    Color purple = new Color(200, 0, 200);
                    Color white = new Color(255, 255, 255);


                    colorSettings.MaxColor = red;
                    colorSettings.SetIntermediateColors(new List<AnalysisDisplayColorEntry>
                                                {
                                                    new AnalysisDisplayColorEntry(red, -20),
                                                    new AnalysisDisplayColorEntry(orange, -10),
                                                    new AnalysisDisplayColorEntry(lightGreen, 0),
                                                    new AnalysisDisplayColorEntry(orange, 10),

                                                });

                    colorSettings.MinColor = purple;

                    AnalysisDisplayLegendSettings legendSettings = new AnalysisDisplayLegendSettings
                    {
                        NumberOfSteps = 10,
                        Rounding = 0.05,
                        ShowDataDescription = false,
                        ShowLegend = true

                    };

                    analysisDisplayStyle = AnalysisDisplayStyle.CreateAnalysisDisplayStyle(doc, "Display Style 2", coloredSurfaceSettings, colorSettings, legendSettings);
                }
                else
                {
                    analysisDisplayStyle = displayStyle.Cast<AnalysisDisplayStyle>().ElementAt<AnalysisDisplayStyle>(0);
                }
                doc.ActiveView.AnalysisDisplayStyleId = analysisDisplayStyle.Id;

                SpatialFieldManager sfm = SpatialFieldManager.GetSpatialFieldManager(doc.ActiveView);

                try
                {
                    sfm.Clear();
                }
                catch { }

                if (sfm == null)
                {
                    sfm = SpatialFieldManager.CreateSpatialFieldManager(doc.ActiveView, 1);
                }

                IList<UV> uvPts = new List<UV>();
                BoundingBoxUV bb = target_face.GetBoundingBox();
                UV min = bb.Min;
                UV max = bb.Max;

                for (i = 0; i != visual_id.Count(); i++)
                {
                    visual_coord[i] = reverse_rotate.OfPoint(visual_coord[i]);
                    uvPts.Add(new UV(visual_coord[i].X, visual_coord[i].Y) - new UV(x2.X, x2.Y));
                        
                }

                UV faceCenter = new UV((max.U + min.U) / 2, (max.V + min.V) / 2);
                Transform computeDerivatives = target_face.ComputeDerivatives(faceCenter);
                XYZ faceCenterNormal = computeDerivatives.BasisZ;
                XYZ faceCenterNormalMultiplied = faceCenterNormal.Normalize().Multiply(1);
                Transform transform = Transform.CreateTranslation(faceCenterNormalMultiplied);

                FieldDomainPointsByUV pnts = new FieldDomainPointsByUV(uvPts);
                IList<ValueAtPoint> valList = new List<ValueAtPoint>();

                for (i = 0; i != visual_id.Count(); i++)
                {
                    valList.Add(new ValueAtPoint(new List<double> { visual_settlement[i] }));
                }

                FieldValues vals = new FieldValues(valList);

                AnalysisResultSchema resultSchema = new AnalysisResultSchema("鋼筋應力", "Description");
                resultSchema.SetUnits(new List<string> { "kg/cm^2" }, new List<double> { 1 });

                sfm.SetMeasurementNames(new List<string> { "資料 1" });

                int idx = sfm.AddSpatialFieldPrimitive(target_face, transform);
                int schemaIndex = sfm.RegisterResult(resultSchema);
                sfm.UpdateSpatialFieldPrimitive(idx, pnts, vals, schemaIndex);
                
                /* plot settlement
                visual_data.Sort((x, y) => x.Item3.X.CompareTo(y.Item3.X));

                DataTable dt = new DataTable();
                dt.Columns.Add("distance");
                dt.Columns.Add("settlement");

                List<double> distance_cumsum_list = new List<double>{ 0 };
                for (i = 0; i < visual_data.Count() - 1; i++)
                {
                    double distance_delta = visual_data[i].Item3.DistanceTo(visual_data[i + 1].Item3);
                    double cumsum = distance_cumsum_list[i] + distance_delta;

                    dt.Rows.Add(Convert.ToInt32(cumsum), visual_data[i].Item2);
                    distance_cumsum_list.Add(cumsum);
                }
                Form3 form3 = new Form3(dt);
                form3.Show();*/
            }

            TaskDialog.Show("1", "done");
            tran.Commit();
            


            /*
            tran.Start("start2");
            i = 0;
            foreach (ElementId elementId in elementIds)
            {

                Element element = doc.GetElement(elementId);
                TaskDialog.Show("1", elementId.ToString());

                element.LookupParameter("儀器編號").Set(equipment_id[i]);


                //monitor.LookupParameter("N").SetValueString(coordinate_n[i].ToString());
                //monitor.LookupParameter("E").SetValueString(coordinate_e[i].ToString());

                i++;
                if (i >= 3)
                {
                    break;
                }
            }
            tran.Commit();*/

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


        public Tuple<XYZ, XYZ> LinearRegression(List<XYZ> monitor_point_list)
        {
            int n = monitor_point_list.Count();
            double x = 0;
            double y = 0;
            double sum_x = 0;
            double sum_y = 0;
            double sum_x2 = 0;
            double sum_xy = 0;

            for (int i = 0; i < n; i++)
            {
                x = monitor_point_list[i].X;
                y = monitor_point_list[i].Y;
                sum_x += x;
                sum_y += y;
                sum_x2 += x * x;
                sum_xy += x * y;
            }

            double a = (sum_y * sum_x2 - sum_x * sum_xy) / (n * sum_x2 - sum_x * sum_x);
            double b = (n * sum_xy - sum_x * sum_y) / (n * sum_x2 - sum_x * sum_x);

            double x_min = monitor_point_list.OrderBy(i => i.X).FirstOrDefault().X;
            double y_min = a * x_min + b;
            double x_max = monitor_point_list.OrderBy(i => i.X).LastOrDefault().X;
            double y_max = a * x_max + b;

            XYZ start_point = new XYZ(x_min, y_min, 0);
            XYZ end_point = new XYZ(x_max, y_max, 0);
            Tuple<XYZ, XYZ> start_end = Tuple.Create(start_point, end_point);

            return start_end;
        }

        static void ExportExcel()
        {

            Excel.Application myexcelApplication = new Excel.Application();
            if (myexcelApplication != null)
            {
                Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                myexcelWorksheet.Cells[1, 1] = "Value 1";
                myexcelWorksheet.Cells[2, 1] = "Value 2";
                myexcelWorksheet.Cells[3, 1] = "Value 3";

                myexcelApplication.ActiveWorkbook.SaveAs(@"C:\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);

                myexcelWorkbook.Close();
                myexcelApplication.Quit();
            }
        }
    }
}
