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
using System.IO;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class VisualizeMonitor : IExternalEventHandler
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
        public string excel_path
        {
            get;
            set;
        }
        public string monitor_path
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
            List<string> unit_id_list = new List<string>();

            List<string> equipment_id = new List<string>();
            List<string> filtered_equipment_id = new List<string>();
            List<int> equipment_id_index = new List<int> { 0 };

            List<double> settlement = new List<double>();
            List<double> sid_displacement = new List<double>();
            List<double> depth = new List<double>();
            List<int> sid_index = new List<int> { 0 };

            List<string> ssi_id = new List<string>();
            List<double> ssi_displacement = new List<double>();
            List<string> vg_id = new List<string>();
            List<double> axial_force = new List<double>();

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
            //string excel_path = @"\\Mac\Home\Desktop\excavation\forAPI制式化表單-單元名稱、監測儀器data.xlsx";

            try
            {
                dex.SetData(excel_path, 1);
                unit_id_list = dex.PassRowString("單元名稱編號", 0, 1).ToList();
                int unit_amount = unit_id_list.Count();




                foreach (string unit_id in unit_id_list.ToArray())
                {
                    List<string> equipment_id_list = dex.PassColumntString(unit_id, 1, 0).ToList();

                    foreach (string equipment in equipment_id_list)
                    {
                        filtered_equipment_id.Add(equipment);
                    }
                    equipment_id_index.Add(equipment_id_list.Count());
                }
                dex.CloseEx();

                for (i = 2; 2 + unit_amount != i; i++)
                {
                    dex.SetData(excel_path, i);
                    List<double> depth_list = dex.PassColumnDouble("觀測深度", 1, 0).ToList();
                    List<double> settlement_list = dex.PassColumnDouble("位移量", 1, 0).ToList();

                    var ds_zip = depth_list.Zip(settlement_list, (d, s) => new { Depth = d, Settlement = s });
                    foreach (var ds in ds_zip)
                    {
                        depth.Add(ds.Depth);
                        sid_displacement.Add(ds.Settlement);
                    }

                    sid_index.Add(depth_list.Count());
                    dex.CloseEx();
                }

                dex.SetData(excel_path, 2 + unit_amount);
                ssi_id = dex.PassColumntString("儀器編號", 1, 0).ToList();
                ssi_displacement = dex.PassColumnDouble("位移量", 1, 0).ToList();
                dex.CloseEx();

                dex.SetData(excel_path, 3 + unit_amount);
                vg_id = dex.PassColumntString("儀器編號", 1, 0).ToList();
                axial_force = dex.PassColumnDouble("軸力", 1, 0).ToList();
                dex.CloseEx();

            }
            catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.ToString()); }

            //dex.SetData(@"\\Mac\Home\Deskt op\excavation\20211209_制式化表單-監測儀器data.xlsx", 4);
            dex.SetData(monitor_path, 4);
            try
            {
                coordinate_n = dex.PassColumnDouble("N座標", 3, 0).ToList();
                coordinate_e = dex.PassColumnDouble("E座標", 3, 0).ToList();
                equipment_id = dex.PassColumntString("儀器編號", 3, 0).ToList();
                settlement = dex.PassColumnDouble("沉陷量", 3, 0).ToList();
                alert = dex.PassRowDouble("高行動值", 3, 0).ToList();

                dex.CloseEx();
            }
            catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

            List<XYZ> monitor_point_list = new List<XYZ>();
            for (i = 0; i != coordinate_n.Count(); i++)
            {
                coordinate_e[i] = coordinate_e[i] * 1000 / 304.8 - projectPosition.EastWest;
                coordinate_n[i] = coordinate_n[i] * 1000 / 304.8 - projectPosition.NorthSouth;
                XYZ monitor_point = new XYZ(coordinate_e[i], coordinate_n[i], 0);
                monitor_point_list.Add(monitor_point);
            }
            tran.Start("start");

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
                colorSettings.ColorSettingsType = AnalysisDisplayStyleColorSettingsType.GradientColor;

                Color deepRed = new Color(166, 0, 0);
                Color red = new Color(255, 44, 01);
                Color green = new Color(0, 253, 0);
                Color lightGreen = new Color(128, 255, 12);
                Color orange = new Color(255, 205, 0);
                Color yellow = new Color(255, 255, 0);
                Color purple = new Color(200, 0, 200);
                Color white = new Color(255, 255, 255);



                colorSettings.MinColor = lightGreen;
                colorSettings.SetIntermediateColors(new List<AnalysisDisplayColorEntry>
                                                {
                                                    new AnalysisDisplayColorEntry(orange),
                                                    new AnalysisDisplayColorEntry(red),
                                                });
                colorSettings.MaxColor = purple;


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
                doc.Delete(sfm.Id);
            }
            catch { TaskDialog.Show("!", "There is no spatial field manager"); }

            sfm = SpatialFieldManager.CreateSpatialFieldManager(doc.ActiveView, 1);
            sfm.SetMeasurementNames(new List<string> { "資料 1" });

            AnalysisResultSchema resultSchema_displacement = new AnalysisResultSchema("位移量", "Description");
            AnalysisResultSchema resultSchema_settlement = new AnalysisResultSchema("沉陷量", "Description2");
            resultSchema_displacement.SetUnits(new List<string> { "mm" }, new List<double> { 1 });
            resultSchema_settlement.SetUnits(new List<string> { "mm" }, new List<double> { 1 });
            int schemaIndex_displacement = sfm.RegisterResult(resultSchema_displacement);
            int schemaIndex_settlement = sfm.RegisterResult(resultSchema_settlement);

            // read angle.tmp
            string temp_path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            temp_path = Path.Combine(temp_path, "temp/angle.txt");
            double angle = Convert.ToDouble(File.ReadAllText(temp_path));

            // rotate all element
            ICollection<ElementId> element_ids = new FilteredElementCollector(doc).WhereElementIsNotElementType().WhereElementIsViewIndependent().ToElementIds();
            ICollection<ElementId> textnote_ids = new FilteredElementCollector(doc).OfClass(typeof(TextNote)).ToElementIds();
            Line origin_axis = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(0, 0, 10));
            //double angle = -48.6 / 180 * Math.PI;
            ElementTransformUtils.RotateElements(doc, element_ids, origin_axis, angle);
            ElementTransformUtils.RotateElements(doc, textnote_ids, origin_axis, angle);

            // rotate base point
            BasePoint basePoint = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).FirstOrDefault() as BasePoint;
            basePoint.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).Set(0);


            double previous_distance = 100000;
            XYZ target_mid = new XYZ(0, 0, 0);
            List<FamilyInstance> target_frame_list = new List<FamilyInstance>();
            List<FamilyInstance> familyInstance = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("支撐")).ToList();
            //try
            //{
            for (int unit_index = 0; unit_index != unit_id_list.Count(); unit_index++)
            {
                TaskDialog.Show("index", unit_index.ToString());
                string detectObject = "監測_土中傾度管";

                // find target wall by unit id
                Wall target_wall = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>()
                            .Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == unit_id_list[unit_index]).First();

                IList<Wall> wall_list = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>().ToList();

                LocationCurve target_lc = target_wall.Location as LocationCurve;
                XYZ target_start = target_lc.Curve.GetEndPoint(0);
                XYZ target_end = target_lc.Curve.GetEndPoint(1);
                target_mid = (target_start + target_end) / 2;

                double target_slope = (target_start.Y - target_end.Y) / (target_start.X - target_end.X);
                double target_normal = -1 / target_slope;
                double target_normal_b = target_mid.Y - target_normal * target_mid.X;

                double except_distance = 50;
                previous_distance = 10000;
                Wall source_wall = target_wall;

                // find unit on the other side
                foreach (Wall w in wall_list)
                {
                    LocationCurve source_lc = w.Location as LocationCurve;
                    XYZ source_start = source_lc.Curve.GetEndPoint(0);
                    XYZ source_end = source_lc.Curve.GetEndPoint(1);
                    XYZ source_mid = (source_start + source_end) / 2;

                    double source_slope = (source_start.Y - source_end.Y) / (source_start.X - source_end.X);

                    double mid_distance = target_mid.DistanceTo(source_mid);
                    double project_distance = Math.Abs(target_normal * source_mid.X + target_normal_b - source_mid.Y);
                    if (project_distance < previous_distance && mid_distance > except_distance && Math.Abs(target_slope - source_slope) < 1)
                    {
                        previous_distance = project_distance;
                        source_wall = w;
                    }
                }

                double width = 10;
                LocationCurve select_source_lc = source_wall.Location as LocationCurve;
                XYZ select_source_start = select_source_lc.Curve.GetEndPoint(0);
                XYZ select_source_end = select_source_lc.Curve.GetEndPoint(1);
                XYZ select_source_mid = (select_source_start + select_source_end) / 2;

                // get monitor family symbol
                FamilySymbol familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == detectObject).First();

                List<string> visual_id = new List<string>();
                List<double> visual_settlement = new List<double>();
                List<XYZ> visual_coord = new List<XYZ>();
                i = 0;

                // find monitor id in excel
                List<string> temp = filtered_equipment_id.GetRange(equipment_id_index[unit_index], equipment_id_index[unit_index + 1]);
                for (i = 0; i != coordinate_n.Count(); i++)
                {
                    string monitor_id = equipment_id[i];

                    if (temp.Contains(monitor_id))
                    {
                        FamilyInstance monitor3 = doc.Create.NewFamilyInstance(monitor_point_list[i], familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        monitor3.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(equipment_id[i] + "/" + settlement[i]);

                        visual_id.Add(equipment_id[i]);
                        visual_settlement.Add(settlement[i]);
                        visual_coord.Add(monitor_point_list[i]);
                    }
                }

                XYZ sum_coord = visual_coord.Aggregate(func: (a, b) => { return a + b; });
                XYZ center = sum_coord / visual_coord.Count();
                List<int> index = FindNearest(target_mid, visual_coord);
                XYZ nearest_visual_coord = visual_coord[index[0]];
                XYZ farest_visual_coord = visual_coord[index[1]];

                XYZ axis = new XYZ(center.X, center.Y, 0);
                Line axis_z = Line.CreateBound(axis, axis + new XYZ(0, 0, 10));

                Transform rotate = Transform.CreateRotationAtPoint(new XYZ(0, 0, 10), Math.Atan(target_normal), new XYZ(axis.X, axis.Y, 0));
                Transform reverse_rotate = Transform.CreateRotationAtPoint(new XYZ(0, 0, 10), -Math.Atan(target_normal), new XYZ(axis.X, axis.Y, 0));

                Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, 0));
                SketchPlane sp = SketchPlane.Create(doc, plane);
                Arc arc3 = Arc.Create(new XYZ(axis.X, axis.Y, 0), 1.05, 0.0, 2.0 * Math.PI, XYZ.BasisX, XYZ.BasisY);
                ModelCurve mc3 = doc.Create.NewModelCurve(arc3, sp);

                XYZ x1 = new XYZ(2 * nearest_visual_coord.X - center.X, center.Y + 1 * width, 0);
                XYZ x2 = new XYZ(2 * nearest_visual_coord.X - center.X, center.Y - 1 * width, 0);
                XYZ x3 = new XYZ(2 * farest_visual_coord.X - center.X, center.Y - 1 * width, 0);
                XYZ x4 = new XYZ(2 * farest_visual_coord.X - center.X, center.Y + 1 * width, 0);

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

                // find target face of floor
                previous_distance = 10000;
                PlanarFace target_face = null;
                GeometryElement geometryElement = floor.get_Geometry(new Options());
                foreach (GeometryObject geometryObject in geometryElement)
                {
                    Solid solid = geometryObject as Solid;

                    if (solid != null)
                    {
                        FaceArray faceArray = solid.Faces;
                        foreach (PlanarFace planarFace in faceArray)
                        {
                            double distance = center.DistanceTo(planarFace.Origin);
                            if (distance <= previous_distance)
                            {
                                previous_distance = distance;
                                target_face = planarFace;
                            }
                        }
                    }
                }

                BoundingBoxUV bb = target_face.GetBoundingBox();
                UV min = bb.Min;
                UV max = bb.Max;

                UV faceCenter = new UV((max.U + min.U) / 2, (max.V + min.V) / 2);
                Transform computeDerivatives = target_face.ComputeDerivatives(faceCenter);
                XYZ faceCenterNormal = computeDerivatives.BasisZ;
                XYZ faceCenterNormalMultiplied = faceCenterNormal.Normalize().Multiply(1);
                Transform transform = Transform.CreateTranslation(faceCenterNormalMultiplied);

                // uv points
                IList<UV> uvPts = new List<UV>();
                for (i = 0; i != visual_id.Count(); i++)
                {
                    visual_coord[i] = reverse_rotate.OfPoint(visual_coord[i]);
                    uvPts.Add(new UV(visual_coord[i].X, visual_coord[i].Y) - new UV(x2.X, x2.Y));
                }

                // val points
                IList<ValueAtPoint> valList = new List<ValueAtPoint>();
                for (i = 0; i != visual_id.Count(); i++)
                {
                    valList.Add(new ValueAtPoint(new List<double> { visual_settlement[i] }));
                }

                int idx = sfm.AddSpatialFieldPrimitive(target_face, transform);
                sfm.UpdateSpatialFieldPrimitive(idx, new FieldDomainPointsByUV(uvPts), new FieldValues(valList), schemaIndex_settlement);

                // find target face of unit
                previous_distance = 10000;
                target_face = null;
                geometryElement = target_wall.get_Geometry(new Options());

                XYZ visual_vector = nearest_visual_coord - farest_visual_coord;
                foreach (GeometryObject geometryObject in geometryElement)
                {
                    Solid solid = geometryObject as Solid;

                    if (solid != null)
                    {
                        FaceArray faceArray = solid.Faces;
                        foreach (PlanarFace planarFace in faceArray)
                        {
                            double dot = visual_vector.DotProduct(planarFace.FaceNormal);
                            double mul = visual_vector.GetLength() * planarFace.FaceNormal.GetLength();
                            double distance = nearest_visual_coord.DistanceTo(planarFace.Origin);
                            if (Math.Abs(dot + mul) < 0.01 && distance <= previous_distance)
                            {
                                previous_distance = distance;
                                target_face = planarFace;
                            }
                        }
                    }
                }

                bb = target_face.GetBoundingBox();
                min = bb.Min;
                max = bb.Max;

                faceCenter = new UV((max.U + min.U) / 2, (max.V + min.V) / 2);
                computeDerivatives = target_face.ComputeDerivatives(faceCenter);
                faceCenterNormal = computeDerivatives.BasisZ;
                faceCenterNormalMultiplied = faceCenterNormal.Normalize().Multiply(1);
                transform = Transform.CreateTranslation(faceCenterNormalMultiplied);

                int n = sid_index[unit_index + 1];
                double delta = (max.V - min.V) / n;

                // uv points
                uvPts = new List<UV>();
                List<double> temp_depth = depth.GetRange(sid_index[unit_index], sid_index[unit_index + 1]);
                for (i = 0; i < n; i++)
                {
                    uvPts.Add(new UV(max.U / 2, max.V - temp_depth[i] * 3.281));
                    if (max.V - temp_depth[i] * 3.281 < 0)
                    {
                        break;
                    }
                }

                // val points
                valList = new List<ValueAtPoint>();
                List<double> temp_sid = sid_displacement.GetRange(sid_index[unit_index], sid_index[unit_index + 1]);
                for (i = 0; i < uvPts.Count(); i++)
                {
                    valList.Add(new ValueAtPoint(new List<double> { temp_sid[i] }));
                }

                idx = sfm.AddSpatialFieldPrimitive(target_face, transform);
                sfm.UpdateSpatialFieldPrimitive(idx, new FieldDomainPointsByUV(uvPts), new FieldValues(valList), schemaIndex_displacement);

                sfm.LegendPosition = new XYZ(target_mid.X, target_mid.Y, 0);

                // find nearest frame
                previous_distance = 10000;
                FamilyInstance t = familyInstance[0];
                foreach (FamilyInstance f in familyInstance)
                {
                    LocationCurve locationCurve = f.Location as LocationCurve;
                    XYZ start = locationCurve.Curve.GetEndPoint(0);
                    XYZ end = locationCurve.Curve.GetEndPoint(1);

                    double start_distance = target_mid.DistanceTo(start);
                    double end_distance = target_mid.DistanceTo(end);
                    if (start_distance < previous_distance || end_distance < previous_distance)
                    {
                        if (start_distance < end_distance) { previous_distance = start_distance; }
                        else { previous_distance = end_distance; }

                        t = f;
                    }
                }
                target_frame_list.Add(t);

                LocationCurve t_lc = t.Location as LocationCurve;

                // put text note
                List<Level> level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name.Contains("斷面typeS1-開挖階數")).ToList();
                for (i = 0; i != vg_id.Count(); i += 2)
                {
                    XYZ t_mid = (t_lc.Curve.GetEndPoint(0) + t_lc.Curve.GetEndPoint(1)) / 2;
                    double z = Math.Round(level[i / 2].Elevation);
                    XYZ point = new XYZ(t_mid.X, t_mid.Y, z);
                    string text = vg_id[i] + ": " + axial_force[i] + "\n" + vg_id[i + 1] + ": " + axial_force[i + 1];
                    TextNote note = TextNote.Create(doc, view.Id, point, text, textNoteOptions);
                }

            }
            //}
            //catch (Exception e) { TaskDialog.Show("Error", e.ToString()); }

            ICollection<ElementId> not_hide_id = new List<ElementId>();
            foreach (FamilyInstance t in target_frame_list)
            {
                LocationCurve t_lc = t.Location as LocationCurve;
                double t_length = t_lc.Curve.Length;
                XYZ t_start = t_lc.Curve.GetEndPoint(0);

                foreach (FamilyInstance f in familyInstance)
                {
                    LocationCurve lc = f.Location as LocationCurve;
                    double length = lc.Curve.Length;
                    XYZ start = lc.Curve.GetEndPoint(0);

                    t_start = new XYZ(t_start.X, t_start.Y, 0);
                    start = new XYZ(start.X, start.Y, 0);

                    if (t_start.DistanceTo(start) < 10)// && Math.Abs(length - t_length) < 10)
                    {
                        not_hide_id.Add(f.Id);
                    }
                }
            }

            List<FamilyInstance> middle_column = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("中間樁")).ToList();
            List<FamilyInstance> slope_column = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains("斜撐")).ToList();
            ICollection<ElementId> hide_id = new List<ElementId>();
            foreach (FamilyInstance f in familyInstance)
            {
                if (!not_hide_id.Contains(f.Id))
                {
                    hide_id.Add(f.Id);
                }
            }
            foreach (FamilyInstance f in middle_column)
            {
                if (!not_hide_id.Contains(f.Id))
                {
                    hide_id.Add(f.Id);
                }
            }
            foreach (FamilyInstance f in slope_column)
            {
                if (!not_hide_id.Contains(f.Id))
                {
                    hide_id.Add(f.Id);
                }
            }
            view.HideElementsTemporary(hide_id);

            TaskDialog.Show("1", "done");
            tran.Commit();
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }

        public List<int> FindNearest(XYZ element_point, List<XYZ> point_list)
        {
            double distance = element_point.DistanceTo(point_list[0]);
            int i = 0;
            List<int> index = new List<int> { 0, 0 };

            foreach (XYZ point in point_list)
            {
                if (element_point.DistanceTo(point) < distance)
                {
                    distance = element_point.DistanceTo(point);
                    index[0] = i;
                }
                else if (element_point.DistanceTo(point) > distance)
                {
                    distance = element_point.DistanceTo(point);
                    index[1] = i;
                }
                i++;
            }

            return index;

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