using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Analysis;
using ExReaderConsole;
using System.Diagnostics;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class DisplayColor : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;

            Transaction tran = new Transaction(doc);
            SubTransaction subtran = new SubTransaction(doc);
            tran.Start("start");

            subtran.Start();
            AnalysisDisplayStyle analysisDisplayStyle = null;

            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector1.OfClass(typeof(AnalysisDisplayStyle)).ToElements();
            var displayStyle = from element in collection
                               where element.Name == "Display Style 1"
                               select element;

            // If display style does not already exist in the document, create it
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
                                                        new AnalysisDisplayColorEntry(red, -3000),
                                                        new AnalysisDisplayColorEntry(orange, -2500),
                                                        new AnalysisDisplayColorEntry(lightGreen, 2500),
                                                        new AnalysisDisplayColorEntry(orange, 3000),

                                                    });
                
                colorSettings.MinColor = purple;

                AnalysisDisplayLegendSettings legendSettings = new AnalysisDisplayLegendSettings
                {
                    NumberOfSteps = 10,
                    Rounding = 0.05,
                    ShowDataDescription = false,
                    ShowLegend = true
                    
                };                


                analysisDisplayStyle = AnalysisDisplayStyle.CreateAnalysisDisplayStyle(doc, "Display Style 1", coloredSurfaceSettings, colorSettings, legendSettings);
            }
            else
            {
                analysisDisplayStyle = displayStyle.Cast<AnalysisDisplayStyle>().ElementAt<AnalysisDisplayStyle>(0);
            }

            // now assign the display style to the view
            doc.ActiveView.AnalysisDisplayStyleId = analysisDisplayStyle.Id;
            subtran.Commit();

            subtran.Start();
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

            ProjectPosition projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);

            XYZ translationVector = new XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation);

            TaskDialog.Show("1", translationVector.ToString());

            /*
            ElementId ele_id = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select a face").ElementId;
            Element ele = doc.GetElement(ele_id);
            //int idx = sfm.AddSpatialFieldPrimitive(reference);
            //TaskDialog.Show("1", idx.ToString());
            //Wall wall = doc.GetElement(reference) as Wall;
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.wall);
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> elementss = collector.WherePasses(filter).ToElements();
            foreach (Element element in elementss)
            {
            double x = ele.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble();
            double y = ele.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble();
            double elevation = ele.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble();
            TaskDialog.Show("1", x.ToString());
            */


            //90 * 0.74029 * 304.8
            /*
            IList<UV> uvPts = new List<UV>();
            BoundingBoxUV bb = face.GetBoundingBox();
            UV min = bb.Min;
            UV max = bb.Max;
            uvPts.Add(new UV(max.U, max.V));
            uvPts.Add(new UV(min.U, max.V));

            uvPts.Add(new UV(max.U, max.V/2));
            uvPts.Add(new UV(min.U, max.V/2));

            uvPts.Add(new UV(max.U, min.V));
            uvPts.Add(new UV(min.U, min.V));

            UV faceCenter = new UV((max.U + min.U) / 2, (max.V + min.V) / 2);
            Transform computeDerivatives = face.ComputeDerivatives(faceCenter);
            XYZ faceCenterNormal = computeDerivatives.BasisZ;

            XYZ faceCenterNormalMultiplied = faceCenterNormal.Normalize().Multiply(2.5);
            Transform transform = Transform.CreateTranslation(faceCenterNormalMultiplied);

            //SpatialFieldManager sfm = SpatialFieldManager.CreateSpatialFieldManager(doc.ActiveView, 1);

            //int idx = sfm.AddSpatialFieldPrimitive(face, transform);
            TaskDialog.Show("1", idx.ToString());

            FieldDomainPointsByUV pnts = new FieldDomainPointsByUV(uvPts);

            IList<ValueAtPoint> valList = new List<ValueAtPoint>();
            valList.Add(new ValueAtPoint(new List<double> { -133 }));
            valList.Add(new ValueAtPoint(new List<double> { 29 }));

            valList.Add(new ValueAtPoint(new List<double> { 4000 }));
            valList.Add(new ValueAtPoint(new List<double> { -430 }));

            valList.Add(new ValueAtPoint(new List<double> { -158 }));
            valList.Add(new ValueAtPoint(new List<double> { -2900 }));
            FieldValues vals = new FieldValues(valList);

            AnalysisResultSchema resultSchema = new AnalysisResultSchema("鋼筋應力", "Description");
            resultSchema.SetUnits(new List<string> { "kg/cm^2" }, new List<double> { 1 });

            sfm.SetMeasurementNames(new List<string> { "資料 1" });

            int schemaIndex = sfm.RegisterResult(resultSchema);
            sfm.UpdateSpatialFieldPrimitive(idx, pnts, vals, schemaIndex);

            */
            subtran.Commit();
            tran.Commit();

            return Result.Succeeded;
        }
    }
}