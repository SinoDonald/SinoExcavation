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

            List<string> equipment_id = new List<string>();
            List<double> settlement = new List<double>();
            List<double> depth = new List<double>();

            Transaction tran = new Transaction(doc);
            SubTransaction subtran = new SubTransaction(doc);
            tran.Start("start");

            subtran.Start();
            AnalysisDisplayStyle analysisDisplayStyle = null;

            FilteredElementCollector collector1 = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector1.OfClass(typeof(AnalysisDisplayStyle)).ToElements();
            var displayStyle = from element in collection
                               where element.Name == "Display Style 2"
                               select element;

            //讀取資料
            ExReader dex = new ExReader();
            dex.SetData(@"\\Mac\Home\Desktop\excavation\20211209_制式化表單-監測儀器data.xlsx", 1);
            try
            {
                equipment_id = dex.PassColumntString("儀器編號", 1).ToList();
                settlement = dex.PassColumnDouble("位移量", 1).ToList();
                depth = dex.PassColumnDouble("觀測深度", 1).ToList();

                dex.CloseEx();
            }
            catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

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
                                                        new AnalysisDisplayColorEntry(red, -20),
                                                        new AnalysisDisplayColorEntry(orange, 10),
                                                        new AnalysisDisplayColorEntry(lightGreen, 20),
                                                        new AnalysisDisplayColorEntry(orange, 30),

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

            Reference reference = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select a face");
            int idx = sfm.AddSpatialFieldPrimitive(reference);

            Face face = doc.GetElement(reference).GetGeometryObjectFromReference(reference) as Face;

            ProjectPosition projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);

            XYZ translationVector = new XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation);

            TaskDialog.Show("1", translationVector.ToString());

            IList<UV> uvPts = new List<UV>();
            BoundingBoxUV bb = face.GetBoundingBox();
            UV min = bb.Min;
            UV max = bb.Max;

            int n = settlement.Count();
            double delta = (max.V - min.V) / n;

            for (int i = 0; i < n; i++)
            {
                uvPts.Add(new UV(max.U, max.V - delta * i));
                //uvPts.Add(new UV(min.U, max.V - delta * i));
            }

            UV faceCenter = new UV((max.U + min.U) / 2, (max.V + min.V) / 2);
            Transform computeDerivatives = face.ComputeDerivatives(faceCenter);
            XYZ faceCenterNormal = computeDerivatives.BasisZ;

            XYZ faceCenterNormalMultiplied = faceCenterNormal.Normalize().Multiply(2.5);
            Transform transform = Transform.CreateTranslation(faceCenterNormalMultiplied);

            TaskDialog.Show("1", idx.ToString());

            FieldDomainPointsByUV pnts = new FieldDomainPointsByUV(uvPts);

            IList<ValueAtPoint> valList = new List<ValueAtPoint>();

            for (int i = 0; i < n; i++)
            {
                valList.Add(new ValueAtPoint(new List<double> { settlement[i] }));
            }

            FieldValues vals = new FieldValues(valList);

            AnalysisResultSchema resultSchema = new AnalysisResultSchema("鋼筋應力", "Description");
            resultSchema.SetUnits(new List<string> { "kg/cm^2" }, new List<double> { 1 });

            sfm.SetMeasurementNames(new List<string> { "資料 1" });

            int schemaIndex = sfm.RegisterResult(resultSchema);
            sfm.UpdateSpatialFieldPrimitive(idx, pnts, vals, schemaIndex);

            
            subtran.Commit();
            tran.Commit();

            return Result.Succeeded;
        }
    }
}