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
                //TaskDialog.Show("1", );
                AnalysisDisplayColoredSurfaceSettings coloredSurfaceSettings = new AnalysisDisplayColoredSurfaceSettings();
                coloredSurfaceSettings.ShowGridLines = true;

                AnalysisDisplayColorSettings colorSettings = new AnalysisDisplayColorSettings();
                Color orange = new Color(255, 205, 0);
                Color purple = new Color(200, 0, 200);
                colorSettings.MaxColor = orange;
                colorSettings.MinColor = purple;

                AnalysisDisplayLegendSettings legendSettings = new AnalysisDisplayLegendSettings();
                legendSettings.NumberOfSteps = 10;
                legendSettings.Rounding = 0.05;
                legendSettings.ShowDataDescription = false;
                legendSettings.ShowLegend = true;

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
            if (null == sfm)
            {
                sfm = SpatialFieldManager.CreateSpatialFieldManager(doc.ActiveView, 1);
            }

            Reference reference = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select a face");
            int idx = sfm.AddSpatialFieldPrimitive(reference);

            Face face = doc.GetElement(reference).GetGeometryObjectFromReference(reference) as Face;

            IList<UV> uvPts = new List<UV>();
            BoundingBoxUV bb = face.GetBoundingBox();
            UV min = bb.Min;
            UV max = bb.Max;
            uvPts.Add(new UV(min.U, min.V));
            uvPts.Add(new UV(max.U, max.V));

            FieldDomainPointsByUV pnts = new FieldDomainPointsByUV(uvPts);

            List<double> doubleList = new List<double>();
            IList<ValueAtPoint> valList = new List<ValueAtPoint>();
            doubleList.Add(0);
            valList.Add(new ValueAtPoint(doubleList));
            doubleList.Clear();
            doubleList.Add(10);
            valList.Add(new ValueAtPoint(doubleList));
            TaskDialog.Show("1", valList.ToString());
            FieldValues vals = new FieldValues(valList);

            AnalysisResultSchema resultSchema = new AnalysisResultSchema("Schema Name", "Description");
            int schemaIndex = sfm.RegisterResult(resultSchema);
            sfm.UpdateSpatialFieldPrimitive(idx, pnts, vals, schemaIndex);
            subtran.Commit();
            tran.Commit();


            return Result.Succeeded;
        }
    }
}