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
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace SinoExcavation_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class MakeSheet : IExternalEventHandler
    {
        public OpenFileDialog openFileDialog
        {
            get;
            set;
        }
        public int sectionLineNumber
        {
            get;
            set;
        }
        public IList<string> sectionName
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            Transaction t = new Transaction(doc);

            //開始圖紙建立
            t.Start("創建圖紙");
            ElementId elementId = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks).ToElements().First().Id;
            //建立圖紙
            ViewSheet viewSheet = ViewSheet.Create(doc, ElementId.InvalidElementId);
            //載入圖紙
            Autodesk.Revit.DB.View view = viewSheet as Autodesk.Revit.DB.View;
            DWGImportOptions dWGImportOptions = new DWGImportOptions();
            dWGImportOptions.ColorMode = ImportColorMode.Preserved;
            dWGImportOptions.Placement = ImportPlacement.Centered;
            dWGImportOptions.Unit = ImportUnit.Centimeter;

            //將instance匯入圖紙上
            LinkLoadResult linkLoadResult = new LinkLoadResult();
            ImportInstance toz = ImportInstance.Create(doc, view, openFileDialog.FileName, dWGImportOptions, out linkLoadResult);
            int view_count = sectionLineNumber; //選擇放入之數量

            //選出欲匯入圖紙的剖面視圖
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter viewersFilter = new ElementCategoryFilter(BuiltInCategory.OST_Viewers);
            IList<Element> viewersList = collector.WherePasses(viewersFilter).WhereElementIsNotElementType().Where(y => y.Name[0] == '剖').ToList<Element>();
            BoundingBoxXYZ xYZ = toz.get_BoundingBox(view);

            //剖面數＝1 
            if (view_count == 1)
            {
                //算座標
                XYZ loc = (xYZ.Max + xYZ.Min) / 2;
                XYZ[] locations = { loc };

                //放入視圖
                SetViewport(doc, view_count, locations, viewSheet);
            }

            //剖面數＝2
            if (view_count == 2)
            {
                //算座標
                
                double x = (xYZ.Max.X + xYZ.Min.X) / 2;
                double z = (xYZ.Max.Y + xYZ.Min.Y) / 2;
                double y1 = xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75;
                double y2 = xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25;
                XYZ loc = new XYZ(x, y1, z);
                XYZ loc2 = new XYZ(x, y2, z);

                XYZ[] locations = { loc, loc2 };
                //放入視圖
                SetViewport(doc, view_count, locations, viewSheet);
            }

            //剖面數＝4
            if (view_count == 4)
            {
                //算座標
                
                double y1 = xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75;
                double y2 = xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25;
                double z = (xYZ.Max.Y + xYZ.Min.Y) / 2;
                double x1 = xYZ.Max.X * 0.25 + xYZ.Min.X * 0.75;
                double x2 = xYZ.Max.X * 0.75 + xYZ.Min.X * 0.25;
                XYZ loc = new XYZ(x1, y1, z);
                XYZ loc2 = new XYZ(x2, y1, z);
                XYZ loc3 = new XYZ(x1, y2, z);
                XYZ loc4 = new XYZ(x2, y2, z);

                XYZ[] locations = { loc, loc2, loc3, loc4 };
                //放入視圖
                SetViewport(doc, view_count, locations, viewSheet);
            }

            //剖面數＝6
            if (view_count == 6)
            {
                //算座標
                
                double y1 = xYZ.Max.Y * 0.25 + xYZ.Min.Y * 0.75;
                double y2 = xYZ.Max.Y * 0.75 + xYZ.Min.Y * 0.25;
                double z = (xYZ.Max.Y + xYZ.Min.Y) / 2;
                double x1 = (xYZ.Max.X - xYZ.Min.X) * (0.25) + xYZ.Min.X;
                double x2 = (xYZ.Max.X - xYZ.Min.X) / 2 + xYZ.Min.X;
                double x3 = (xYZ.Max.X - xYZ.Min.X) * (0.75) + xYZ.Min.X;
                XYZ loc = new XYZ(x1, y1, z);
                XYZ loc2 = new XYZ(x2, y1, z);
                XYZ loc3 = new XYZ(x3, y1, z);
                XYZ loc4 = new XYZ(x1, y2, z);
                XYZ loc5 = new XYZ(x2, y2, z);
                XYZ loc6 = new XYZ(x3, y2, z);
                XYZ[] locations = { loc, loc2, loc3, loc4, loc5, loc6 };

                //放入視圖
                SetViewport(doc, view_count, locations, viewSheet);
            }
            
            t.Commit();
            //完成剖面圖紙出圖

            //切往所建立圖紙之視圖
            uidoc.ActiveView = view;

            TaskDialog.Show("出圖", "完成出圖!!!");
            //throw new NotImplementedException();
        }

        //放入視圖
        public void SetViewport(Document doc, int Numbers, XYZ[] locations, ViewSheet viewSheet)
        {
            for (int i = 0; i < Numbers; i++)
            {
                try
                {
                    //剖面圖
                    ViewSection viewSection = new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Where(y => y.Name == sectionName[i]).First();
                    ViewSection depentviewsection = doc.GetElement(viewSection.Duplicate(ViewDuplicateOption.AsDependent)) as ViewSection;

                    //剖面圖參數
                    depentviewsection.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC).Set("自訂");
                    depentviewsection.Scale = Numbers * 100;

                    Viewport viewport1 = Viewport.Create(doc, viewSheet.Id, depentviewsection.Id, locations[i]);
                }
                catch
                {
                    //平面圖
                    ViewPlan viewPlane = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().Where(y => y.Name == sectionName[i]).First();
                    ViewPlan depentviewplan = doc.GetElement(viewPlane.Duplicate(ViewDuplicateOption.AsDependent)) as ViewPlan;

                    //平面圖參數
                    depentviewplan.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC).Set("自訂");
                    depentviewplan.Scale = Numbers * 100;

                    Viewport viewport1 = Viewport.Create(doc, viewSheet.Id, depentviewplan.Id, locations[i]);
                }
            }

        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
