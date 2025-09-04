using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;

namespace SinoExcavation_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class PutDetect : IExternalEventHandler
    {

        public string detectObject
        {
            get;
            set;
        }
        public string Levelname
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            try
            {
                Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);

                Document doc = uidoc.Document;

                Transaction T = new Transaction(doc);
                View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().First();

                ViewPlan viewPlan;
                //在選取階層後，判斷是否已經建置該階層的level並指定此level的視圖(try的部分)
                //若沒有，則新增一個該階層的level與視圖(catch部分)
                try
                {
                    viewPlan = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().Where(x => x.Name == Levelname).First();
                }
                catch
                {
                    Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name == Levelname).First();

                    ViewFamilyType viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().Where(x => x.Name == "樓板平面圖").First();


                    ExReader cop = new ExReader();

                    T.Start("create view");

                    viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, level.Id);
                    //創建視圖範圍
                    PlanViewRange planViewRange = viewPlan.GetViewRange();

                    planViewRange.SetOffset(PlanViewPlane.ViewDepthPlane, -2 * 1000 / 304.8);

                    planViewRange.SetOffset(PlanViewPlane.BottomClipPlane, -1 * 1000 / 304.8);

                    viewPlan.SetViewRange(planViewRange);

                    T.Commit();
                }




                //更改當前視圖
                uidoc.ActiveView = viewPlan;

                //指定所選擇之監測儀器
                FamilySymbol familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == detectObject).First();

                //進入繪製階段
                try
                {
                    uidoc.PromptForFamilyInstancePlacement(familySymbol);
                }
                catch { };

                //繪製完跳出提示視窗
                TaskDialog.Show("通知", "放置完成");

                //回到3D視圖
                uidoc.ActiveView = view3D;

                T.Start("刪除視圖");

                if (viewPlan.Name != "樓層 0")
                {
                    doc.Delete(viewPlan.Id);
                }

                T.Commit();
            }
            catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message + e.StackTrace); }


        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }

    }
}
