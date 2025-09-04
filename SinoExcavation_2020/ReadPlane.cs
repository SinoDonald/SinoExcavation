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

namespace SinoExcavation_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class ReadPlane : IExternalEventHandler
    {
        public delegate void ReturnPlane(IList<string> planes_name);//委派
        public event ReturnPlane ReturnPlanes;//事件
        public void Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            //載入目前有的平面圖及剖面圖
            IList<string> Viewsections = new FilteredElementCollector(doc).OfClass(typeof(ViewSection)).Cast<ViewSection>().Select(x => x.Name).ToList();
            IList<string> Viewplanes = new FilteredElementCollector(doc).OfClass(typeof(ViewPlan)).Cast<ViewPlan>().Select(x => x.Name).ToList();
            IList<string> all_view = Viewsections;
            all_view = all_view.Concat(Viewplanes).ToList();
            ReturnPlanes(all_view);//執行委派的方法到這個事件
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
