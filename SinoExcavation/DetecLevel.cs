using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;

namespace SinoExcavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class DetecLevel : IExternalEventHandler
    {
        public IList<string> FloorName
        {
            get;
            set;
        }
        //為了讓樓層能夠顯示在監測儀器的combobox內，用此程式先抓取目前有的樓層，放入public參數內
        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            //蒐集樓層
            IList<Level> Level_list = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList().ToList();

            FloorName = new List<string>();
            //加入參數
            foreach (Level lev in Level_list)
            {
                FloorName.Add(lev.Name);
            }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
