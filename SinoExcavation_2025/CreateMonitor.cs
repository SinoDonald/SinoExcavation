using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using ExReaderConsole;

namespace SinoExcavation_2025
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
        public string excel_path
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            Transaction tran = new Transaction(doc);

            List<double> coordinate_n = new List<double>();
            List<double> coordinate_e = new List<double>();
            List<string> equipment_id = new List<string>();
            ProjectPosition projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);

            int i = 0;
            tran.Start("start");

            View view = doc.ActiveView;

            TextNoteOptions textNoteOptions = new TextNoteOptions();
            TextNoteType tType = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();
            textNoteOptions.HorizontalAlignment = HorizontalTextAlignment.Left;
            textNoteOptions.TypeId = tType.Id;

            //讀取資料
            ExReader dex = new ExReader();
            dex.SetData(excel_path, 1);
            try
            {
                coordinate_n = dex.PassColumnDouble("N座標", 3, 0).ToList();
                coordinate_e = dex.PassColumnDouble("E座標", 3, 0).ToList();
                equipment_id = dex.PassColumntString("儀器編號", 3, 0).ToList();

                dex.CloseEx();
            }
            catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

            List<ElementId> elementIds = new List<ElementId>();


            i = 0;
            for (i = 0; i != coordinate_n.Count(); i++)
            {

                coordinate_e[i] = coordinate_e[i] * 1000 / 304.8 - projectPosition.EastWest;
                coordinate_n[i] = coordinate_n[i] * 1000 / 304.8 - projectPosition.NorthSouth;

                XYZ origin = new XYZ(coordinate_e[i], coordinate_n[i], 0);
                string detectObject = "監測_土中傾度管";

                FamilySymbol familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == detectObject).First();
                //familySymbol.LookupParameter("儀器編號").Set(equipment_id[i]);
                //familySymbol.LookupParameter("N").SetValueString(coordinate_n[i].ToString());
                //familySymbol.LookupParameter("E").SetValueString(coordinate_e[i].ToString());
                FamilyInstance monitor = doc.Create.NewFamilyInstance(origin, familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                TextNote note = TextNote.Create(doc, view.Id, origin, equipment_id[i], textNoteOptions);

                elementIds.Add(monitor.Id);
                //TaskDialog.Show("1", monitor.Id.ToString());


                //if (i >= 10) { break; }
            }

            
            tran.Commit();


            TaskDialog.Show("done", i.ToString());
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}