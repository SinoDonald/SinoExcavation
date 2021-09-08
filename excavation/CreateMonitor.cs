using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using ExReaderConsole;

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
            Transaction tran = new Transaction(doc);

            List<double> coordinate_n = new List<double>();
            List<double> coordinate_e = new List<double>();
            List<string> equipment_id = new List<string>();
            ProjectPosition projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero);

            int i = 0;
            tran.Start("start");
            if (files_path[0].Contains("dwg"))
            {
                View view = doc.ActiveView;
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

                        }
                    }
                    catch { }
                }
                doc.Delete(dwg.Id);

            } else if (files_path[0].Contains("xls") || files_path[0].Contains("XLS"))
            {
                View view = doc.ActiveView;

                TextNoteOptions textNoteOptions = new TextNoteOptions();
                TextNoteType tType = new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Cast<TextNoteType>().FirstOrDefault();
                textNoteOptions.HorizontalAlignment = HorizontalTextAlignment.Left;
                textNoteOptions.TypeId = tType.Id;

                //讀取資料
                ExReader dex = new ExReader();
                dex.SetData(files_path[0], 1);
                try
                {
                    coordinate_n = dex.PassMonitorDouble("N座標", 3).ToList();
                    coordinate_e = dex.PassMonitorDouble("E座標", 3).ToList();
                    equipment_id = dex.PassMonitortString("儀器編號", 3).ToList();

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
                
            }
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

            TaskDialog.Show("done", i.ToString());
        }

        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
