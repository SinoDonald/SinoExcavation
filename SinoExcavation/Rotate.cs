using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using ExReaderConsole;
using System.IO;


namespace SinoExcavation
{
    class Rotate : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            Transaction tran = new Transaction(doc);

            tran.Start("rotate");
            string temp_path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            temp_path = Path.Combine(temp_path, @"temp\angle.txt");

            ICollection<ElementId> element_ids = new FilteredElementCollector(doc).WhereElementIsNotElementType().WhereElementIsViewIndependent().ToElementIds();
            ICollection<ElementId> textnote_ids = new FilteredElementCollector(doc).OfClass(typeof(TextNote)).ToElementIds();
            
            Line origin_axis = Line.CreateBound(new XYZ(0, 0, 0), new XYZ(0, 0, 10));
            double angle = 0;
            if (File.Exists(temp_path))
            {
                // read existing file
                angle = -Convert.ToDouble(File.ReadAllText(temp_path));
                
            }
            else
            {
                List<double> slope_list = new List<double>();
                IList<Element> walls = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Walls).ToElements();
                foreach (Element e in walls)
                {
                    Plane plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, new XYZ(0, 0, 0));
                    SketchPlane sp = SketchPlane.Create(doc, plane);

                    Wall wall = e as Wall;
                    LocationCurve lc = wall.Location as LocationCurve;
                    int length = Convert.ToInt32(lc.Curve.Length / 10);

                    XYZ start_point = lc.Curve.GetEndPoint(0);
                    start_point = new XYZ(start_point.X, start_point.Y, 0);

                    XYZ end_point = lc.Curve.GetEndPoint(1);
                    end_point = new XYZ(end_point.X, end_point.Y, 0);

                    double wall_slope = Math.Atan((start_point.X - end_point.X) / (start_point.Y - end_point.Y));
                    slope_list.Add(wall_slope);

                }

                // find the most slope in slope list as the rotation angle
                angle = slope_list.GroupBy(i => i).OrderByDescending(g => g.Count()).Select(g => g.Key).First();
            }

            // rotate element
            ElementTransformUtils.RotateElements(doc, element_ids, origin_axis, angle);
            try
            {
                ElementTransformUtils.RotateElements(doc, textnote_ids, origin_axis, angle);
            }
            catch {  }

            if (Math.Abs(angle) > 0.001)
            {
                // save temp file
                using (StreamWriter outputFile = new StreamWriter(temp_path))
                {
                    outputFile.WriteLine((angle).ToString());
                }

                temp_path = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                temp_path = Path.Combine(temp_path, @"temp\");
                TaskDialog.Show("success", "angle.tmp was save at: " + temp_path);
            }
            else
            {
                TaskDialog.Show("warning", "angle is too small");
            }

            tran.Commit();
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
