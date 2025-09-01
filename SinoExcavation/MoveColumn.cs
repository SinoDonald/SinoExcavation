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

    class MoveColumn : IExternalEventHandler
    {
        public string SectionName
        {
            get;
            set;
        }
        public IList<double> ShiftData
        {
            //0:shift_x; 1:shift_y;
            get;
            set;
        }

        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //撈選該斷面之柱
            ICollection<ElementId> columns_id = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList()
                                                   where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面"+SectionName
                                                   select x.Id).ToList();

            //撈選出該斷面之房間
            Room room = null;
            foreach (SpatialElement spelement in new FilteredElementCollector(doc).OfClass(typeof(SpatialElement)).Cast<SpatialElement>().ToList())
            {
                Room r = spelement as Room;
                if (r.Name.Split(' ')[0] == "斷面" + SectionName)
                {
                    room = r;
                }
            }

            IList<BoundarySegment> boundarySegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions())[0];
            int wall_count = boundarySegments.Count;

            //取得連續壁內座標點
            XYZ[] innerwall_points = new XYZ[wall_count + 1];
            for (int i = 0; i < wall_count; i++)
            {
                //inner
                innerwall_points[i] = boundarySegments[i].GetCurve().Tessellate()[0];
            }
            innerwall_points[wall_count] = innerwall_points[0];

            Transaction trans_2 = new Transaction(doc);
            trans_2.Start("交易開始");
            //開始移動中間樁
            ElementTransformUtils.MoveElements(doc, columns_id, new XYZ(ShiftData[0]*1000/304.8, ShiftData[1] * 1000 / 304.8, 0));
            ICollection<FamilyInstance> columns_instance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList()
                                                            where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + SectionName
                                                            select x).ToList();
            foreach(FamilyInstance column in columns_instance)
            {
                if (!IsInPolygon((column.Location as LocationPoint).Point, innerwall_points))
                {
                    doc.Delete(column.Id);
                }
            }
            
            trans_2.Commit();
        }

        //判斷點是否位於開挖範圍內
        public static bool IsInPolygon(XYZ checkPoint, XYZ[] polygonPoints)
        {
            bool inside = false;
            int pointCount = polygonPoints.Count<XYZ>();
            XYZ p1, p2;
            for (int i = 0, j = pointCount - 1; i < pointCount; j = i, i++)
            {
                p1 = polygonPoints[i];
                p2 = polygonPoints[j];
                if (checkPoint.Y < p2.Y)
                {
                    if (p1.Y <= checkPoint.Y)
                    {
                        if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) > (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                        {

                            inside = (!inside);
                        }
                    }
                }
                else if (checkPoint.Y < p1.Y)
                {

                    if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) < (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                    {
                        inside = (!inside);
                    }
                }
            }
            return inside;
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
