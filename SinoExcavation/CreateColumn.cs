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

    class CreateColumn : IExternalEventHandler
    {
        //讀取深開挖資料xlsx檔案
        public IList<string> files_path
        {
            get;
            set;
        }
        public IList<double> ShiftData
        {
            //0:column_dis; 1:x_remainder; 2:y_remainder
            get;
            set;
        }
        //從使用者介面讀取x,y值（平移量）
        public IList<double> xy_shift
        {
            get;
            set;
        }

        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            foreach (string file_path in files_path)
            {
                //API CODE START

                //偏移量
                double xshift = xy_shift[0];
                double yshift = xy_shift[1];
                try
                {
                    //讀取資料
                    ExReader dex = new ExReader();
                    dex.SetData(file_path, 1);
                    try
                    {
                        dex.PassColumnData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

                    //篩選並集合樓層元件
                    ICollection<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                    Level levdeep = levels.Where(x => x.Name.Contains("斷面" + dex.section) && !(x.Name.Contains("擋土壁深度"))).OrderBy(x => x.Elevation).ToList()[0];


                    //建立房間元件藉此得到連續壁內座標點
                    Transaction trans_1 = new Transaction(doc);
                    trans_1.Start("create room & clear columns");
                    double position_x_average = 0;
                    double position_y_average = 0;
                    try
                    {
                        dex.excaRange.RemoveAt(dex.excaRange.Count - 1);
                        position_x_average = (dex.excaRange.Average(x => x.Item1) - xshift) * 1000 / 304.8;
                        position_y_average = (dex.excaRange.Average(x => x.Item2) - yshift) * 1000 / 304.8;
                    } catch { }
                    Room room = null;


                    //選出該斷面房間
                    foreach (SpatialElement spelement in new FilteredElementCollector(doc).OfClass(typeof(SpatialElement)).Cast<SpatialElement>().ToList())
                    {
                        Room r = spelement as Room;
                        if (r.Name.Split(' ')[0] == "斷面" + dex.section)
                        {
                            room = r;
                        }
                    }
                    //若沒有則建立新房間
                    if (room == null)
                    {
                        room = doc.Create.NewRoom(levdeep, new UV(position_x_average, position_y_average));
                        room.Name = "斷面" + dex.section;
                        room.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section);
                    }

                    ICollection<ElementId> old_columns = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>() where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == String.Format("斷面{0}:中間樁", dex.section) select x.Id).ToList();
                    
                    //刪除既存在之柱
                    if (old_columns.Count != 0)
                    {
                        doc.Delete(old_columns);
                    }

                    trans_1.Commit();
                    int wall_count = 0;
                    IList<BoundarySegment> boundarySegments = null;
                    if(room.GetBoundarySegments(new SpatialElementBoundaryOptions()).Count != 0)
                    {
                        //撈選該房間之邊界資料
                        boundarySegments = room.GetBoundarySegments(new SpatialElementBoundaryOptions())[0];
                        wall_count = boundarySegments.Count;
                    }
                    else
                    {
                        wall_count = dex.excaRange.Count() - 1;
                    }
                    XYZ[] innerwall_points = null;
                    XYZ[] for_check_innerwall_points = null;
                    try
                    {

                        //取得連續壁內座標點
                        innerwall_points = new XYZ[wall_count + 1];
                        for_check_innerwall_points = new XYZ[wall_count + 1];//用來計算支撐中間樁之擺放界線
                        for (int i = 0; i < wall_count; i++)
                        {
                            //inner
                            innerwall_points[i] = boundarySegments[i].GetCurve().Tessellate()[0];
                            for_check_innerwall_points[i] = boundarySegments[i].GetCurve().Tessellate()[0];//用來計算支撐中間樁之擺放界線
                        }
                        //將最後一點設為第一點
                        innerwall_points[wall_count] = innerwall_points[0];
                        for_check_innerwall_points[wall_count] = for_check_innerwall_points[0];
                    }
                    catch
                    {
                        //取得連續壁內座標點
                        innerwall_points = new XYZ[wall_count + 1];
                        for_check_innerwall_points = new XYZ[wall_count + 1];//用來計算支撐中間樁之擺放界線
                        for (int i = 0; i != dex.excaRange.Count(); i++)
                        {
                            innerwall_points[i] = new XYZ(dex.excaRange[i].Item1 - xshift, dex.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;
                            for_check_innerwall_points[i] = new XYZ(dex.excaRange[i].Item1 - xshift, dex.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;
                        }
                    }

                    //取得所有XY數值
                    List<double> Xs = new List<double>();
                    List<double> Ys = new List<double>();
                    double[] slope = new double[wall_count];
                    double[] bias = new double[wall_count];
                    for (int i = 0; i < (innerwall_points.Count<XYZ>()); i++)
                    {
                        Xs.Add(innerwall_points[i].X);
                        Ys.Add(innerwall_points[i].Y);
                        if (i < slope.Count())
                        {
                            if (innerwall_points[i + 1].X - innerwall_points[i].X == 0)
                            {
                                slope[i] = 20172018;
                                bias[i] = innerwall_points[i + 1].X;
                            }
                            else
                            {
                                slope[i] = (innerwall_points[i + 1].Y - innerwall_points[i].Y) / (innerwall_points[i + 1].X - innerwall_points[i].X);
                                if (slope[i] == 0 || Math.Abs(slope[i]) < 0.0000001)
                                {
                                    bias[i] = innerwall_points[i + 1].Y;
                                }
                                else
                                {
                                    bias[i] = innerwall_points[i + 1].Y - slope[i] * innerwall_points[i + 1].X;
                                }
                            }
                        }
                    }

                    Transaction trans_2 = new Transaction(doc);
                    trans_2.Start("交易開始");
                    //開始建立中間樁
                    string column_bury = ((dex.column[0].Item6 - (dex.column[0].Item2 - dex.column[0].Item3)) * 1000).ToString();
                    double columns_dis = dex.centralCol[0] * 1000 / 304.8;//中間樁間距
                    if (ShiftData[0] != 0)
                    {
                        columns_dis = ShiftData[0] * 1000 / 304.8;
                    }
                    //選出中間樁族群。如過類型存在就直接用，不在就在根據匯入的資料建立一個類型
                    ICollection<FamilySymbol> columns_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                    foreach (FamilySymbol column_type in columns_familyinstance)
                    {
                        if (column_type.Name == "中間樁")
                        {
                            FamilySymbol eleFamSym = column_type;
                            bool exist = false;
                            foreach(FamilySymbol columnTypeCheck in columns_familyinstance)
                            {
                                if(columnTypeCheck.Name == "中間樁-" + dex.column[0].Item5)
                                {
                                    exist = true;
                                    eleFamSym = columnTypeCheck;
                                }
                            }
                            if (!exist) { eleFamSym = column_type.Duplicate("中間樁-" + dex.column[0].Item5) as FamilySymbol; }
                            string columnType = dex.column[0].Item5.Split('H')[1];
                            double b = 0, h = 0, s = 0, t = 0;
                            double.TryParse(columnType.Split('x')[0], out b);
                            double.TryParse(columnType.Split('x')[1], out h);
                            double.TryParse(columnType.Split('x')[2], out s);
                            double.TryParse(columnType.Split('x')[3], out t);
                            eleFamSym.LookupParameter("h").Set(b / 304.8);
                            eleFamSym.LookupParameter("b").Set(h / 304.8);
                            eleFamSym.LookupParameter("s").Set(s / 304.8);
                            eleFamSym.LookupParameter("t").Set(t / 304.8);
                            for (int j = 0; j < (Math.Abs(Ys.Max() - Ys.Min()) / columns_dis - 1) - ShiftData[2]; j++)
                            {

                                for (int i = 0; i < (Math.Abs(Xs.Max() - Xs.Min()) / columns_dis - 1) - ShiftData[1]; i++)
                                {

                                    XYZ column_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                                    if (IsInPolygon(column_location, innerwall_points) == true)
                                    {
                                        FamilyInstance column_instance = doc.Create.NewFamilyInstance(column_location, eleFamSym, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Column);
                                        column_instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).SetValueString("0");//給定中間樁長度
                                        column_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(String.Format("斷面{0}:中間樁", dex.section));
                                        column_instance.get_Parameter(BuiltInParameter.DOOR_NUMBER).Set((columns_dis / 1000 * 304.8).ToString());
                                        column_instance.LookupParameter("樁深埋入深度").SetValueString((dex.column[0].Item3 * 1000).ToString());
                                        column_instance.LookupParameter("型鋼埋入深度").SetValueString(column_bury);
                                        column_instance.LookupParameter("樁徑").SetValueString((dex.column[0].Item4 * 1000 / 2).ToString());

                                        // 中間柱方向 : 0 = 工, other = H  
                                        if (dex.column[0].Item7 != 0)
                                        {
                                            Line axis = Line.CreateBound(column_location, column_location + XYZ.BasisZ);
                                            ElementTransformUtils.RotateElement(doc, column_instance.Id, axis, Math.PI / 2);

                                        }


                                    }
                                }
                            }
                        }
                    }
                    trans_2.Commit();

                }
                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); break; }
            }

            TaskDialog.Show("done", "中間樁建置完畢");
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
