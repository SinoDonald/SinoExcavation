using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class sheet_pile_NoCAD : IExternalEventHandler
    {
        //讀取深開挖資料xlsx檔案
        public IList<string> files_path
        {
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
            Document doc = app.ActiveUIDocument.Document;

            foreach (string file_path in files_path)
            {
                //讀取資料
                ExReader sheet = new ExReader();
                sheet.SetData(file_path, 1);
                try
                {
                    sheet.PassWallData();
                    sheet.CloseEx();
                }
                catch (Exception e) { sheet.CloseEx(); TaskDialog.Show("Error", e.Message + e.StackTrace); }

                Transaction tran = new Transaction(doc);
                SubTransaction subtran = new SubTransaction(doc);

                tran.Start("放鋼板樁");

                //開挖各階之深度輸入
                List<double> height = new List<double>();
                foreach (var data in sheet.excaLevel)
                    height.Add(data.Item2 * -1);

                //偏移量
                double xshift = xy_shift[0];
                double yshift = xy_shift[1];

                //建立開挖階數            
                try
                {
                    Level[] levlist = new Level[height.Count()];
                    for (int i = 0; i != height.Count(); i++)
                    {
                        levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                        levlist[i].Name = String.Format("斷面{0}-開挖階數" + (i + 1).ToString(), sheet.section);
                    }

                }
                catch { }
                Level wall_level = Level.Create(doc, sheet.wall_high * 1000 * -1 / 304.8);
                try { wall_level.Name = String.Format("斷面{0}-擋土壁深度", sheet.section); } catch { }


                //須回到原點
                XYZ[] points = new XYZ[sheet.excaRange.Count()];

                //依照xlsx內給定的xy座標建立開挖範圍
                //讀取座標建立points
                for (int i = 0; i != sheet.excaRange.Count(); i++)
                    points[i] = new XYZ(sheet.excaRange[i].Item1 - xshift, sheet.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;

                Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().First();
                FamilySymbol sheet_pile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "鋼板樁_v2").ToList().First();
                sheet_pile.Activate();
                sheet_pile.LookupParameter("高度").SetValueString((sheet.wall_high * 1000).ToString());

                FamilyInstance instance1 = doc.Create.NewFamilyInstance(XYZ.Zero, sheet_pile, wall_level, StructuralType.Brace);
                double B = double.Parse(instance1.LookupParameter("B").AsValueString());
                double h = double.Parse(instance1.LookupParameter("h").AsValueString());
                double t = double.Parse(instance1.LookupParameter("t").AsValueString());
                double distance = (B + 2 * t * Math.Tan(15 * Math.PI / 180)) / 304.8;
                doc.Delete(instance1.Id);

                IList<Curve> wall_profileloops = new List<Curve>();
                
                for (int i = 0; i < points.Count() - 1; i++)
                {
                    subtran.Start();
                    Line edge_line = Line.CreateBound(points[i], points[i + 1]);

                    wall_profileloops.Add(edge_line);

                    double slope = edge_line.Direction.Y / edge_line.Direction.X;
                    int pile_num = (int)(edge_line.Length / distance) + 1;

                    List<ElementId> pre_nor_rotate_list = new List<ElementId>();
                    List<XYZ> pre_nor_rotate_point_list = new List<XYZ>();
                    List<ElementId> pre_rotate_list = new List<ElementId>();
                    List<XYZ> pre_rotate_point_list = new List<XYZ>();

                    //check line
                    int rotate_condition = 1;
                    if (i != 0)
                    {
                        Line check_line = Line.CreateBound(points[i - 1], points[i]);
                        if (slope == (check_line.Direction.Y / check_line.Direction.X))
                        {
                            if (((int)(check_line.Length / distance) + 1) % 2 == 1)
                            {
                                rotate_condition = 0;
                            }
                            else
                            {
                                rotate_condition = 1;
                            }
                        }
                    }

                    int dir_check = 0;
                    //如果往左，要先把元件往左邊移一個
                    if (edge_line.Direction.X < 0)
                    {
                        dir_check = 1;
                    }
                    subtran.Commit();
                    XYZ nomal_vector = XYZ.BasisZ.CrossProduct(edge_line.Direction);
                    subtran.Start();
                    for (int j = 0; j < pile_num; j++)
                    {
                        XYZ put_point = points[i] + (j + dir_check) * distance * edge_line.Direction - (h / 304.8) * nomal_vector;
                        FamilyInstance instance2 = doc.Create.NewFamilyInstance(put_point, sheet_pile, level, StructuralType.NonStructural);

                        ElementTransformUtils.RotateElement(doc, instance2.Id, Line.CreateBound(put_point, (put_point + XYZ.BasisZ)), Math.Atan(slope));
                        if (j % 2 == rotate_condition)
                        {
                            pre_rotate_list.Add(instance2.Id);
                            pre_rotate_point_list.Add(put_point);
                        }
                        //X方向為負數，會建置在內側
                        if (edge_line.Direction.X < 0)
                        {
                            pre_nor_rotate_list.Add(instance2.Id);
                            pre_nor_rotate_point_list.Add(put_point);
                        }
                        instance2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(j.ToString());
                    }
                    subtran.Commit();

                    //rotate elements
                    subtran.Start();
                    
                    for (int j = 0; j < pre_rotate_list.Count(); j++)
                    {
                        XYZ rot_sta;
                        if (edge_line.Direction.X >= 0)
                        {
                            rot_sta = pre_rotate_point_list[j] + (B / 2 / 304.8) * edge_line.Direction;// - (h / 304.8) * nomal_vector
                        }
                        else
                        {
                            rot_sta = pre_rotate_point_list[j] - (B / 2 / 304.8) * edge_line.Direction;// + (h / 304.8) * nomal_vector
                        }
                        XYZ rot_end = rot_sta + XYZ.BasisZ;
                        ElementTransformUtils.RotateElement(doc, pre_rotate_list[j], Line.CreateBound(rot_sta, rot_end), Math.PI);
                    }

                    subtran.Commit();

                    //X向為負數，整條元件要對座標線再旋轉一次
                    subtran.Start();
                    if (pre_nor_rotate_list.Count != 0)
                    {
                        for (int k = 0; k < pre_nor_rotate_list.Count; k++)
                        {
                            XYZ rot_sta;
                            rot_sta = pre_nor_rotate_point_list[k] - (B / 2 / 304.8) * edge_line.Direction;
                            XYZ rot_end = rot_sta + XYZ.BasisZ;
                            ElementTransformUtils.RotateElement(doc, pre_nor_rotate_list[k], Line.CreateBound(rot_sta, rot_end), Math.PI);
                        }
                    }
                    subtran.Commit();
                }

                tran.Commit();
            }

        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
