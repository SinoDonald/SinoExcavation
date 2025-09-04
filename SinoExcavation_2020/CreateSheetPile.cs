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
        public OpenFileDialog openFileDialog
        {
            get;
            set;
        }

        public void Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            Transaction tran = new Transaction(doc);

            string dwg_file_name = "";
            Transform project_transform = null;
            GeometryElement geoLines = null;
            try
            {
                dwg_file_name = openFileDialog.FileName;
            }
            catch { }
            
            if (dwg_file_name.Contains(".dwg"))
            {
                tran.Start("CAD");
                //插入CAD
                Autodesk.Revit.DB.View view = doc.ActiveView;
                DWGImportOptions dWGImportOptions = new DWGImportOptions();
                dWGImportOptions.ColorMode = ImportColorMode.Preserved;
                dWGImportOptions.Placement = ImportPlacement.Origin;
                LinkLoadResult linkLoadResult = new LinkLoadResult();
                ImportInstance toz = ImportInstance.Create(doc, view, dwg_file_name, dWGImportOptions, out linkLoadResult);
                toz.Pinned = false;
                ElementTransformUtils.MoveElement(doc, toz.Id, new XYZ(xy_shift[0], xy_shift[1], 0));

                tran.Commit();

                //取得CAD
                project_transform = toz.GetTotalTransform();
                GeometryElement geometryElement = toz.get_Geometry(new Options());
                geoLines = (geometryElement.First() as GeometryInstance).SymbolGeometry;
            }

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
                List<Line> lines = new List<Line>();
                if (dwg_file_name.Contains(".dwg"))
                {
                    IList<String> target_section = new List<String>();
                    target_section.Add(sheet.section);
                    IList<XYZ> allXYZs = new List<XYZ>();
                    foreach (var v in geoLines)
                    {
                        allXYZs.Clear();
                        try//處理polyline
                        {
                            PolyLine pline = v as PolyLine;
                            GraphicsStyle check = doc.GetElement(pline.GraphicsStyleId) as GraphicsStyle;

                            //檢查是否為所要圖層
                            if (target_section.Contains(check.GraphicsStyleCategory.Name))
                            {

                                //撈取所有點位
                                foreach (XYZ p in pline.GetCoordinates())
                                {
                                    allXYZs.Add(project_transform.OfPoint(p - new XYZ(xshift, yshift, 0)));
                                }

                                bool Clockdirection = ClockwiseDirection(allXYZs);

                                //建置線段//
                                if (Clockdirection == true)//若為順時針需要倒轉
                                {
                                    allXYZs = allXYZs.Reverse().ToList();
                                }
                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);
                                    lines.Add(line);
                                }
                            }
                        }
                        catch (Exception) { }

                    }

                    //處理Arc順逆時針問題
                    foreach (var v in geoLines)
                    {
                        allXYZs.Clear();
                        try
                        {
                            //處理ARC
                            Arc pline = v as Arc;
                            //檢查是否為所要圖層
                            if (target_section.Contains((doc.GetElement(pline.GraphicsStyleId) as GraphicsStyle).GraphicsStyleCategory.Name))
                            {
                                //撈取所有點位
                                foreach (XYZ p in pline.Tessellate())
                                {
                                    allXYZs.Add(project_transform.OfPoint(p - new XYZ(xshift, yshift, 0)));
                                }

                                for (int i = 0; i < allXYZs.Count() - 1; i++)
                                {
                                    Line line = Line.CreateBound(allXYZs[i], allXYZs[i + 1]);
                                    lines.Add(line);
                                }
                            }
                        }
                        catch (Exception) { }
                    }
                }
                else
                {
                    //依照xlsx內給定的xy座標建立開挖範圍
                    //讀取座標建立points
                    for (int i = 0; i != sheet.excaRange.Count() - 1; i++)
                    {
                        XYZ point1 = new XYZ(sheet.excaRange[i].Item1 - xshift, sheet.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;
                        XYZ point2 = new XYZ(sheet.excaRange[i + 1].Item1 - xshift, sheet.excaRange[i + 1].Item2 - yshift, 0) * 1000 / 304.8;

                        lines.Add(Line.CreateBound(point1, point2));
                    }
                }

                Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().First();

                FamilySymbol sheet_pile = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "鋼板樁_v2").ToList().First();
                sheet_pile.Activate();

                sheet_pile.LookupParameter("高度").SetValueString((sheet.wall_high * 1000).ToString());

                FamilyInstance instance1 = doc.Create.NewFamilyInstance(XYZ.Zero, sheet_pile, wall_level, StructuralType.Brace);
                double B = double.Parse(instance1.LookupParameter("B").AsValueString());
                double h = double.Parse(instance1.LookupParameter("h").AsValueString());//h
                double t = double.Parse(instance1.LookupParameter("t").AsValueString());//t
                double distance = (B + 2 * t * Math.Tan(15 * Math.PI / 180)) / 304.8;

                doc.Delete(instance1.Id);

                IList<Curve> wall_profileloops = new List<Curve>();

                for (int i = 0; i < lines.Count(); i++)
                {
                    subtran.Start();
                    Line edge_line = lines[i];

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
                        Line check_line = lines[i - 1];
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
                        XYZ put_point = edge_line.GetEndPoint(0) + (j + dir_check) * distance * edge_line.Direction - (h / 304.8) * nomal_vector;
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

        private bool ClockwiseDirection(IList<XYZ> points)
        {
            //計算多邊形邊界線順時鐘/逆時鐘
            bool clock = true;
            int i, j, k;
            int count = 0;
            double z;
            int n = points.Count;
            for (i = 0; i < n; i++)
            {
                j = (i + 1) % n;
                k = (i + 2) % n;
                z = (points[j].X - points[i].X) * (points[k].Y - points[j].Y);
                z -= (points[j].Y - points[i].Y) * (points[k].X - points[j].X);
                if (z < 0)
                {
                    count--;
                }
                else if (z > 0)
                {
                    count++;
                }
            }
            if (count > 0)
            {
                clock = false;
            }

            return clock;

        }
    }
}
