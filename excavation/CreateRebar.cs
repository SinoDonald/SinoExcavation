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

    class CreateRebar : IExternalEventHandler
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
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            foreach(string file_path in files_path)
            {
                try
                {
                    //讀取資料
                    ExReader dex = new ExReader();
                    dex.SetData(file_path, 1); //先假設在第一頁
                    try
                    {
                        dex.PassRebarData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.ToString()); }

                    // 要不要合體!!!!!!
                    bool combine_or_not = true;


                    //偏移量
                    //double xshift = xy_shift[0];
                    //double yshift = xy_shift[1];

                    //建立平面圖
                    Transaction trans0 = new Transaction(doc);
                    trans0.Start("Creat View");
                    ViewFamilyType viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().Where(x => x.Name == "樓板平面圖").First();
                    double elevation = 0;
                    Level LL = Level.Create(doc, elevation);
                    ViewPlan viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, LL.Id);
                    LL.Name = "放置鋼筋籠平面(elev. = 0)";
                    trans0.Commit();

                    //更改當前視圖
                    uidoc.ActiveView = viewPlan;

                    //建立牆壁 母單元
                    Transaction trans = new Transaction(doc);
                    trans.Start("Creat Wall");
                    //牆壁高度
                    Level wall_level = Level.Create(doc, dex.wall_high * 1000 * -1 / 304.8);

                    WallType wallType = null;
                    ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
                    //檢查擋土壁
                    if (walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (dex.wall_width * 1000).ToString() + "mm").ToList().Count != 0)
                        wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (dex.wall_width * 1000).ToString() + "mm").ToList().First();
                    //建立擋土壁新類型
                    if (wallType == null)
                    {
                        wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁").ToList().First();
                        WallType new_wallFamSym = wallType.Duplicate("連續壁" + dex.wall_width * 1000 + "mm") as WallType;
                        CompoundStructure ly = new_wallFamSym.GetCompoundStructure();
                        ly.SetLayerWidth(0, dex.wall_width * 1000 / 304.8);
                        new_wallFamSym.SetCompoundStructure(ly);
                        wallType = new_wallFamSym;
                    }
                    //牆壁位置（之後改成給使用者在平面圖點選）
                    double wall_length = (dex.f_length + dex.f_connector * 2) * 1000 / 304.8;
    
                    TaskDialog.Show("Test", "請點選欲放置鋼筋籠位置");
                    
                    XYZ start = uidoc.Selection.PickPoint();
                    XYZ end;
                    if (combine_or_not==true)
                    {
                        end = new XYZ(start.X + wall_length*2 - dex.f_connector * 1000/304.8, start.Y, start.Z);
                    }
                    else
                    {
                        end = new XYZ(start.X + wall_length, start.Y, start.Z);
                    }

                    //XYZ end = new XYZ(start.X + wall_length, start.Y, start.Z);
                    Line geomLine = Line.CreateBound(start, end);
                    Wall w = Wall.Create(doc, geomLine, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, true);
                    trans.Commit();
                    //TaskDialog.Show("TEST", "母單元牆壁建置完畢");


                    //開始建立鋼筋               
                    double width = dex.wall_width * 1000 / 304.8;
                    double pro = dex.protection_width * 10 / 304.8;
                    double ver_r_diameter = 0;
                    double ver_r_extra_diameter = 0;
                    double ver_e_diameter = 0;
                    double ver_e_extra_diameter = 0;
                    double hor_diameter = 0;
                    double shear_rebar_first = 0;
                    double index = pro;
                    double ver_t_diameter = 0;

                    Element host = document.GetElement(w.Id);
                    ICollection<RebarBarType> rebar_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(RebarBarType)).Cast<RebarBarType>().ToList();
                    RebarHookType hookType90 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 90 度").First();
                    RebarHookType hookType135 = new FilteredElementCollector(doc).OfClass(typeof(RebarHookType)).Cast<RebarHookType>().ToList().Where(x => x.Name == "鐙/箍 - 135 度").First();

                    
                    Transaction trans2 = new Transaction(doc);
                    trans2.Start("Create Main rebar!");

                    //垂直筋擋土側
                    string rebar_size_r = dex.vertical_r_rebar[0].Item3.Split('D')[1] + 'M';
                    
                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_r)
                        {
                            //RebarHostData rebarHostData = RebarHostData.GetRebarHostData(host);
                            //TaskDialog.Show("TEST", rebarHostData.IsValidHost().ToString());
                            ver_r_diameter = rebar_type.BarDiameter;
                            double S_v_r = dex.vertical_r_rebar[0].Item4 * 10 / 304.8;
                            XYZ normal = new XYZ(1, 0, 0);
                            int number = Convert.ToInt32(wall_length / S_v_r );
                            //TaskDialog.Show("TEST","垂直筋擋土側數量 = "+number.ToString());

                            for (int i = 0; i < number; i++)
                            {
                                XYZ origin = new XYZ(start.X + pro + i* S_v_r, start.Y + width/2 - pro, start.Z);
                                XYZ rebarLineEnd = new XYZ(origin.X, origin.Y, origin.Z - dex.wall_high * 1000 / 304.8);
                                IList<Curve> curves = new List<Curve>();
                                Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                curves.Add(rebarLine);
                                Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                                
                                //找出第一根剪力筋的位置: 1000mm 後的第一根主筋
                                index += S_v_r;
                                if (index > 1000/304.8)
                                {
                                    shear_rebar_first = index;
                                    index = -100000000;
                                    //TaskDialog.Show("TEST", "FIRST SHEAR REBAR = " + (shear_rebar_first*304.8).ToString());
                                }
                            }

                        }

                    }

                    //垂直筋開挖側
                    string rebar_size_e = dex.vertical_e_rebar[0].Item3.Split('D')[1] + 'M';

                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_e)
                        {
                            ver_e_diameter = rebar_type.BarDiameter;
                            double S_v_e = dex.vertical_e_rebar[0].Item4 * 10 / 304.8;
                            XYZ normal = new XYZ(1, 0, 0);
                            int number = Convert.ToInt32(wall_length / S_v_e);
                            //TaskDialog.Show("TEST", "垂直筋開挖側數量 = " + number.ToString());

                            for (int i = 0; i < number; i++)
                            {
                                XYZ origin = new XYZ(start.X + pro + i * S_v_e, start.Y - width/2 + pro, start.Z);
                                XYZ rebarLineEnd = new XYZ(origin.X, origin.Y, origin.Z - dex.wall_high * 1000 / 304.8);
                                IList<Curve> curves = new List<Curve>();
                                Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                curves.Add(rebarLine);
                                Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);

                            }
                        }
                    }

                    trans2.Commit();
                    //TaskDialog.Show("TEST", "主筋建置完畢");

                    //水平筋
                    Transaction trans3 = new Transaction(doc);
                    trans3.Start("Create H rebar!");

                    string rebar_size_h = dex.horizontal_rebar[0].Item3.Split('D')[1] + 'M';

                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_h)
                        {
                            hor_diameter = rebar_type.BarDiameter;
                            double S_h = dex.horizontal_rebar[0].Item4 * 10 / 304.8;
                            
                            XYZ normal = new XYZ(0, 0, 1);
                            int number = Convert.ToInt32(dex.wall_high * 1000 / 304.8 / S_h);
                            //TaskDialog.Show("TEST", "水平筋數量 = " + number.ToString());

                            //起點偏移量 (主筋半徑+副筋半徑 除以2)
                            double n = (ver_r_diameter + hor_diameter) / 2;

                            for (int i = 0; i < number; i++)
                            {
                                //擋土側
                                XYZ origin = new XYZ(start.X, start.Y + width/2 - pro - n, start.Z - i*S_h);
                                XYZ rebarLineEnd = new XYZ(origin.X + wall_length, origin.Y, origin.Z);
                                IList<Curve> curves = new List<Curve>();
                                Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                curves.Add(rebarLine);
                                Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                                
                                //開挖側
                                XYZ origin2 = new XYZ(start.X, start.Y - width / 2 + pro + n, start.Z - i*S_h);
                                XYZ rebarLineEnd2 = new XYZ(origin2.X + wall_length, origin2.Y, origin2.Z);
                                IList<Curve> curves2 = new List<Curve>();
                                Line rebarLine2 = Line.CreateBound(origin2, rebarLineEnd2);
                                
                                curves2.Add(rebarLine2);
                                Rebar re2 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves2, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                            }
                        }

                        //主筋擋土側加筋
                        for (int j = 0; j < dex.vertical_r_rebar.Count(); j++)
                        {
                            string rebar_extra_r = dex.vertical_r_rebar[j].Item5;

                            string extra_r = (rebar_extra_r == "X") ? rebar_extra_r : rebar_extra_r.Split('D')[1] + "M";
                            if (rebar_type.Name == extra_r)
                            {

                                ver_r_extra_diameter = rebar_type.BarDiameter;
                                double s = dex.vertical_r_rebar[j].Item6 * 10 / 304.8;
                                double start_depth = dex.vertical_r_rebar[j].Item1;
                                double end_depth = dex.vertical_r_rebar[j].Item2;
                                //TaskDialog.Show("TEST", rebar_type.Name + " " + start_depth.ToString() + " " + end_depth.ToString());

                                XYZ normal = new XYZ(1, 0, 0);
                                int number = Convert.ToInt32(wall_length / s);
                                //TaskDialog.Show("TEST", "Number of rebar = " + number.ToString());

                                //起點偏移量 (主筋半徑+主筋加筋半徑+副筋直徑)
                                double n = (ver_r_diameter + ver_r_extra_diameter) / 2 + hor_diameter;

                                for (int i = 0; i < number; i++)
                                {
                                    XYZ origin = new XYZ(start.X +　pro +i*s, start.Y + width/2 - pro - n, start.Z -start_depth);
                                    XYZ rebarLineEnd = new XYZ(origin.X, origin.Y, origin.Z - end_depth);
                                    IList<Curve> curves = new List<Curve>();
                                    Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                    curves.Add(rebarLine);
                                    Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);

                                }
                            }
                        }

                        //主筋開挖側加筋
                        for (int j = 0; j < dex.vertical_e_rebar.Count(); j++)
                        {
                            string rebar_extra_e= dex.vertical_e_rebar[j].Item5;

                            string extra_e = (rebar_extra_e == "X") ? rebar_extra_e : rebar_extra_e.Split('D')[1] + "M";
                            if (rebar_type.Name == extra_e)
                            {

                                ver_e_extra_diameter = rebar_type.BarDiameter;
                                double s = dex.vertical_e_rebar[j].Item6 * 10 / 304.8;
                                double start_depth = dex.vertical_e_rebar[j].Item1;
                                double end_depth = dex.vertical_e_rebar[j].Item2;
                                //TaskDialog.Show("TEST", rebar_type.Name + " " + start_depth.ToString() + " " + end_depth.ToString());

                                XYZ normal = new XYZ(1, 0, 0);
                                int number = Convert.ToInt32(wall_length / s);
                                //TaskDialog.Show("TEST", "Number of rebar = " + number.ToString());

                                //起點偏移量 (主筋半徑+主筋加筋半徑+副筋直徑)
                                double n = (ver_e_diameter + ver_e_extra_diameter) / 2 + hor_diameter;

                                for (int i = 0; i < number; i++)
                                {
                                    XYZ origin = new XYZ(start.X + pro + i * s, start.Y - width/2 + pro + n, start.Z -start_depth);
                                    XYZ rebarLineEnd = new XYZ(origin.X, origin.Y, origin.Z - end_depth);
                                    IList<Curve> curves = new List<Curve>();
                                    Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                    curves.Add(rebarLine);
                                    Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);

                                }
                            }
                        }
                    }

                    trans3.Commit();
                    //TaskDialog.Show("TEST", "水平筋建置完畢");

                    //剪力筋
                    Transaction trans4 = new Transaction(doc);
                    trans4.Start("Create shear rebar!");

                    for (int j = 0; j < dex.shear_rebar_depth.Count(); j++)
                    { 
                        double start_depth = dex.shear_rebar_depth[j].Item1 * 1000 / 304.8;
                        double end_depth = dex.shear_rebar_depth[j].Item2 * 1000 / 304.8;

                        string h_rebar = dex.shear_rebar[j].Item2.Split('D')[1] + "M";
                        string v_rebar = dex.shear_rebar[j].Item5.Split('D')[1] + "M";
                        double h_space = dex.shear_rebar[j].Item3 * 10 / 304.8;
                        double v_space = dex.shear_rebar[j].Item6 * 10 / 304.8;

                        foreach (RebarBarType rebar_type in rebar_familyinstance)
                        {
                            if (rebar_type.Name == h_rebar)
                            {
                                //垂直段數量
                                XYZ normal = new XYZ(0,0,1);
                                int number_v = Convert.ToInt32((end_depth - start_depth) / v_space);
                                //TaskDialog.Show("TEST", "剪力筋垂直方向數量 = " + number_v.ToString());

                                for (int k = 0; k < number_v; k++)
                                {
                                    double h_diameter = rebar_type.BarDiameter;
                                    //水平段數量
                                    int number_h = Convert.ToInt32((dex.f_length * 1000 / 304.8) / h_space);
                                    //TaskDialog.Show("TEST", "剪力筋水平方向數量 = " + number_h.ToString());
                                    
                                    for (int i = 0; i < number_h; i++)
                                    {
                                        XYZ origin = new XYZ(start.X + shear_rebar_first + (ver_r_diameter + h_diameter) / 2 + i * h_space, start.Y + width/2 - pro + ver_e_diameter / 2 + h_diameter, start.Z -start_depth - k * v_space);
                                        XYZ rebarLineEnd = new XYZ(origin.X, - width/2 + pro - ver_e_diameter / 2 - h_diameter, origin.Z);

                                        IList<Curve> curves = new List<Curve>();
                                        Line rebarLine = Line.CreateBound(origin, rebarLineEnd);
                                        RebarHookType startHook = hookType90;
                                        RebarHookType endHook = hookType135;
                                        //彎鉤90度和135度上下左右都要錯開
                                        if (k % 2 == 0)
                                        {
                                            startHook = (i % 2 == 0) ? hookType90 : hookType135;
                                            endHook = (i % 2 == 0) ? hookType135 : hookType90;
                                        }
                                        else
                                        {
                                            startHook = (i % 2 == 0) ? hookType135 : hookType90;
                                            endHook = (i % 2 == 0) ? hookType90 : hookType135;
                                        }                                  
                                        curves.Add(rebarLine);
                                        Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.StirrupTie, rebar_type, startHook, endHook, host, normal, curves, RebarHookOrientation.Right, RebarHookOrientation.Right, true, true);
                                    }
                                }
                            }
                        }
                    }
                    trans4.Commit();
                    //TaskDialog.Show("TEST", "剪力筋建置完畢");
                    

                    //公單元     
                    //建立牆壁
                    Transaction trans5 = new Transaction(doc);
                    trans5.Start("Creat Wall");

                    double wall_length_m = (dex.m_length) * 1000 / 304.8;
                    XYZ start2;
                    XYZ end2;
                    if (combine_or_not == true)
                    {
                        start2 = new XYZ(start.X + wall_length - dex.f_connector * 1000/304.8 , start.Y, start.Z);
                    }
                    else
                    {
                        start2 = new XYZ(start.X + wall_length + 500 / 304.8, start.Y, start.Z);
                        end2 = new XYZ(start2.X + wall_length_m, start2.Y, start2.Z);

                        Line geomLine2 = Line.CreateBound(start2, end2);
                        Wall w2 = Wall.Create(doc, geomLine2, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, true);
                        host = document.GetElement(w2.Id);
                    }
                    //start2 = new XYZ(start.X + wall_length + 500 / 304.8, start.Y, start.Z);
                    //end2 = new XYZ(start2.X + wall_length_m, start2.Y, start2.Z);    
                    //Line geomLine2 = Line.CreateBound(start2, end2);
                    //Wall w2 = Wall.Create(doc, geomLine2, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, true);

                    trans5.Commit();
                    //TaskDialog.Show("TEST", "公單元牆壁建置完畢");

                    //host = document.GetElement(w2.Id);

                    //開始建鋼筋
                    //水平筋
                    Transaction trans6 = new Transaction(doc);
                    trans6.Start("Create Main rebar!");
                    //string rebar_size_h = dex.horizontal_rebar[0].Item3.Split('D')[1] + 'M';

                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_h)
                        {
                            hor_diameter = rebar_type.BarDiameter;
                            XYZ normal = new XYZ(0, 0, 1);
                            double S_h = dex.horizontal_rebar[0].Item4 * 10 / 304.8;
                            int number = Convert.ToInt32(dex.wall_high * 1000 / 304.8 / S_h);

                            for (int i = 0; i < number; i++)
                            {
                                //擋土側
                                XYZ origin = new XYZ(start2.X + (dex.f_connector - dex.m_connector) * 1000 / 304.8, start2.Y + width / 2 - dex.m_connector2 * 1000 / 304.8, start2.Z - i * S_h);
                                XYZ point1 = new XYZ(origin.X + dex.m_connector * 1000 / 304.8, origin.Y, origin.Z);
                                XYZ point2 = new XYZ(point1.X + dex.m_connector2 * 1000 / 304.8, start2.Y + width / 2 - pro, origin.Z);
                                XYZ point3 = new XYZ(point2.X + dex.f_length * 1000 / 304.8 - 2*(dex.m_connector2 * 1000 / 304.8), point2.Y, origin.Z);
                                XYZ point4 = new XYZ(point3.X + dex.m_connector2 * 1000 / 304.8, point1.Y, origin.Z);
                                XYZ endpoint = new XYZ(point4.X + dex.m_connector * 1000 / 304.8, origin.Y, origin.Z);
                            
                                IList<Curve> curves = new List<Curve>();
                            
                                Line Line1 = Line.CreateBound(origin, point1);
                                Line Line2 = Line.CreateBound(point1, point2);
                                Line Line3 = Line.CreateBound(point2, point3);
                                Line Line4 = Line.CreateBound(point3, point4);
                                Line Line5 = Line.CreateBound(point4, endpoint);
                            
                                curves.Add(Line1);
                                curves.Add(Line2);
                                curves.Add(Line3);
                                curves.Add(Line4);
                                curves.Add(Line5);

                                Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);

                                //開挖側
                                origin = new XYZ(start2.X + (dex.f_connector - dex.m_connector) * 1000 / 304.8, start2.Y - width / 2 + dex.m_connector2 * 1000 / 304.8, start2.Z - i * S_h);
                                point1 = new XYZ(origin.X + dex.m_connector * 1000 / 304.8, origin.Y, origin.Z);
                                point2 = new XYZ(point1.X + dex.m_connector2 * 1000 / 304.8, start2.Y - width / 2 + pro, origin.Z);
                                point3 = new XYZ(point2.X + dex.f_length * 1000 / 304.8 - 2 * (dex.m_connector2 * 1000 / 304.8), point2.Y, origin.Z);
                                point4 = new XYZ(point3.X + dex.m_connector2 * 1000 / 304.8, point1.Y, origin.Z);
                                endpoint = new XYZ(point4.X + dex.m_connector * 1000 / 304.8, origin.Y, origin.Z);

                                IList<Curve> curves2 = new List<Curve>();

                                Line1 = Line.CreateBound(origin, point1);
                                Line2 = Line.CreateBound(point1, point2);
                                Line3 = Line.CreateBound(point2, point3);
                                Line4 = Line.CreateBound(point3, point4);
                                Line5 = Line.CreateBound(point4, endpoint);

                                curves2.Add(Line1);
                                curves2.Add(Line2);
                                curves2.Add(Line3);
                                curves2.Add(Line4);
                                curves2.Add(Line5);

                                Rebar re2 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves2, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);

                            }
                        }
                    }
                    trans6.Commit();
                    //TaskDialog.Show("TEST","水平筋建立完畢");

                    //垂直筋
                    Transaction trans7 = new Transaction(doc);
                    trans7.Start("Create Horizontal rebar!");
                    //擋土側
                    //string rebar_size_r = dex.vertical_r_rebar[0].Item3.Split('D')[1] + 'M';

                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_r)
                        {
                            ver_r_diameter = rebar_type.BarDiameter;
                            double S_v_r = dex.vertical_r_rebar[0].Item4 * 10 / 304.8;
                            XYZ normal = new XYZ(1, 0, 0);
                            int number = Convert.ToInt32(((dex.f_length - dex.m_connector2 * 2)*1000/304.8)/ S_v_r);
                            //TaskDialog.Show("TEST", "垂直筋擋土側數量 = " + (number+1).ToString());

                            for (int i = 0; i < number; i++)
                            {
                                XYZ origin = new XYZ(start2.X + dex.f_connector*1000/304.8 + dex.m_connector2*1000/304.8 + i*S_v_r, start2.Y + width/2 - pro + (ver_r_diameter + hor_diameter)/2, start2.Z );
                                XYZ rebarLineEnd = new XYZ(origin.X, origin.Y, origin.Z - dex.wall_high * 1000 / 304.8);
                                IList<Curve> curves = new List<Curve>();
                                Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                curves.Add(rebarLine);
                                Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                            }
                            //除不盡最後補一根
                            XYZ final = new XYZ(start2.X + wall_length_m - dex.f_connector * 1000 / 304.8 - dex.m_connector2 * 1000 / 304.8, start2.Y + width / 2 - pro + (ver_r_diameter + hor_diameter) / 2, start2.Z);
                            XYZ finalEnd = new XYZ(final.X, final.Y, final.Z - dex.wall_high * 1000 / 304.8); ;

                            IList<Curve> curve = new List<Curve>();
                            Line line = Line.CreateBound(final, finalEnd);

                            curve.Add(line);
                            Rebar re2 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curve, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                        }
                    }

                    //開挖側
                    //string rebar_size_e = dex.vertical_e_rebar[0].Item3.Split('D')[1] + 'M';

                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_r)
                        {
                            ver_e_diameter = rebar_type.BarDiameter;
                            double S_v_e = dex.vertical_e_rebar[0].Item4 * 10 / 304.8;
                            XYZ normal = new XYZ(1, 0, 0);
                            int number = Convert.ToInt32(((dex.f_length - dex.m_connector2 * 2) * 1000 / 304.8) / S_v_e);
                            //TaskDialog.Show("TEST", "垂直筋開挖側數量 = " + (number + 1).ToString());

                            for (int i = 0; i < number; i++)
                            {
                                XYZ origin = new XYZ(start2.X + dex.f_connector * 1000 / 304.8 + dex.m_connector2 * 1000 / 304.8 + i * S_v_e, start2.Y - width / 2 + pro - (ver_r_diameter + hor_diameter) / 2, start2.Z);
                                XYZ rebarLineEnd = new XYZ(origin.X, origin.Y, origin.Z - dex.wall_high * 1000 / 304.8);
                                IList<Curve> curves = new List<Curve>();
                                Line rebarLine = Line.CreateBound(origin, rebarLineEnd);

                                curves.Add(rebarLine);
                                Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                            }
                            //除不盡最後補一根
                            XYZ final = new XYZ(start2.X + wall_length_m - dex.f_connector * 1000 / 304.8 - dex.m_connector2 * 1000 / 304.8, start2.Y - width / 2 + pro - (ver_r_diameter + hor_diameter) / 2, start2.Z);
                            XYZ finalEnd = new XYZ(final.X, final.Y, final.Z - dex.wall_high * 1000 / 304.8); ;

                            IList<Curve> curve = new List<Curve>();
                            Line line = Line.CreateBound(final, finalEnd);

                            curve.Add(line);
                            Rebar re2 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curve, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                        }
                    }

                    //榫尾兩端
                    string rebar_size_t = dex.m_main.Split('D')[1] + 'M';

                    foreach (RebarBarType rebar_type in rebar_familyinstance)
                    {
                        if (rebar_type.Name == rebar_size_t)
                        {
                            ver_t_diameter = rebar_type.BarDiameter;
                            double S_v_t = dex.m_space * 10 / 304.8;
                            XYZ normal = new XYZ(1, 0, 0);
                            int number = Convert.ToInt32(((dex.m_connector) * 1000 / 304.8) / S_v_t);
                            //TaskDialog.Show("TEST", "榫尾數量 = " + (number + 1).ToString());
                            for (int i = 0; i < number+1; i++)
                            {
                                XYZ origin1 = new XYZ(start2.X + (dex.f_connector - dex.m_connector) * 1000 / 304.8 + i*S_v_t, start2.Y + width / 2 - dex.m_connector2*1000/304.8 + (ver_r_diameter + hor_diameter) / 2, start2.Z);
                                XYZ rebarLineEnd1 = new XYZ(origin1.X, origin1.Y, origin1.Z - dex.wall_high * 1000 / 304.8);
                                XYZ origin2 = new XYZ(start2.X + (dex.m_length - dex.f_connector + dex.m_connector) * 1000 / 304.8 - i*S_v_t, origin1.Y, origin1.Z);
                                XYZ rebarLineEnd2 = new XYZ(origin2.X, origin2.Y, origin2.Z - dex.wall_high * 1000 / 304.8);
                                XYZ origin3 = new XYZ(origin1.X, start2.Y - width / 2 + dex.m_connector2 * 1000 / 304.8 - (ver_r_diameter + hor_diameter) / 2, start2.Z);
                                XYZ rebarLineEnd3 = new XYZ(origin3.X, origin3.Y, origin3.Z - dex.wall_high * 1000 / 304.8);
                                XYZ origin4 = new XYZ(origin2.X, origin3.Y, start2.Z);
                                XYZ rebarLineEnd4 = new XYZ(origin4.X, origin4.Y, origin4.Z - dex.wall_high * 1000 / 304.8);

                                IList<Curve> curves1 = new List<Curve>();
                                IList<Curve> curves2 = new List<Curve>();
                                IList<Curve> curves3 = new List<Curve>();
                                IList<Curve> curves4 = new List<Curve>();

                                Line rebarLine1 = Line.CreateBound(origin1, rebarLineEnd1);
                                Line rebarLine2 = Line.CreateBound(origin2, rebarLineEnd2);
                                Line rebarLine3 = Line.CreateBound(origin3, rebarLineEnd3);
                                Line rebarLine4 = Line.CreateBound(origin4, rebarLineEnd4);

                                curves1.Add(rebarLine1);
                                curves2.Add(rebarLine2);
                                curves3.Add(rebarLine3);
                                curves4.Add(rebarLine4);

                                Rebar re1 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves1, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                                Rebar re2 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves2, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                                Rebar re3 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves3, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);
                                Rebar re4 = Rebar.CreateFromCurves(doc, RebarStyle.Standard, rebar_type, null, null, host, normal, curves4, RebarHookOrientation.Left, RebarHookOrientation.Left, true, true);

                            }
                        }
                    }

                    trans7.Commit();
                    //TaskDialog.Show("TEST", "垂直筋建立完畢");

                    //剪力筋
                    Transaction trans8 = new Transaction(doc);
                    trans8.Start("Create shear rebar");

                    for (int j = 0; j < dex.shear_rebar_depth.Count(); j++)
                    {
                        double start_depth = dex.shear_rebar_depth[j].Item1 * 1000 / 304.8;
                        double end_depth = dex.shear_rebar_depth[j].Item2 * 1000 / 304.8;

                        string h_rebar = dex.shear_rebar[j].Item2.Split('D')[1] + "M";
                        string v_rebar = dex.shear_rebar[j].Item5.Split('D')[1] + "M";
                        double h_space = dex.shear_rebar[j].Item3 * 10 / 304.8;
                        double v_space = dex.shear_rebar[j].Item6 * 10 / 304.8;
                        double t_space = dex.m_space * 10 / 304.8;

                        foreach (RebarBarType rebar_type in rebar_familyinstance)
                        {
                            if (rebar_type.Name == h_rebar)
                            {
                                //垂直段數量
                                XYZ normal = new XYZ(0, 0, 1);
                                int number_v = Convert.ToInt32((end_depth - start_depth) / v_space); 

                                RebarHookType startHook = hookType90;
                                RebarHookType endHook = hookType135;
                                
                                for (int k = 0; k < number_v; k++)
                                {
                                    double h_diameter = rebar_type.BarDiameter;
                                    //水平段數量
                                    double middle = (dex.f_length - dex.m_connector2 * 2) * 1000 / 304.8;
                                    int number_h = Convert.ToInt32( middle / h_space);
                                    //int number = ((number_h + 1) * h_space < middle) ? number_h + 1 : number_h;
                                    
                                    for (int i = 0; i < number_h; i++)
                                    {
                                        XYZ origin = new XYZ(start2.X + dex.f_connector * 1000 / 304.8 + dex.m_connector2 * 1000 / 304.8 + (ver_r_diameter + h_diameter) / 2 + i * h_space, start2.Y + width/2 - pro + hor_diameter/2 + ver_r_diameter + h_diameter,  -start_depth - k * v_space);
                                        XYZ rebarLineEnd = new XYZ(origin.X, start2.Y - width/2 + pro - hor_diameter/2 - ver_e_diameter - h_diameter, origin.Z);
                                        IList<Curve> curves = new List<Curve>();
                                        Line rebarLine = Line.CreateBound(origin, rebarLineEnd);
                                        
                                        
                                        //彎鉤90度和135度上下左右都要錯開
                                        if (k % 2 == 0)
                                        {
                                            startHook = (i % 2 == 0) ? hookType90 : hookType135;
                                            endHook = (i % 2 == 0) ? hookType135 : hookType90;
                                        }
                                        else
                                        {
                                            startHook = (i % 2 == 0) ? hookType135 : hookType90;
                                            endHook = (i % 2 == 0) ? hookType90 : hookType135;
                                        }
                                        curves.Add(rebarLine);
                                        Rebar re = Rebar.CreateFromCurves(doc, RebarStyle.StirrupTie, rebar_type, startHook, endHook, host, normal, curves, RebarHookOrientation.Right, RebarHookOrientation.Right, true, true);
                                    }
                                    XYZ final = new XYZ(start2.X + wall_length_m - dex.f_connector * 1000 / 304.8 - dex.m_connector2 * 1000 / 304.8 + (ver_r_diameter + h_diameter) / 2, start2.Y + width / 2 - pro + hor_diameter / 2 + ver_r_diameter + h_diameter, -start_depth - k * v_space);
                                    XYZ finalEnd = new XYZ(final.X, start2.Y - width / 2 + pro - hor_diameter / 2 - ver_e_diameter - h_diameter, final.Z) ;

                                    IList<Curve> curve = new List<Curve>();
                                    Line line = Line.CreateBound(final, finalEnd);
                                    
                                    startHook = (startHook == hookType135) ? hookType90 : hookType135;
                                    endHook = (endHook == hookType135) ? hookType90 : hookType135;
                                    
                                    curve.Add(line);
                                    Rebar re2 = Rebar.CreateFromCurves(doc, RebarStyle.StirrupTie, rebar_type, startHook, endHook, host, normal, curve, RebarHookOrientation.Right, RebarHookOrientation.Right, true, true);

                                    int number_t = Convert.ToInt32(((dex.m_connector) * 1000 / 304.8) / t_space);
                                    for (int i = 0; i < number_t+1; i++)
                                    {
                                        XYZ origin1 = new XYZ(start2.X + (dex.f_connector - dex.m_connector) * 1000 / 304.8 + (ver_t_diameter + h_diameter) / 2 + i * t_space, start2.Y + width / 2 - dex.m_connector2 * 1000 / 304.8 + hor_diameter / 2 + ver_r_diameter + h_diameter, -start_depth - k * v_space);
                                        XYZ rebarLineEnd1 = new XYZ(origin1.X, start2.Y - width / 2 + dex.m_connector2 * 1000 / 304.8 - hor_diameter / 2 - ver_t_diameter - h_diameter, origin1.Z);


                                        XYZ origin2 = new XYZ(start2.X + (dex.m_length - dex.f_connector + dex.m_connector) * 1000 / 304.8 + (ver_t_diameter + h_diameter) / 2 - i * t_space, origin1.Y, origin1.Z);
                                        XYZ rebarLineEnd2 = new XYZ(origin2.X, rebarLineEnd1.Y, origin2.Z);

                                        IList<Curve> curves1 = new List<Curve>();
                                        Line rebarLine1 = Line.CreateBound(origin1, rebarLineEnd1);
                                        curves1.Add(rebarLine1);

                                        IList<Curve> curves2 = new List<Curve>();
                                        Line rebarLine2 = Line.CreateBound(origin2, rebarLineEnd2);
                                        curves2.Add(rebarLine2);

                                        if (k % 2 == 0)
                                        {
                                            startHook = (i % 2 == 0) ? hookType90 : hookType135;
                                            endHook = (i % 2 == 0) ? hookType135 : hookType90;
                                        }
                                        else
                                        {
                                            startHook = (i % 2 == 0) ? hookType135 : hookType90;
                                            endHook = (i % 2 == 0) ? hookType90 : hookType135;
                                        }
                                        Rebar re3 = Rebar.CreateFromCurves(doc, RebarStyle.StirrupTie, rebar_type, startHook, endHook, host, normal, curves1, RebarHookOrientation.Right, RebarHookOrientation.Right, true, true);
                                        Rebar re4 = Rebar.CreateFromCurves(doc, RebarStyle.StirrupTie, rebar_type, startHook, endHook, host, normal, curves2, RebarHookOrientation.Right, RebarHookOrientation.Right, true, true);

                                    }

                                }


                                
                            }
                        }
                    }

                    trans8.Commit();
                    //TaskDialog.Show("TEST","剪力筋建立完畢");

                }
                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); break; }
            }

            
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
        public void CopyPaste(ElementId id, double startDepth, double endDepth, double v_space , Document doc)
        {
            int number = Convert.ToInt32((endDepth - startDepth) / v_space);
            //TaskDialog.Show("TEST", "number of copy = " + number.ToString());
            Element element = doc.GetElement(id);
            for (int i = 1; i < number; i++)
            {
                XYZ loc = new XYZ(0, 0, -i*v_space);
                ElementTransformUtils.CopyElement(doc, id, loc);
            }
        }
    }
}

