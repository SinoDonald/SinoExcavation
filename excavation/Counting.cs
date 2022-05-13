using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class Counting : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        public string count_excel
        {
            get;
            set;
        }
        public string path
        {
            get;
            set;
        }
        public string new_name
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            try
            {
                Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
                UIDocument uidoc = new UIDocument(document);
                Document doc = uidoc.Document;
                Transaction t = new Transaction(doc);
                List<List<string>> total_mud = new List<List<string>>();
                foreach (string file_path in files_path)
                {
                    ExReader reader = new ExReader();
                    reader.SetData(file_path, 1);
                    reader.PassWallData();
                    reader.PassColumnData();
                    reader.PassSideData();
                    reader.CloseEx();
                    string check_str = reader.section;

                    //連續壁部分

                    Level S1 = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name == "斷面" + check_str + "-擋土壁深度").ToList().First();
                    List<Wall> wall_list = new FilteredElementCollector(doc).OfClass(typeof(Wall)).Cast<Wall>().Where(x => x.LevelId == S1.Id).ToList();

                    TaskDialog.Show("1", S1.Name);
                    TaskDialog.Show("1", wall_list.Count().ToString());
                    //找到長度
                    double total_l = 0;
                    foreach (Wall a in wall_list)
                    {
                        // * 304.8 / 1000
                        LocationCurve locationCurve = a.Location as LocationCurve;
                        double length = locationCurve.Curve.Length * 304.8 / 1000;
                        total_l += length;
                    }
                    total_l = Math.Round(total_l, 3);

                    //找到深度
                    //WALL_USER_HEIGHT_PARAM => 不連續高度
                    double depth = double.Parse(wall_list[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsValueString()) / 1000;


                    //中間樁部分

                    IList<FamilyInstance> mid_zhang = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Where(x => x.Name.Contains("中間樁")).Cast<FamilyInstance>().ToList();
                    FamilyInstance Column = mid_zhang.First();
                    IList<double> count_mid = new List<double>(); 
                    double col_depth = double.Parse(Column.LookupParameter("樁深埋入深度").AsValueString()) / 1000;
                    double ste_depth = (double.Parse(Column.LookupParameter("型鋼埋入深度").AsValueString()) / 1000);
                    string rad = (double.Parse(Column.LookupParameter("樁徑").AsValueString()) / 1000).ToString();
                    IList<FamilyInstance> temp = new List<FamilyInstance>();
                    foreach (FamilyInstance fami in mid_zhang)
                    {
                        Element check_level = doc.GetElement(fami.LevelId);


                        if (check_level.Name.Contains(check_str))
                        {
                            temp.Add(fami);
                        }
                    }
                    count_mid.Add(temp.Count);

                    //回撐部分
                    List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>> mid_back_t_list = new List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>>();
                    List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>> mid_back_s_list = new List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>>();
                    List<string> back_floor = new List<string>();
                    for (int lev = 0; lev != reader.supLevel.Count(); lev++)
                    {
                        try
                        {
                            int.Parse(reader.supLevel[lev].Item1.ToString());
                        }
                        catch { back_floor.Add(reader.supLevel[lev].Item1.ToString()); }
                    }
                    foreach (string floorname in back_floor)
                    {
                        List<FamilyInstance> side_steel = new List<FamilyInstance>(); //抓圍囹
                        List<FamilyInstance> support = new List<FamilyInstance>(); //抓支撐
                        List<FamilyInstance> test = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
                        foreach (FamilyInstance famtest in test)
                        {
                            try
                            {
                                string comment = famtest.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
                                if (comment != null && comment.Contains(check_str + "-" + floorname + "-圍囹"))
                                {
                                    side_steel.Add(famtest);
                                }
                                if (comment != null && comment.Contains(check_str + "-" + floorname + "-支撐"))
                                {
                                    support.Add(famtest);
                                }
                            }
                            catch { }
                        }

                        List<string> steel_class = new List<string>(); //圍囹的型鋼類型
                        List<string> support_class = new List<string>(); //支撐的型鋼類型

                        //載入類型
                        foreach (FamilyInstance FamIn in side_steel)
                        {
                            LocationCurve locu = FamIn.Location as LocationCurve;
                            steel_class.Add(FamIn.Name);
                        }
                        foreach (FamilyInstance FamIn in support)
                        {
                            LocationCurve locu = FamIn.Location as LocationCurve;
                            support_class.Add(FamIn.Name);
                        }

                        //篩選重複的類型
                        for (int a = 0; a < steel_class.Count(); a++)
                        {
                            for (int b = steel_class.Count() - 1; b > a; b--)
                            {
                                if (steel_class[a] == steel_class[b])
                                {
                                    steel_class.RemoveAt(b);
                                }
                            }
                        }
                        for (int a = 0; a < support_class.Count(); a++)
                        {
                            for (int b = support_class.Count() - 1; b > a; b--)
                            {
                                if (support_class[a] == support_class[b])
                                {
                                    support_class.RemoveAt(b);
                                }
                            }
                        }

                        //根據每一類型找出長度
                        List<Tuple<string, List<string>, List<string>>> small_t_list = new List<Tuple<string, List<string>, List<string>>>();
                        List<Tuple<string, List<string>, List<string>>> small_s_list = new List<Tuple<string, List<string>, List<string>>>();
                        foreach (string classname in steel_class)
                        {
                            List<string> length_list = new List<string>();  //長度list
                            List<string> amount_list = new List<string>();  //支數list

                            foreach (FamilyInstance Fi in side_steel)
                            {
                                if (Fi.Name == classname)
                                {
                                    double loline = double.Parse(Fi.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    length_list.Add(Math.Round(loline, 3).ToString()); //將同類型的長度新增到list
                                }

                            }

                            //list裡已有同類型的型鋼的長度，現在將過濾重複的長度
                            for (int a = 0; a < length_list.Count(); a++)
                            {
                                for (int b = length_list.Count() - 1; b > a; b--)
                                {
                                    if (length_list[a] == length_list[b])
                                    {
                                        length_list.RemoveAt(b);
                                    }
                                }
                            }

                            //長度都不重複了，利用長度與類型計算支數

                            foreach (string length in length_list)
                            {
                                List<FamilyInstance> amount_fami = new List<FamilyInstance>();
                                foreach (FamilyInstance fami in side_steel)
                                {
                                    double curve = double.Parse(fami.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    if (Math.Round(curve, 3).ToString() == length)
                                    {
                                        amount_fami.Add(fami);
                                    }
                                }
                                amount_list.Add(amount_fami.Count().ToString());
                            }

                            //所以現在每一個類型都有兩個list一個是長度一個是支數

                            Tuple<string, List<string>, List<string>> small_t = Tuple.Create(classname, length_list, amount_list);
                            small_t_list.Add(small_t);
                        }

                        foreach (string classname in support_class)
                        {
                            List<string> length_list = new List<string>();  //長度list
                            List<string> amount_list = new List<string>();  //支數list

                            foreach (FamilyInstance Fi in support)
                            {
                                if (Fi.Name == classname)
                                {
                                    double loline = double.Parse(Fi.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    length_list.Add(Math.Round(loline, 3).ToString()); //將同類型的長度新增到list
                                }

                            }

                            //list裡已有同類型的型鋼的長度，現在將過濾重複的長度
                            for (int a = 0; a < length_list.Count(); a++)
                            {
                                for (int b = length_list.Count() - 1; b > a; b--)
                                {
                                    if (length_list[a] == length_list[b])
                                    {
                                        length_list.RemoveAt(b);
                                    }
                                }
                            }

                            //長度都不重複了，利用長度與類型計算支數

                            foreach (string length in length_list)
                            {
                                List<FamilyInstance> amount_fami = new List<FamilyInstance>();
                                foreach (FamilyInstance fami in support)
                                {
                                    double curve = double.Parse(fami.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    if (Math.Round(curve, 3).ToString() == length)
                                    {
                                        amount_fami.Add(fami);
                                    }
                                }
                                amount_list.Add(amount_fami.Count().ToString());
                            }

                            //所以現在每一個類型都有兩個list一個是長度一個是支數

                            Tuple<string, List<string>, List<string>> small_t = Tuple.Create(classname, length_list, amount_list);
                            small_s_list.Add(small_t);
                        }
                        Tuple<string, List<Tuple<string, List<string>, List<string>>>> mid_t = Tuple.Create("回撐階數" + floorname, small_t_list);
                        Tuple<string, List<Tuple<string, List<string>, List<string>>>> mid_s = Tuple.Create("回撐階數" + floorname, small_s_list);
                        mid_back_t_list.Add(mid_t);
                        mid_back_s_list.Add(mid_s);
                    }


                    //圍囹部分
                    List<Level> level_count = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name.Contains("斷面" + check_str + "-開挖階數")).ToList();

                    List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>> mid_t_list = new List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>>();
                    List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>> mid_s_list = new List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>>();
                    List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>> mid_p_list = new List<Tuple<string, List<Tuple<string, List<string>, List<string>>>>>();
                    for (int j = 1; j < level_count.Count(); j++)
                    {
                        List<FamilyInstance> side_steel = new List<FamilyInstance>(); //抓圍囹
                        List<FamilyInstance> support = new List<FamilyInstance>(); //抓支撐
                        List<FamilyInstance> slope = new List<FamilyInstance>(); //抓斜撐
                        List<FamilyInstance> test = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().ToList();
                        foreach (FamilyInstance famtest in test)
                        {
                            try
                            {
                                string comment = famtest.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();
                                if (comment != null && comment.Contains(check_str + "-" + j.ToString() + "-圍囹"))
                                {
                                    side_steel.Add(famtest);
                                }
                                if (comment != null && comment.Contains(check_str + "-" + j.ToString() + "-支撐"))
                                {
                                    support.Add(famtest);
                                }
                                if (comment != null && comment.Contains(check_str + "-" + j.ToString() + "-斜撐"))
                                {
                                    slope.Add(famtest);
                                }
                            }
                            catch { }
                        }

                        List<string> steel_class = new List<string>(); //圍囹的型鋼類型
                        List<string> support_class = new List<string>(); //支撐的型鋼類型
                        IList<string> slope_class = new List<string>();//斜撐的型鋼類型
                                                                       //載入類型
                        foreach (FamilyInstance FamIn in side_steel)
                        {
                            LocationCurve locu = FamIn.Location as LocationCurve;
                            steel_class.Add(FamIn.Name);
                        }
                        foreach (FamilyInstance FamIn in support)
                        {
                            LocationCurve locu = FamIn.Location as LocationCurve;
                            support_class.Add(FamIn.Name);
                        }
                        foreach (FamilyInstance famin in slope)
                        {
                            if (famin.Name == "斜撐")
                            {
                                string H = famin.LookupParameter("H").AsValueString();
                                string B = famin.LookupParameter("B").AsValueString();
                                string tw = famin.LookupParameter("tw").AsValueString();
                                string tf = famin.LookupParameter("tf").AsValueString();
                                string name = string.Format("H{0}x{1}x{2}x{3}", H, B, tw, tf);
                                slope_class.Add(name);
                            }
                            else
                            {
                                slope_class.Add(famin.Name);
                            }
                        }

                        //篩選重複的類型
                        for (int a = 0; a < steel_class.Count(); a++)
                        {
                            for (int b = steel_class.Count() - 1; b > a; b--)
                            {
                                if (steel_class[a] == steel_class[b])
                                {
                                    steel_class.RemoveAt(b);
                                }
                            }
                        }
                        for (int a = 0; a < support_class.Count(); a++)
                        {
                            for (int b = support_class.Count() - 1; b > a; b--)
                            {
                                if (support_class[a] == support_class[b])
                                {
                                    support_class.RemoveAt(b);
                                }
                            }
                        }
                        for (int a = 0; a < slope_class.Count(); a++)
                        {
                            for (int b = slope_class.Count() - 1; b > a; b--)
                            {
                                if (slope_class[a] == slope_class[b])
                                {
                                    slope_class.RemoveAt(b);
                                }
                            }
                        }
                        //根據每一類型找出長度

                        List<Tuple<string, List<string>, List<string>>> small_t_list = new List<Tuple<string, List<string>, List<string>>>();
                        List<Tuple<string, List<string>, List<string>>> small_s_list = new List<Tuple<string, List<string>, List<string>>>();
                        List<Tuple<string, List<string>, List<string>>> small_p_list = new List<Tuple<string, List<string>, List<string>>>();
                        foreach (string classname in steel_class) //圍囹的長度支數排列
                        {
                            List<string> length_list = new List<string>();  //長度list
                            List<string> amount_list = new List<string>();  //支數list

                            foreach (FamilyInstance Fi in side_steel)
                            {
                                if (Fi.Name == classname)
                                {
                                    double loline = double.Parse(Fi.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    length_list.Add(Math.Round(loline, 3).ToString()); //將同類型的長度新增到list
                                }

                            }

                            //list裡已有同類型的型鋼的長度，現在將過濾重複的長度
                            for (int a = 0; a < length_list.Count(); a++)
                            {
                                for (int b = length_list.Count() - 1; b > a; b--)
                                {
                                    if (length_list[a] == length_list[b])
                                    {
                                        length_list.RemoveAt(b);
                                    }
                                }
                            }

                            //長度都不重複了，利用長度與類型計算支數

                            foreach (string length in length_list)
                            {
                                List<FamilyInstance> amount_fami = new List<FamilyInstance>();
                                foreach (FamilyInstance fami in side_steel)
                                {
                                    double curve = double.Parse(fami.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    if (Math.Round(curve, 3).ToString() == length)
                                    {
                                        amount_fami.Add(fami);
                                    }
                                }
                                amount_list.Add(amount_fami.Count().ToString());
                            }

                            //所以現在每一個類型都有兩個list一個是長度一個是支數

                            Tuple<string, List<string>, List<string>> small_t = Tuple.Create(classname, length_list, amount_list);
                            small_t_list.Add(small_t);
                        }

                        foreach (string classname in support_class) //支撐的長度支數排列
                        {
                            List<string> length_list = new List<string>();  //長度list
                            List<string> amount_list = new List<string>();  //支數list

                            foreach (FamilyInstance Fi in support)
                            {
                                if (Fi.Name == classname)
                                {
                                    double loline = double.Parse(Fi.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    length_list.Add(Math.Round(loline, 3).ToString()); //將同類型的長度新增到list
                                }

                            }

                            //list裡已有同類型的型鋼的長度，現在將過濾重複的長度
                            for (int a = 0; a < length_list.Count(); a++)
                            {
                                for (int b = length_list.Count() - 1; b > a; b--)
                                {
                                    if (length_list[a] == length_list[b])
                                    {
                                        length_list.RemoveAt(b);
                                    }
                                }
                            }

                            //長度都不重複了，利用長度與類型計算支數

                            foreach (string length in length_list)
                            {
                                List<FamilyInstance> amount_fami = new List<FamilyInstance>();
                                foreach (FamilyInstance fami in support)
                                {
                                    double curve = double.Parse(fami.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    if (Math.Round(curve, 3).ToString() == length)
                                    {
                                        amount_fami.Add(fami);
                                    }
                                }
                                amount_list.Add(amount_fami.Count().ToString());
                            }

                            //所以現在每一個類型都有兩個list一個是長度一個是支數

                            Tuple<string, List<string>, List<string>> small_t = Tuple.Create(classname, length_list, amount_list);
                            small_s_list.Add(small_t);
                        }

                        foreach (string classname in slope_class)
                        {
                            List<string> length_list = new List<string>();  //長度list
                            List<string> amount_list = new List<string>();  //支數list

                            foreach (FamilyInstance Fi in slope)
                            {
                                if (Fi.Name == classname)
                                {
                                    double loline = double.Parse(Fi.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                    length_list.Add(Math.Round(loline, 3).ToString()); //將同類型的長度新增到list
                                }
                                else
                                {
                                    double thick = double.Parse(Fi.LookupParameter("支撐厚度").AsValueString());
                                    double radius = double.Parse(Fi.LookupParameter("中間樁直徑").AsValueString());
                                    double angle = Math.Atan(1500 / (1500 - thick / 1 - radius / 2)); //斜撐角度
                                    string slope_length = (Math.Round((1500 / Math.Sin(angle)) / 1000, 3)).ToString();
                                    length_list.Add(slope_length);
                                }
                            }

                            //list裡已有同類型的型鋼的長度，現在將過濾重複的長度
                            for (int a = 0; a < length_list.Count(); a++)
                            {
                                for (int b = length_list.Count() - 1; b > a; b--)
                                {
                                    if (length_list[a] == length_list[b])
                                    {
                                        length_list.RemoveAt(b);
                                    }
                                }
                            }

                            //長度都不重複了，利用長度與類型計算支數

                            foreach (string length in length_list)
                            {
                                List<FamilyInstance> amount_fami = new List<FamilyInstance>();
                                foreach (FamilyInstance fami in slope)
                                {
                                    if (fami.Name == classname)
                                    {
                                        double curve = double.Parse(fami.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM).AsValueString()) / 1000;
                                        if (Math.Round(curve, 3).ToString() == length)
                                        {
                                            amount_fami.Add(fami);
                                        }
                                    }
                                    else
                                    {
                                        double thick = double.Parse(fami.LookupParameter("支撐厚度").AsValueString());
                                        double radius = double.Parse(fami.LookupParameter("中間樁直徑").AsValueString());
                                        double angle = Math.Atan(1500 / (1500 - thick / 1 - radius / 2)); //斜撐角度
                                        string slope_length = (Math.Round((1500 / Math.Sin(angle)) / 1000, 3)).ToString();
                                        if (slope_length == length)
                                        {
                                            amount_fami.Add(fami);
                                        }
                                    }
                                }
                                amount_list.Add(amount_fami.Count().ToString());
                            }

                            //所以現在每一個類型都有兩個list一個是長度一個是支數

                            Tuple<string, List<string>, List<string>> small_t = Tuple.Create(classname, length_list, amount_list);
                            small_p_list.Add(small_t);
                        }
                        Tuple<string, List<Tuple<string, List<string>, List<string>>>> mid_t = Tuple.Create("階數" + j, small_t_list);
                        Tuple<string, List<Tuple<string, List<string>, List<string>>>> mid_s = Tuple.Create("階數" + j, small_s_list);
                        Tuple<string, List<Tuple<string, List<string>, List<string>>>> mid_p = Tuple.Create("階數" + j, small_p_list);
                        mid_t_list.Add(mid_t);
                        mid_s_list.Add(mid_s);
                        mid_p_list.Add(mid_p);
                    }



                    //監測儀器部分
                    List<FamilyInstance> sincere_rhombus = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name.Contains("壁中傾度管")).ToList();
                    List<FamilyInstance> hollow_rhombus = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name.Contains("土中傾度管")).ToList();
                    List<FamilyInstance> sincere_triangle = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name.Contains("連續壁沉陷觀測點")).ToList();
                    List<FamilyInstance> hollow_circle = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name.Contains("觀測井")).ToList();
                    List<FamilyInstance> sincere_circle = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name.Contains("水壓計")).ToList();
                    List<FamilyInstance> arrowhead = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name.Contains("支撐應變計")).ToList();


                    //土方
                    //該斷面所有開挖階數
                    List<Level> level_depth = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().Where(x => x.Name.Contains("斷面" + check_str + "-開挖階數")).ToList();
                    //該段面的空間
                    SpatialElement room = new FilteredElementCollector(doc).OfClass(typeof(SpatialElement)).Cast<SpatialElement>().Where(x => x.Name.Contains(check_str)).ToList().First();
                    double area = (room.Area * 304.8 * 304.8 / 1000 / 1000);
                    area = Math.Round(area, 3);
                    List<string> mud_square = new List<string>();
                    for (int g = 0; g != level_depth.Count(); g++)
                    {
                        if (g == 0)
                        {
                            double de = double.Parse(level_depth[g].get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()) * (-1) / 1000;
                            string mud = area.ToString() + "*" + de.ToString();
                            mud_square.Add(mud);
                        }
                        else
                        {
                            double de = (double.Parse(level_depth[g - 1].get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString())
                                - double.Parse(level_depth[g].get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString())) / 1000;
                            string mud = area.ToString() + "*" + de.ToString();
                            mud_square.Add(mud);
                        }

                    }

                    total_mud.Add(mud_square);
                    //資料讀取完畢
                    //開始寫進excel

                    try
                    {
                        Excel.Application Eapp = new Excel.Application();

                        /*
                        if (files_path[0] == file_path)
                        {
                            string example = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase).ToString().Substring(6);
                            example += @"\數量表_test.xls";
                            Excel.Workbook EWB2 = Eapp.Workbooks.Open(example);
                            EWB2.SaveAs(path, Type.Missing, "", "", Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, 1, false, Type.Missing, Type.Missing, Type.Missing);
                        }*/
                        Excel.Workbook EWb = Eapp.Workbooks.Open(path);
                        Excel.Worksheet EWs = EWb.Worksheets[1];
                        Excel.Worksheet DataSheet = EWb.Worksheets[5]; //型鋼資料庫
                        Excel.Range ERa_whole = EWs.UsedRange;
                        Excel.Range Data = DataSheet.UsedRange;

                        Excel.Range StrRow = EWs.Rows["8:15"];
                        StrRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                        EWs.Cells[8, 1] = "斷面" + check_str;
                        EWs.Cells[9, 1] = wall_list[0].Name;
                        EWs.Cells[9, 2] = "m2";
                        EWs.Cells[9, 3] = "=" + total_l.ToString() + "*" + depth.ToString();
                        EWs.Cells[9, 4] = total_l.ToString() + "*" + depth.ToString();
                        EWs.Cells[9, 5] = "長x深";
                        EWs.Cells[10, 1] = "混凝土";
                        EWs.Cells[10, 2] = "m3";
                        EWs.Cells[10, 3] = "=" + total_l.ToString() + "*" + depth.ToString() + "*" + reader.wall_width.ToString();
                        EWs.Cells[10, 4] = total_l.ToString() + "*" + depth.ToString() + "*" + reader.wall_width.ToString();
                        EWs.Cells[10, 5] = "長x深x厚度";

                        //中間樁部分

                        EWs.Cells[11, 1] = "中間樁";
                        EWs.Cells[11, 2] = "支";
                        EWs.Cells[11, 4] = count_mid[0].ToString();
                        Excel.Range col_unit = Data.Find(reader.column[0].Item5, MatchCase: true);

                        int col_unit_row = col_unit.Row;
                        string col_unit_str = DataSheet.Cells[col_unit_row, 14].Value2.ToString();

                        string burry = (reader.column[0].Item6 - ste_depth).ToString();
                        EWs.Cells[12, 1] = "H型鋼(" + reader.column[0].Item5 + ")";
                        EWs.Cells[12, 2] = "t";
                        EWs.Cells[12, 4] = burry + "*" + count_mid[0].ToString() + "*" + col_unit_str;
                        EWs.Cells[12, 3] = "=" + burry + "*" + count_mid[0].ToString() + "*" + col_unit_str;
                        EWs.Cells[12, 5] = "長 x 支數 x 單位重";


                        EWs.Cells[13, 1] = "H型鋼(" + reader.column[0].Item5 + ")(埋入段)";
                        EWs.Cells[13, 2] = "t";
                        EWs.Cells[13, 4] = ste_depth.ToString() + "*" + count_mid[0].ToString() + "*" + col_unit_str;
                        EWs.Cells[13, 3] = "=" + ste_depth.ToString() + "*" + count_mid[0].ToString() + "*" + col_unit_str;
                        EWs.Cells[13, 5] = "長 x 支數 x 單位重";

                        EWs.Cells[14, 1] = "混凝土";
                        EWs.Cells[14, 2] = "m3";
                        EWs.Cells[14, 4] = string.Format("{0}*3.14159*{1}^2*{2}", reader.column[0].Item3.ToString(), rad, count_mid[0].ToString());
                        EWs.Cells[14, 3] = string.Format("={0}*3.14159*{1}^2*{2}", reader.column[0].Item3.ToString(), rad, count_mid[0].ToString());
                        EWs.Cells[14, 5] = "長 x 面積 x 支";

                        //型鋼部分

                        EWs.Cells[15, 1] = "開挖階數";
                        int count = 0;
                        int big_count = 0;
                        int count1 = 0;

                        for (int e = 0; e != mid_t_list.Count(); e++)//不同階數
                        {

                            if (mid_s_list[0].Item2.Count == 0 && mid_t_list[0].Item2.Count == 0 && mid_p_list[0].Item2.Count == 0)
                            { break; }
                            else
                            {
                                Excel.Range Era2 = EWs.Rows[16 + count1 + count + big_count, Type.Missing];
                                Era2.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[16 + count1 + count + big_count, 1] = mid_t_list[e].Item1;//寫入階數
                                Era2 = null;
                            }

                            if (mid_t_list[e].Item2.Count() != 0)
                            {
                                Excel.Range ERa_Name = EWs.Rows[17 + count1 + count + big_count, Type.Missing];
                                ERa_Name.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[17 + count1 + count + big_count, 1] = mid_t_list[e].Item1 + "圍囹";
                                for (int d = 0; d != mid_t_list[e].Item2.Count(); d++)//寫入不同類型型鋼
                                {

                                    for (int c = 0; c != mid_t_list[e].Item2[d].Item3.Count(); c++)//寫入不同長度
                                    {
                                        try
                                        {
                                            Excel.Range ERa = EWs.Rows[18 + count1 + count + big_count, Type.Missing];
                                            ERa.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                            Excel.Range unit = Data.Find(mid_t_list[e].Item2[d].Item1, MatchCase: true);
                                            int unit_row = unit.Row;
                                            string unit_str = DataSheet.Cells[unit_row, 14].Value2.ToString();

                                            EWs.Cells[18 + count1 + count + big_count, 1] = mid_t_list[e].Item2[d].Item1;//型鋼名稱
                                            EWs.Cells[18 + count1 + count + big_count, 2] = "t";
                                            EWs.Cells[18 + count1 + count + big_count, 4] = mid_t_list[e].Item2[d].Item3[c] + "*" + mid_t_list[e].Item2[d].Item2[c] + "*" + unit_str;//支數*長度
                                            EWs.Cells[18 + count1 + count + big_count, 3] = "=" + mid_t_list[e].Item2[d].Item3[c] + "*" + mid_t_list[e].Item2[d].Item2[c] + "*" + unit_str;
                                            EWs.Cells[18 + count1 + count + big_count, 5] = "支數*長度*單位重\n" + mid_t_list[e].Item2[d].Item1;
                                            ERa = null;
                                            unit = null;
                                            count++;
                                        }
                                        catch (Exception ae) { reader.CloseEx(); TaskDialog.Show("error", Environment.NewLine + ae.Message); }
                                    }

                                }
                                ERa_Name = null;
                                count1++;
                            }

                            if (mid_s_list[e].Item2.Count() != 0)
                            {
                                Excel.Range ERa_Name = EWs.Rows[17 + count1 + count + big_count, Type.Missing];
                                ERa_Name.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[17 + count1 + count + big_count, 1] = mid_s_list[e].Item1 + "支撐";
                                for (int d = 0; d != mid_s_list[e].Item2.Count(); d++)//寫入不同類型型鋼
                                {

                                    for (int c = 0; c != mid_s_list[e].Item2[d].Item3.Count(); c++)//寫入不同長度
                                    {
                                        Excel.Range ERa = EWs.Rows[18 + count1 + count + big_count, Type.Missing];
                                        ERa.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                        Excel.Range unit = Data.Find(mid_s_list[e].Item2[d].Item1, MatchCase: true);
                                        int unit_row = unit.Row;
                                        string unit_str = DataSheet.Cells[unit_row, 14].Value2.ToString();
                                        EWs.Cells[18 + count1 + count + big_count, 1] = mid_s_list[e].Item2[d].Item1;//型鋼名稱
                                        EWs.Cells[18 + count1 + count + big_count, 2] = "t";
                                        EWs.Cells[18 + count1 + count + big_count, 4] = mid_s_list[e].Item2[d].Item3[c] + "*" + mid_s_list[e].Item2[d].Item2[c] + "*" + unit_str;//支數*長度
                                        EWs.Cells[18 + count1 + count + big_count, 3] = "=" + mid_s_list[e].Item2[d].Item3[c] + "*" + mid_s_list[e].Item2[d].Item2[c] + "*" + unit_str;
                                        EWs.Cells[18 + count1 + count + big_count, 5] = "支數*長度*單位重\n" + mid_s_list[e].Item2[d].Item1;
                                        ERa = null;
                                        unit = null;
                                        count++;
                                    }

                                }
                                ERa_Name = null;
                                count1++;
                            }
                            if (mid_p_list[e].Item2.Count() != 0)
                            {
                                Excel.Range ERa_Name = EWs.Rows[17 + count1 + count + big_count, Type.Missing];
                                ERa_Name.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[17 + count1 + count + big_count, 1] = mid_p_list[e].Item1 + "斜撐";
                                for (int d = 0; d != mid_p_list[e].Item2.Count(); d++)//寫入不同類型型鋼
                                {

                                    for (int c = 0; c != mid_p_list[e].Item2[d].Item3.Count(); c++)//寫入不同長度
                                    {
                                        Excel.Range ERa = EWs.Rows[18 + count1 + count + big_count, Type.Missing];
                                        ERa.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                        Excel.Range unit = Data.Find(mid_p_list[e].Item2[d].Item1, MatchCase: true);
                                        int unit_row = unit.Row;
                                        string unit_str = DataSheet.Cells[unit_row, 14].Value2.ToString();
                                        EWs.Cells[18 + count1 + count + big_count, 1] = mid_p_list[e].Item2[d].Item1;//型鋼名稱
                                        EWs.Cells[18 + count1 + count + big_count, 2] = "t";
                                        EWs.Cells[18 + count1 + count + big_count, 4] = mid_p_list[e].Item2[d].Item3[c] + "*" + mid_p_list[e].Item2[d].Item2[c] + "*" + unit_str;//支數*長度
                                        EWs.Cells[18 + count1 + count + big_count, 3] = "=" + mid_p_list[e].Item2[d].Item3[c] + "*" + mid_p_list[e].Item2[d].Item2[c] + "*" + unit_str;
                                        EWs.Cells[18 + count1 + count + big_count, 5] = "支數*長度*單位重\n" + mid_p_list[e].Item2[d].Item1;
                                        ERa = null;
                                        unit = null;
                                        count++;
                                    }

                                }
                                ERa_Name = null;
                                count1++;
                            }

                            big_count++;
                        }

                        //回填部分

                        for (int e = 0; e != mid_back_t_list.Count(); e++)//不同階數
                        {

                            if (mid_back_t_list[0].Item2.Count == 0 && mid_back_s_list[0].Item2.Count == 0)
                            { break; }
                            else
                            {
                                Excel.Range Era2 = EWs.Rows[18 + count1 + count + big_count, Type.Missing];
                                Era2.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[18 + count1 + count + big_count, 1] = mid_back_t_list[e].Item1;//寫入階數
                                Era2 = null;
                            }
                            if (mid_back_t_list[e].Item2.Count() != 0)
                            {
                                Excel.Range ERa_Name = EWs.Rows[19 + count1 + count + big_count, Type.Missing];
                                ERa_Name.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[19 + count1 + count + big_count, 1] = mid_back_t_list[e].Item1 + "圍囹";
                                for (int d = 0; d != mid_back_t_list[e].Item2.Count(); d++)//寫入不同類型型鋼
                                {

                                    for (int c = 0; c != mid_back_t_list[e].Item2[d].Item3.Count(); c++)//寫入不同長度
                                    {
                                        try
                                        {
                                            Excel.Range ERa = EWs.Rows[20 + count1 + count + big_count, Type.Missing];
                                            ERa.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                            Excel.Range unit = Data.Find(mid_back_t_list[e].Item2[d].Item1, MatchCase: true);
                                            int unit_row = unit.Row;
                                            string unit_str = DataSheet.Cells[unit_row, 14].Value2.ToString();

                                            EWs.Cells[20 + count1 + count + big_count, 1] = mid_back_t_list[e].Item2[d].Item1;//型鋼名稱
                                            EWs.Cells[20 + count1 + count + big_count, 2] = "t";
                                            EWs.Cells[20 + count1 + count + big_count, 4] = mid_back_t_list[e].Item2[d].Item3[c] + "*" + mid_back_t_list[e].Item2[d].Item2[c] + "*" + unit_str;//支數*長度
                                            EWs.Cells[20 + count1 + count + big_count, 3] = "=" + mid_back_t_list[e].Item2[d].Item3[c] + "*" + mid_back_t_list[e].Item2[d].Item2[c] + "*" + unit_str;
                                            EWs.Cells[20 + count1 + count + big_count, 5] = "支數*長度*單位重\n" + mid_back_t_list[e].Item2[d].Item1;
                                            ERa = null;
                                            unit = null;
                                            count++;
                                        }
                                        catch (Exception ae) { reader.CloseEx(); TaskDialog.Show("error", Environment.NewLine + ae.Message); }
                                    }

                                }
                                ERa_Name = null;
                                count1++;
                            }
                            if (mid_back_s_list[e].Item2.Count() != 0)
                            {
                                Excel.Range ERa_Name = EWs.Rows[19 + count1 + count + big_count, Type.Missing];
                                ERa_Name.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                EWs.Cells[19 + count1 + count + big_count, 1] = mid_back_s_list[e].Item1 + "支撐";
                                for (int d = 0; d != mid_back_s_list[e].Item2.Count(); d++)//寫入不同類型型鋼
                                {

                                    for (int c = 0; c != mid_back_s_list[e].Item2[d].Item3.Count(); c++)//寫入不同長度
                                    {
                                        try
                                        {
                                            Excel.Range ERa = EWs.Rows[20 + count1 + count + big_count, Type.Missing];
                                            ERa.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                            Excel.Range unit = Data.Find(mid_back_s_list[e].Item2[d].Item1, MatchCase: true);
                                            int unit_row = unit.Row;
                                            string unit_str = DataSheet.Cells[unit_row, 14].Value2.ToString();

                                            EWs.Cells[20 + count1 + count + big_count, 1] = mid_back_s_list[e].Item2[d].Item1;//型鋼名稱
                                            EWs.Cells[20 + count1 + count + big_count, 2] = "t";
                                            EWs.Cells[20 + count1 + count + big_count, 4] = mid_back_s_list[e].Item2[d].Item3[c] + "*" + mid_back_s_list[e].Item2[d].Item2[c] + "*" + unit_str;//支數*長度
                                            EWs.Cells[20 + count1 + count + big_count, 3] = "=" + mid_back_s_list[e].Item2[d].Item3[c] + "*" + mid_back_s_list[e].Item2[d].Item2[c] + "*" + unit_str;
                                            EWs.Cells[20 + count1 + count + big_count, 5] = "支數*長度*單位重\n" + mid_back_s_list[e].Item2[d].Item1;
                                            ERa = null;
                                            unit = null;
                                            count++;
                                        }
                                        catch (Exception ae) { reader.CloseEx(); TaskDialog.Show("error", Environment.NewLine + ae.Message); }
                                    }

                                }
                                ERa_Name = null;
                                count1++;
                            }

                            big_count++;
                        }

                        StrRow = null;

                        Excel.Worksheet EWs2 = EWb.Worksheets[2];
                        Excel.Range ERa2_whole = EWs2.UsedRange;
                        EWs2.Cells[11, 2] = "壁中傾度管";
                        EWs2.Cells[11, 3] = "個";
                        EWs2.Cells[11, 4] = sincere_rhombus.Count().ToString();
                        EWs2.Cells[12, 2] = "土中傾度管";
                        EWs2.Cells[12, 3] = "個";
                        EWs2.Cells[12, 4] = hollow_rhombus.Count().ToString();
                        EWs2.Cells[13, 2] = "連續壁沉陷觀測點";
                        EWs2.Cells[13, 3] = "個";
                        EWs2.Cells[13, 4] = sincere_triangle.Count().ToString();
                        EWs2.Cells[14, 2] = "觀測井";
                        EWs2.Cells[14, 3] = "個";
                        EWs2.Cells[14, 4] = hollow_circle.Count().ToString();
                        EWs2.Cells[15, 2] = "水壓計";
                        EWs2.Cells[15, 3] = "個";
                        EWs2.Cells[15, 4] = sincere_circle.Count().ToString();
                        EWs2.Cells[16, 2] = "支撐應變計";
                        EWs2.Cells[16, 3] = "個";
                        EWs2.Cells[16, 4] = arrowhead.Count().ToString();


                        EWb.Save();

                        EWs = null;
                        DataSheet = null;
                        Data = null;
                        EWs2 = null;
                        col_unit = null;
                        EWb.Close();
                        EWb = null;
                        Eapp.Quit();
                        Eapp = null;
                    }
                    catch (Exception e) { TaskDialog.Show("error", e.ToString()); }
                }
                //最後寫入土方量
                try
                {

                    Excel.Application Eapp3 = new Excel.Application();

                    Excel.Workbook EWb3 = Eapp3.Workbooks.Open(path);

                    Excel.Worksheet EWs3 = EWb3.Worksheets[1];
                    int mcount = 0;
                    int mscount = 0;
                    for (int u = 0; u != files_path.Count(); u++)
                    {
                        ExReader ex = new ExReader();
                        ex.SetData(files_path[u], 1);
                        ex.PassWallData();
                        ex.CloseEx();
                        Excel.Range Era3 = EWs3.Rows[8 + mcount + mscount, Type.Missing];
                        Era3.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                        EWs3.Cells[8 + mcount + mscount, 1] = "斷面" + ex.section + "土方量";
                        for (int x = 0; x != total_mud[u].Count(); x++)
                        {
                            Excel.Range ERa4 = EWs3.Rows[9 + mcount + mscount, Type.Missing];
                            ERa4.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                            EWs3.Cells[9 + mcount + mscount, 1] = "階數" + (x + 1).ToString();
                            EWs3.Cells[9 + mcount + mscount, 2] = "m3";
                            EWs3.Cells[9 + mcount + mscount, 3] = "=" + total_mud[u][x];
                            EWs3.Cells[9 + mcount + mscount, 4] = total_mud[u][x];
                            EWs3.Cells[9 + mcount + mscount, 5] =
                            ERa4 = null;
                            mscount++;
                        }
                        mcount++;
                    }
                    EWb3.Save();
                    EWs3 = null;
                    EWb3.Close();
                    EWb3 = null;
                    Eapp3.Quit();
                    Eapp3 = null;

                }
                catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message); }
                TaskDialog.Show("寫入完成", "寫入結束。");
            }
            catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message + e.StackTrace); }

        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
