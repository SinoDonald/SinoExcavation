using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;

namespace SinoExcavation_2025
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class OutputExcel : IExternalEventHandler
    {
        //深開挖類型
        public string excavationType
        {
            get;
            set;
        }

        //excel表格樣板
        public string FilePath
        {
            get;
            set;
        }

        public void Execute(UIApplication app)
        {

            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;


            

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter wallFilter = new ElementCategoryFilter(BuiltInCategory.OST_Walls);
            IList<Element> wallList = collector.WherePasses(wallFilter).WhereElementIsNotElementType().ToElements();

            collector = new FilteredElementCollector(doc);
            ElementCategoryFilter columnFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
            IList<Element> columnList = collector.WherePasses(columnFilter).WhereElementIsNotElementType().ToElements();

            collector = new FilteredElementCollector(doc);
            ElementCategoryFilter levelFilter = new ElementCategoryFilter(BuiltInCategory.OST_Levels);
            IList<Element> levelList = collector.WherePasses(levelFilter).WhereElementIsNotElementType().ToElements();

            collector = new FilteredElementCollector(doc);
            ElementCategoryFilter frameFilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            IList<Element> frameList = collector.WherePasses(frameFilter).WhereElementIsNotElementType().ToElements();

            collector = new FilteredElementCollector(doc);
            ElementCategoryFilter floorFilter = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
            IList<Element> floorList = collector.WherePasses(floorFilter).WhereElementIsNotElementType().ToElements();

            collector = new FilteredElementCollector(doc);
            ElementCategoryFilter roomFilter = new ElementCategoryFilter(BuiltInCategory.OST_Rooms);
            IList<Element> roomList = collector.WherePasses(roomFilter).WhereElementIsNotElementType().ToElements();

            //分開支撐與圍囹
            IList<Element> frame_cofferdam_list = new List<Element>();
            IList<Element> frame_support_list = new List<Element>();
            string frame_type = "";
            foreach(Element frame in frameList)
            {
                frame_type = frame.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[2];
                if (frame_type == "支撐")
                {
                    frame_support_list.Add(frame);
                }
                else if (frame_type == "圍囹")
                {
                    frame_cofferdam_list.Add(frame);
                }
            }
            

            IList<IList<Element>> wallListBySection = dividedBySection(wallList);
            IList<IList<Element>> roomListBySection = dividedBySection(roomList);
            IList<IList<Element>> floorListBySection = dividedBySection(floorList);
            IList<IList<Element>> frame_support_bysection_list = dividedBySection(frame_support_list);
            IList<IList<Element>> frame_cofferdam_bysection_list = dividedBySection(frame_cofferdam_list);
            IList<IList<Element>> columnListBySection = dividedBySectionForColumn(columnList);
            IList<IList<Element>> levelListBySection = dividedBySectionForLevel(levelList);
            IList<string> sectionList = new List<string>();
            

            double length = 0; //長度
            double columnDepth = 0; //樁深埋入深度
            double columnDiameter = 0; //樁徑
            double HDepth = 0; //型鋼埋入深度
            double depth = 0; //不連續高度
            double width = 0;
            double height = 0;
            double horizon = 0;
            double stepDepth = 0;
            string section = "";
            

            for (int sectionCount = 0; sectionCount < wallListBySection.Count; sectionCount++)
            {

                //檔案路徑
                string filepath = FilePath;
                //應用程序
                Excel.Application excelAPP = new Excel.Application();
                //檔案
                Excel.Workbook excelWorkbook = excelAPP.Workbooks.Open(filepath);
                //工作表
                Excel.Worksheet excelWorksheet = excelWorkbook.Worksheets["開挖基本設定"];

                ExReaderConsole.ExReader excel = new ExReaderConsole.ExReader();

                try
                {
                    wallList = wallListBySection[sectionCount];
                    string sectionName = wallList[0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[0];
                    for (int selectCount = 0; selectCount < wallListBySection.Count; selectCount++)
                    {
                        try
                        {
                            if (frame_support_bysection_list[selectCount][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains(sectionName))
                            {
                                frame_support_list = frame_support_bysection_list[selectCount];
                            }
                        }
                        catch { }
                        try
                        {
                            if (frame_cofferdam_bysection_list[selectCount][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains(sectionName))
                            {
                                frame_cofferdam_list = frame_cofferdam_bysection_list[selectCount];
                            }
                        }
                        catch { }
                        try
                        {
                            if (roomListBySection[selectCount][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains(sectionName))
                            {
                                roomList = roomListBySection[selectCount];
                            }
                        }
                        catch { }
                        try
                        {
                            if (levelListBySection[selectCount][0].Name.Contains(sectionName))
                            {
                                levelList = levelListBySection[selectCount];
                            }
                        }
                        catch { }
                        try
                        {
                            if (columnListBySection[selectCount][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains(sectionName))
                            {
                                columnList = columnListBySection[selectCount];
                            }
                        }
                        catch { }
                        try
                        {
                            if (floorListBySection[selectCount][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Contains(sectionName))
                            {
                                floorList = floorListBySection[selectCount];
                            }
                        }
                        catch { }


                    }
                    try
                    {
                        if (levelListBySection[wallListBySection.Count][0].Name.Contains(sectionName))
                        {
                            levelList = levelListBySection[wallListBySection.Count];
                        }
                    }
                    catch { }
                    

                    //IList<XYZ> pointList = new List<XYZ>();

                    IList<IList<Element>> supportFrameListwithLevel = new List<IList<Element>>();
                    IList<IList<Element>> cofferdamFrameListwithLevel = new List<IList<Element>>();

                    if (excavationType == "深開挖基礎")
                    {
                        excel.SetData(filepath, 1);

                        var address = excel.FindAddress("分析斷面：");

                        section = wallList[0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[0];
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = section;

                        address = excel.FindAddress("擋土壁型式：");

                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = wallList[0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[1];

                        address = excel.FindAddress("擋土壁長度：");

                        Wall wall = wallList[0] as Wall;
                        double.TryParse(wallList[0].LookupParameter("不連續高度").AsValueString(), out depth);
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = depth / 1000;

                        double.TryParse(wall.WallType.LookupParameter("寬度").AsValueString(), out width);
                        address = excel.FindAddress("連續壁厚度：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = width / 1000;


                        double.TryParse(levelList[levelList.Count - 2].LookupParameter("高程").AsValueString(), out height);
                        double.TryParse(levelListBySection[0][0].LookupParameter("高程").AsValueString(), out horizon);

                        address = excel.FindAddress("開挖深度：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = -height / 1000;
                        address = excel.FindAddress("地表高程：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = -horizon / 1000;

                        address = excel.FindAddress("開挖範圍");
                        /*
                        //找開挖範圍位置 寫入開挖範圍
                        
                        for (int i = 0; i < pointList.Count; i++)
                        {
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2 + 1] = (pointList[i].X * 304.8 / 1000).ToString("0.00");
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2 + 2] = (pointList[i].Y * 304.8 / 1000).ToString("0.00");
                        }*/

                        //找開挖階數位置 寫入開挖街樹
                        address = Tuple.Create(address.Item1 + 2, address.Item2 + 1);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "開挖階數:";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "階數:";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "深度:";

                        int level_count = 0;
                        for (int i = 0; i < levelList.Count - 1; i++)
                        {
                            if (levelList[i].Name.Contains("開挖階數"))
                            {
                                double.TryParse(levelList[i].LookupParameter("高程").AsValueString(), out stepDepth);
                                excelWorksheet.Cells[address.Item1 + i + 1, address.Item2] = i + 1;
                                excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 1] = -stepDepth / 1000;
                                excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 2] = "m";
                                level_count++;
                            }

                        }

                        //寫入支撐資料
                        address = Tuple.Create(address.Item1 + level_count + 3, address.Item2);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "單向支撐或雙向支撐";
                        excelWorksheet.Cells[address.Item1 + 1, address.Item2 - 1] = "支撐階數 : ";
                        excelWorksheet.Cells[address.Item1 + 1, address.Item2] = "階數";
                        excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 1] = "深度";
                        excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 2] = "支數";
                        excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 3] = "型號";
                        excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 4] = "支撐間距";


                        /*for (int i = 0; i < levelList.Count - 2; i++)
                        {
                            supportFrameListwithLevel.Add(new List<Element>());
                        }*/


                        //支撐依照樓層分類
                        for (int i = 0; i < frame_support_list.Count; i++)
                        {
                            bool is_exist = false;
                            string frame_level = frame_support_list[i].LookupParameter("備註").AsString().Split('-')[1];
                            for (int j = 0; j < supportFrameListwithLevel.Count(); j++)
                            {
                                if (frame_level == supportFrameListwithLevel[j][0].LookupParameter("備註").AsString().Split('-')[1])
                                {
                                    is_exist = true;
                                    supportFrameListwithLevel[j].Add(frame_support_list[i]);
                                    break;
                                }
                            }
                            if (is_exist == false)
                            {
                                supportFrameListwithLevel.Add(new List<Element>());
                                supportFrameListwithLevel[supportFrameListwithLevel.Count - 1].Add(frame_support_list[i]);
                            }
                        }

                        //圍囹依照樓層分類
                        for (int i = 0; i < frame_cofferdam_list.Count; i++)
                        {
                            bool is_exist = false;
                            string frame_level = frame_cofferdam_list[i].LookupParameter("備註").AsString().Split('-')[1];
                            for (int j = 0; j < cofferdamFrameListwithLevel.Count(); j++)
                            {
                                if (frame_level == cofferdamFrameListwithLevel[j][0].LookupParameter("備註").AsString().Split('-')[1])
                                {
                                    is_exist = true;
                                    cofferdamFrameListwithLevel[j].Add(frame_cofferdam_list[i]);
                                    break;
                                }
                            }
                            if (is_exist == false)
                            {
                                cofferdamFrameListwithLevel.Add(new List<Element>());
                                cofferdamFrameListwithLevel[cofferdamFrameListwithLevel.Count - 1].Add(frame_cofferdam_list[i]);
                            }
                        }


                        //把支撐依照不同高度分類
                        int countSupoort = 0;
                        int countCofferdam = 0;
                        /*int supportFrameLevel = 0;
                        foreach (Element frame in frame_support_list)
                        {
                            if (frame.LookupParameter("備註").AsString().Contains("支撐"))
                            {
                                try { Double.Parse(frame.LookupParameter("備註").AsString().Split('-')[1].ToString()); } catch { supportFrameListwithRLevel.Add(); }
                                supportFrameListwithLevel[supportFrameLevel - 1].Add(frame);
                            }
                        }*/


                        //remove same level but different depth support frame
                        for (int j = 0; j < supportFrameListwithLevel.Count; j++)
                        {
                            //MessageBox.Show(supportFrameListwithLevel[j].Count.ToString());
                            double tempDepth = 0;
                            double tempDeptht = 0;
                            for (int i = 0; i < supportFrameListwithLevel[j].Count; i++)
                            {
                                if (i == 0)
                                {
                                    double.TryParse(supportFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDepth);
                                }
                                else
                                {
                                    double.TryParse(supportFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDeptht);
                                    if (tempDeptht < tempDepth)
                                    {
                                        double.TryParse(supportFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDepth);
                                    }
                                }
                            }
                            for (int i = supportFrameListwithLevel[j].Count - 1; i >= 0; i--)
                            {

                                double.TryParse(supportFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDeptht);
                                if (Math.Abs(tempDeptht - tempDepth) > 0.0001)
                                {
                                    supportFrameListwithLevel[j].RemoveAt(i);
                                }
                            }
                        }

                        //remove same level but different depth cofferdam frame
                        for (int j = 0; j < cofferdamFrameListwithLevel.Count; j++)
                        {
                            double tempDepth = 0;
                            double tempDeptht = 0;
                            for (int i = 0; i < cofferdamFrameListwithLevel[j].Count; i++)
                            {
                                if (i == 0)
                                {
                                    double.TryParse(cofferdamFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDepth);
                                }
                                else
                                {
                                    double.TryParse(cofferdamFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDeptht);
                                    if (tempDeptht > tempDepth)
                                    {
                                        double.TryParse(cofferdamFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDepth);
                                    }
                                }
                            }
                            for (int i = cofferdamFrameListwithLevel[j].Count - 1; i >= 0; i--)
                            {

                                double.TryParse(cofferdamFrameListwithLevel[j][i].LookupParameter("底部高程").AsValueString(), out tempDeptht);
                                if (Math.Abs(tempDeptht - tempDepth) > 0.0001)
                                {
                                    cofferdamFrameListwithLevel[j].RemoveAt(i);
                                }
                            }
                        }


                        try
                        {
                            //判斷並寫入資料_支撐
                            foreach (IList<Element> supportFrameList in supportFrameListwithLevel)
                            {
                                //找出哪根是最旁邊的
                                int firstFrame = 0;
                                double longestDistance = 0;
                                for (int i = 0; i < supportFrameList.Count(); i++)
                                {
                                    Line supportFrame1 = (supportFrameList[i].Location as LocationCurve).Curve as Line;
                                    for (int j = 0; j < supportFrameList.Count(); j++)
                                    {
                                        Line supportFrame2 = (supportFrameList[j].Location as LocationCurve).Curve as Line;
                                        if (i == 0 && j == 0)
                                        {
                                            //do nothing
                                        }
                                        else
                                        {
                                            if (distanceBetweenLines(supportFrame1, supportFrame2) > longestDistance)
                                            {
                                                longestDistance = distanceBetweenLines(supportFrame1, supportFrame2);
                                                firstFrame = j;
                                            }
                                        }
                                    }
                                }
                                //計算出間距
                                List<double> distanceList = new List<double>();
                                List<double> spacingList = new List<double>();
                                double supportFrameDepth = double.MinValue;
                                foreach (Element frame in supportFrameList)
                                {
                                    double supportFrameDepthTemp;
                                    double.TryParse(frame.LookupParameter("Z 向偏移值").AsValueString(), out supportFrameDepthTemp);
                                    if (supportFrameDepthTemp > supportFrameDepth)
                                    {
                                        supportFrameDepth = supportFrameDepthTemp;
                                    }
                                    Line frameLine1 = (supportFrameList[firstFrame].Location as LocationCurve).Curve as Line;
                                    Line frameLine2 = (frame.Location as LocationCurve).Curve as Line;

                                    distanceList.Add(distanceBetweenLines(frameLine1, frameLine2));
                                }
                                distanceList.Sort();
                                int countSupportFrame = 1;
                                double spacing = 0;
                                bool isOneSupportFrame = true;
                                for (int i = 0; i < distanceList.Count; i++)
                                {
                                    if (i > 0)
                                    {
                                        spacingList.Add(distanceList[i] - distanceList[i - 1]);
                                    }
                                }


                                //計算出間距一單位支數
                                for (int i = 0; i < spacingList.Count; i++)
                                {
                                    spacing += spacingList[i];
                                    if (i > 0)
                                    {

                                        if (spacingList[i] - spacingList[i - 1] > 0.001)
                                        {
                                            //MessageBox.Show(spacingList[i].ToString() + " - " + spacingList[i - 1].ToString());
                                            countSupportFrame = i + 1;
                                            isOneSupportFrame = false;
                                            break;
                                        }
                                    }
                                }

                                if (isOneSupportFrame)
                                {
                                    spacing = spacingList[0];
                                }
                                //寫入資料
                                countSupoort++;
                                excelWorksheet.Cells[address.Item1 + 1 + countSupoort, address.Item2] = supportFrameList[0].LookupParameter("備註").AsString().Split('-')[1].ToString();

                                //double.TryParse(supportFrameList[0].LookupParameter("底部高程").AsValueString(), out supportFrameDepth);
                                excelWorksheet.Cells[address.Item1 + 1 + countSupoort, address.Item2 + 1] = -supportFrameDepth / 1000;
                                excelWorksheet.Cells[address.Item1 + 1 + countSupoort, address.Item2 + 2] = countSupportFrame;
                                excelWorksheet.Cells[address.Item1 + 1 + countSupoort, address.Item2 + 3] = supportFrameList[0].Name;
                                excelWorksheet.Cells[address.Item1 + 1 + countSupoort, address.Item2 + 4] = spacing * 304.8 / 1000;
                            }

                            //寫入圍囹資料
                            address = Tuple.Create(address.Item1 + supportFrameListwithLevel.Count + 3, address.Item2);
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 - 1] = "圍囹階數：";
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2] = "階數";
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 1] = "支數";
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 2] = "型號";


                            //判斷並寫入資料_圍囹
                            foreach (IList<Element> cofferdamFrameList in cofferdamFrameListwithLevel)
                            {
                                //delet unparallel cofferdam frame
                                for (int i = 0; i < cofferdamFrameList.Count; i++)
                                {
                                    Line first_line = (cofferdamFrameList[0].Location as LocationCurve).Curve as Line;
                                    if (i > 0)
                                    {
                                        if (isParallel(first_line, (cofferdamFrameList[i].Location as LocationCurve).Curve as Line) == false)
                                        {
                                            cofferdamFrameList.RemoveAt(i);
                                            i--;
                                        }
                                    }
                                }

                                //找出哪根是最旁邊的
                                int firstFrame = 0;
                                double longestDistance = 0;
                                for (int i = 0; i < cofferdamFrameList.Count(); i++)
                                {
                                    Line supportFrame1 = (cofferdamFrameList[i].Location as LocationCurve).Curve as Line;
                                    for (int j = 0; j < cofferdamFrameList.Count(); j++)
                                    {
                                        Line supportFrame2 = (cofferdamFrameList[j].Location as LocationCurve).Curve as Line;
                                        if (i == 0 && j == 0)
                                        {
                                            //do nothing
                                        }
                                        else
                                        {
                                            if (distanceBetweenLines(supportFrame1, supportFrame2) > longestDistance)
                                            {
                                                longestDistance = distanceBetweenLines(supportFrame1, supportFrame2);
                                                firstFrame = j;
                                            }
                                        }
                                    }
                                }
                                //計算出間距
                                List<double> distanceList = new List<double>();
                                List<double> spacingList = new List<double>();
                                //double cofferdamFrameDepth = double.MinValue;
                                foreach (Element frame in cofferdamFrameList)
                                {
                                    Line frameLine1 = (cofferdamFrameList[firstFrame].Location as LocationCurve).Curve as Line;
                                    Line frameLine2 = (frame.Location as LocationCurve).Curve as Line;

                                    distanceList.Add(distanceBetweenLines(frameLine1, frameLine2));
                                }
                                //MessageBox.Show("9.2");
                                distanceList.Sort();
                                int countCofferdamFrame = 1;
                                double spacing = 0;
                                bool isOneSupportFrame = true;
                                for (int i = 0; i < distanceList.Count; i++)
                                {
                                    if (i > 0)
                                    {
                                        spacingList.Add(distanceList[i] - distanceList[i - 1]);
                                    }
                                }


                                //MessageBox.Show("9.3");
                                //計算出間距一單位支數
                                for (int i = 0; i < spacingList.Count; i++)
                                {
                                    spacing += spacingList[i];
                                    if (i > 0)
                                    {

                                        if (spacingList[i] - spacingList[i - 1] > 0.001)
                                        {
                                            //MessageBox.Show(spacingList[i].ToString() + " - " + spacingList[i - 1].ToString());
                                            countCofferdamFrame = i + 1;
                                            isOneSupportFrame = false;
                                            break;
                                        }
                                    }
                                }

                                if (isOneSupportFrame)
                                {
                                    spacing = spacingList[0];
                                }
                                //寫入資料
                                countCofferdam++;

                                excelWorksheet.Cells[address.Item1 + 1 + countCofferdam, address.Item2] = cofferdamFrameList[0].LookupParameter("備註").AsString().Split('-')[1].ToString();
                                excelWorksheet.Cells[address.Item1 + 1 + countCofferdam, address.Item2 + 1] = countCofferdamFrame;
                                excelWorksheet.Cells[address.Item1 + 1 + countCofferdam, address.Item2 + 2] = cofferdamFrameList[0].Name;
                            }
                        }
                        catch { MessageBox.Show("frame error!!"); }


                        //找表格中間柱位置
                        address = Tuple.Create(address.Item1 + countCofferdam + 5, address.Item2);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "中間柱";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "中間樁間距";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "中間樁長度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 2] = "樁深埋入深度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 3] = "樁徑";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 4] = "型鋼型號";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 5] = "型鋼長度";


                        string first_column = columnList[0].Name;
                        for (int j = 1; j < columnList.Count; j++)
                        {
                            if (columnList[j].Name == first_column)
                            {
                                columnList.RemoveAt(j);
                                j--;
                            }
                            else
                            {
                                first_column = columnList[j].Name;
                            }
                        }

                        for (int j = 0; j < columnList.Count; j++)
                        {
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2] = columnList[j].get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();

                            double.TryParse(columnList[j].LookupParameter("樁深埋入深度").AsValueString(), out columnDepth);
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 2] = columnDepth / 1000;

                            double.TryParse(columnList[j].LookupParameter("樁徑").AsValueString(), out columnDiameter);
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 3] = columnDiameter * 2 / 1000;

                            double.TryParse(columnList[j].LookupParameter("長度").AsValueString(), out length);
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 1] = (columnDepth + length) / 1000;

                            double.TryParse(columnList[j].LookupParameter("型鋼埋入深度").AsValueString(), out HDepth);
                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 5] = (length + HDepth) / 1000;

                            excelWorksheet.Cells[address.Item1 + 1, address.Item2 + 4] = columnList[j].Name.Split('-')[1];
                        }
                        

                        //側牆資料寫入
                        address = Tuple.Create(address.Item1 + columnList.Count + 3, address.Item2);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "側牆:";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "樓層數";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "側牆厚度(m)";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 2] = "混凝土強度(kg/cm2)";

                        //選出側牆
                        IList<Element> side_wall_list = new List<Element>();
                        foreach(Element w in wallList)
                        {
                            bool same_level_type = false;
                            if (w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[1].Contains("側壁"))
                            {
                                for(int j = 0; j < side_wall_list.Count; j++)
                                {
                                    if(w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() == side_wall_list[j].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() && w.Name == side_wall_list[j].Name)
                                    {
                                        same_level_type = true;
                                        break;
                                    }
                                }
                                if(same_level_type == false)
                                {
                                    side_wall_list.Add(w);
                                }
                                
                            }
                        }

                        for(int i = 0; i < side_wall_list.Count; i++)
                        {
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2] = side_wall_list[i].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[1].Split('側')[0];
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 1] = (side_wall_list[i] as Wall).WallType.LookupParameter("寬度").AsDouble() * 304.8/1000;
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 2] = side_wall_list[i].get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();
                        }

                        //樓板回築和回填資料寫入
                        address = Tuple.Create(address.Item1 + side_wall_list.Count + 3, address.Item2);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "樓板回築&回填：";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "深度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "樓層數";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 2] = "厚度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 3] = "混凝土強度(kg/cm2)";

                        for(int j = 0; j < floorList.Count; j++)
                        {
                            Level floor_level = doc.GetElement(floorList[j].LevelId) as Level;
                            excelWorksheet.Cells[address.Item1 + j + 1, address.Item2] = floor_level.Name.Split('-')[1];
                            excelWorksheet.Cells[address.Item1 + j + 1, address.Item2 + 1] = (-(floor_level.LookupParameter("高程").AsDouble() * 304.8 - (floorList[j].LookupParameter("厚度").AsDouble() * 304.8) / 2) / 1000).ToString();
                            excelWorksheet.Cells[address.Item1 + j + 1, address.Item2 + 2] = floorList[j].LookupParameter("厚度").AsDouble() * 304.8 / 1000;
                            excelWorksheet.Cells[address.Item1 + j + 1, address.Item2 + 3] = floorList[j].get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();
                        }

                        TaskDialog.Show("message", "完成匯出" + sectionName);

                    }
                    else if (excavationType == "井式基礎")
                    {
                        //TaskDialog.Show("message", excavationType);
                        //去掉非開挖階數之樓層
                        for (int i = levelList.Count - 1; i >= 0; i--)
                        {
                            if (!levelList[i].Name.Contains("開挖階數"))
                            {
                                levelList.RemoveAt(i);
                            }
                        }

                        //開頭資料
                        double.TryParse(wallList[0].LookupParameter("不連續高度").AsValueString(), out depth);
                        double.TryParse((wallList[0] as Wall).WallType.LookupParameter("寬度").AsValueString(), out width);
                        double.TryParse(levelList[levelList.Count - 1].LookupParameter("高程").AsValueString(), out height);
                        double.TryParse(levelList[0].LookupParameter("高程").AsValueString(), out horizon);
                        excel.SetData(filepath, 1);

                        Tuple<int, int> address = excel.FindAddress("擋土壁長度：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = depth / 1000;
                        address = excel.FindAddress("擋土壁型式：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = wallList[0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[1];
                        address = excel.FindAddress("分析斷面：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = wallList[0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[0];
                        address = excel.FindAddress("擋土壁厚度：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = width / 1000;
                        address = excel.FindAddress("開挖深度：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = -height / 1000;
                        address = excel.FindAddress("地表高程：");
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = -horizon / 1000;

                        //開挖階數
                        //找開挖階數位置 寫入開挖街樹
                        address = Tuple.Create(address.Item1 + 2, address.Item2 + 1);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "開挖階數:";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "階數";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "深度";


                        for (int i = 0; i < levelList.Count - 1; i++)
                        {
                            double.TryParse(levelList[i].LookupParameter("高程").AsValueString(), out stepDepth);
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2] = i + 1;
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 1] = -stepDepth / 1000;
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 2] = "m";
                        }

                        address = Tuple.Create(address.Item1 + levelList.Count + 2, address.Item2);

                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "支撐階數:";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "階數";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "深度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 2] = "支數";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 3] = "型號";

                        IList<IList<Element>> frameListbyDepth = new List<IList<Element>>();

                        //TaskDialog.Show("1", levelList.Count.ToString());
                        //把支撐依照不同高度分類
                        int supportFrameLevel = 0;
                        foreach (Element frame in frameList)
                        {
                            bool alreadyInList = false;
                            int.TryParse(frame.LookupParameter("備註").AsString().Split('-')[1], out supportFrameLevel);
                            for (int j = 0; j < frameListbyDepth.Count; j++)
                            {
                                if (supportFrameLevel.ToString() == frameListbyDepth[j][0].LookupParameter("備註").AsString().Split('-')[1])
                                {
                                    frameListbyDepth[j].Add(frame);
                                    alreadyInList = true;
                                    break;
                                }
                            }
                            if (alreadyInList == false)
                            {
                                frameListbyDepth.Add(new List<Element>());
                                frameListbyDepth[frameListbyDepth.Count - 1].Add(frame);
                            }

                        }

                        double frameDepth = 0;
                        for (int i = 0; i < frameListbyDepth.Count; i++)
                        {
                            double.TryParse(frameListbyDepth[i][0].get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).AsValueString(), out frameDepth);
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2] = frameListbyDepth[i][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[1];
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 1] = frameDepth / 1000;
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 2] = frameListbyDepth[i].Count / 2;
                            excelWorksheet.Cells[address.Item1 + i + 1, address.Item2 + 3] = frameListbyDepth[i][0].Name;
                        }

                        address = Tuple.Create(address.Item1 + frameListbyDepth.Count + 2, address.Item2);
                        excelWorksheet.Cells[address.Item1, address.Item2 - 1] = "樓板回築&回填：";
                        excelWorksheet.Cells[address.Item1, address.Item2] = "樓層數";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 1] = "深度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 2] = "厚度";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 3] = "直徑";
                        excelWorksheet.Cells[address.Item1, address.Item2 + 4] = "混凝土強度(kg/cm2)";
                        double floorDepth = 0;
                        double floorthickness = 0;
                        double floorDiameter = 0;
                        for (int i = 0; i < floorList.Count; i++)
                        {
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2] = floorList[i].Name;
                            double.TryParse(floorList[i].get_Parameter(BuiltInParameter.STRUCTURAL_ELEVATION_AT_BOTTOM).AsValueString(), out floorDepth);
                            double.TryParse(floorList[i].get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsValueString(), out floorthickness);
                            double.TryParse(floorList[i].get_Parameter(BuiltInParameter.HOST_PERIMETER_COMPUTED).AsValueString(), out floorDiameter);
                            floorDiameter = floorDiameter / Math.PI / 1000;
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2] = floorList[i].Name;
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2 + 1] = -floorDepth / 1000;
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2 + 2] = floorthickness / 1000;
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2 + 3] = String.Format("{0:0.##}", floorDiameter);
                            excelWorksheet.Cells[address.Item1 + 1 + i, address.Item2 + 4] = floorList[i].get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString();
                        }
                        TaskDialog.Show("message", "完成匯出" + sectionName);
                    }

                    //save
                    excelWorkbook.SaveAs(Filename: wallList[0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[0]);

                    //關閉即釋放物件
                    excelWorksheet = null;
                    excelWorkbook.Close();
                    excelWorkbook = null;
                    excelAPP.Quit();
                    excelAPP = null;
                    excel.CloseEx();
                    //TaskDialog.Show("message", "finish once!!");
                }
                catch
                {
                    //關閉即釋放物件
                    excelWorksheet = null;
                    excelWorkbook.Close();
                    excelWorkbook = null;
                    excelAPP.Quit();
                    excelAPP = null;
                    excel.CloseEx();
                    //TaskDialog.Show("message", "finish once!!");}

                }
            }
                TaskDialog.Show("通知", "完成匯出");
            //throw new NotImplementedException();
        }

        IList<IList<Element>> dividedBySectionForColumn(IList<Element> List)
        {
            IList<IList<Element>> listBySection = new List<IList<Element>>();
            for (int i = 0; i < List.Count; i++)
            {
                string sectionName = List[i].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0];
                if (i == 0)
                {
                    listBySection.Add(new List<Element>());
                    listBySection[0].Add(List[i]);
                }
                else
                {
                    bool isExist = false;
                    for (int j = 0; j < listBySection.Count; j++)
                    {
                        string sectionNameTemp = listBySection[j][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0];
                        if (sectionName == sectionNameTemp)
                        {
                            listBySection[j].Add(List[i]);
                            isExist = true;
                            break;
                        }
                    }
                    if (isExist == false)
                    {
                        listBySection.Add(new List<Element>());
                        listBySection[listBySection.Count - 1].Add(List[i]);
                    }
                }
            }

            return listBySection;
        }

        IList<IList<Element>> dividedBySectionForLevel(IList<Element> List)
        {
            IList<IList<Element>> listBySection = new List<IList<Element>>();
            for (int i = 0; i < List.Count; i++)
            {
                string sectionName = List[i].Name.Split('-')[0];
                if (i == 0)
                {
                    listBySection.Add(new List<Element>());
                    listBySection[0].Add(List[i]);
                }
                else
                {
                    bool isExist = false;
                    for (int j = 0; j < listBySection.Count; j++)
                    {
                        string sectionNameTemp = listBySection[j][0].Name.Split('-')[0];
                        if (sectionName == sectionNameTemp)
                        {
                            listBySection[j].Add(List[i]);
                            isExist = true;
                            break;
                        }
                    }
                    if (isExist == false)
                    {
                        listBySection.Add(new List<Element>());
                        listBySection[listBySection.Count - 1].Add(List[i]);
                    }
                }
            }

            return listBySection;
        }

        IList<IList<Element>> dividedBySection(IList<Element> List)
        {
            IList<IList<Element>> listBySection = new List<IList<Element>>();
            for (int i = 0; i < List.Count; i++)
            {
                string sectionName = List[i].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[0];
                if (i == 0)
                {
                    listBySection.Add(new List<Element>());
                    listBySection[0].Add(List[i]);
                }
                else
                {
                    bool isExist = false;
                    for (int j = 0; j < listBySection.Count; j++)
                    {
                        string sectionNameTemp = listBySection[j][0].get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split('-')[0];
                        if(sectionName == sectionNameTemp)
                        {
                            listBySection[j].Add(List[i]);
                            isExist = true;
                            break;
                        }

                    }
                    if (isExist == false)
                    {
                        listBySection.Add(new List<Element>());
                        listBySection[listBySection.Count - 1].Add(List[i]);
                    }
                }
            }

            return listBySection;
        }

        public IList<Curve> curveIntersection(IList<Curve> curveList)
        {
            IList<Curve> cList = new List<Curve>();
            XYZ start, end;
            
            for(int i = 0; i < curveList.Count; i++)
            {
                if(i == 0)
                {
                    start = intersectionPoint(curveList[curveList.Count - 1], curveList[i]);
                    end = intersectionPoint(curveList[i], curveList[i + 1]);
                }
                else if(i == curveList.Count - 1)
                {
                    start = intersectionPoint(curveList[i - 1], curveList[i]);
                    end = intersectionPoint(curveList[i], curveList[0]);
                }
                else
                {
                    start = intersectionPoint(curveList[i - 1], curveList[i]);
                    end = intersectionPoint(curveList[i], curveList[i + 1]);
                }
                cList.Add(Line.CreateBound(start, end));
            }
            return cList;
        }

        public XYZ intersectionPoint(Curve C1, Curve C2)
        {
            //ax + by = c
            Line L1 = C1 as Line;
            Line L2 = C2 as Line;
            double a1, b1, c1, a2, b2, c2, x1, y1, x2, y2, dx, dy, d;
            a1 = L1.Direction.Y;
            b1 = -L1.Direction.X;
            a2 = L2.Direction.Y;
            b2 = -L2.Direction.X;
            x1 = L1.GetEndPoint(0).X;
            y1 = L1.GetEndPoint(0).Y;
            x2 = L2.GetEndPoint(0).X;
            y2 = L2.GetEndPoint(0).Y;
            c1 = a1 * x1 + b1 * y1;
            c2 = a2 * x2 + b2 * y2;
            d = a1 * b2 - b1 * a2;
            dx = c1 * b2 - b1 * c2;
            dy = a1 * c2 - c1 * a2;
            XYZ interPoint;
            if (d != 0)
            {
                interPoint = new XYZ(dx / d, dy / d, C1.GetEndPoint(0).Z);
            }
            else
            {
                double distance = distanceBetweenTwoPoints(C1.GetEndPoint(0), C2.GetEndPoint(0));
                interPoint = (C1.GetEndPoint(0) + C2.GetEndPoint(0)) / 2;
                if (distanceBetweenTwoPoints(C1.GetEndPoint(0), C2.GetEndPoint(1)) < distance)
                {
                    distance = distanceBetweenTwoPoints(C1.GetEndPoint(0), C2.GetEndPoint(1));
                    interPoint = (C1.GetEndPoint(0) + C2.GetEndPoint(1)) / 2;
                    if (distanceBetweenTwoPoints(C1.GetEndPoint(1), C2.GetEndPoint(1)) < distance)
                    {
                        distance = distanceBetweenTwoPoints(C1.GetEndPoint(1), C2.GetEndPoint(1));
                        interPoint = (C1.GetEndPoint(1) + C2.GetEndPoint(1)) / 2;
                    }
                }
                else
                {
                    if (distanceBetweenTwoPoints(C1.GetEndPoint(1), C2.GetEndPoint(0)) < distance)
                    {
                        distance = distanceBetweenTwoPoints(C1.GetEndPoint(1), C2.GetEndPoint(0));
                        interPoint = (C1.GetEndPoint(1) + C2.GetEndPoint(0)) / 2;
                    }
                }
            }
            return interPoint;
        }

        public IList<BoundarySegment> lineArrange(IList<BoundarySegment> wallList)
        {
            IList<BoundarySegment> arrangedWall = new List<BoundarySegment>();
            XYZ start;
            XYZ end;
            XYZ target = wallList[0].GetCurve().GetEndPoint(1);
            arrangedWall.Add(wallList[0]);
            wallList.RemoveAt(0);
            int size = wallList.Count;
            while (size+1 > arrangedWall.Count)
            {
                for(int j = 0; j < wallList.Count; j++)
                {
                    start = wallList[j].GetCurve().GetEndPoint(0);
                    end = wallList[j].GetCurve().GetEndPoint(1);
                    if (distanceBetweenTwoPoints(target, start) < 0.01)
                    {
                        arrangedWall.Add(wallList[j]);
                        target = end;
                        wallList.RemoveAt(j);
                        break;
                    }else if((distanceBetweenTwoPoints(target, end) < 0.01)){
                        arrangedWall.Add(wallList[j]);
                        target = start;
                        wallList.RemoveAt(j);
                        break;
                    }
                }
            }
            return arrangedWall;
        }

        
        public double distanceBetweenLines(Line L1, Line L2)
        {
            // y = ax + c
            double a = 0;
            double b = 1;
            if (L1.Direction.X == 0)
            {
                a = -1;
                b = 0;
            }
            else
            {
                a = L1.Direction.Y/L1.Direction.X;
            }
            
            XYZ P1 = L1.GetEndPoint(0);
            XYZ P2 = L2.GetEndPoint(0);
            double c1 = b*P1.Y - a * P1.X;
            double c2 = b*P2.Y - a * P2.X;
            double distance = Math.Abs(c1 - c2)/Math.Pow((Math.Pow(a,2) + b),0.5);
            return distance;
        }

        public double distanceBetweenTwoPoints(XYZ p1, XYZ p2)
        {
            double distance = Math.Pow(Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y,2),0.5);
            return distance;
        }

        public bool isParallel(Line L1, Line L2)
        {
            if (L1.Direction.IsAlmostEqualTo(L2.Direction, 0.001) || L1.Direction.IsAlmostEqualTo(-L2.Direction, 0.001))
            {
                return true;
            }
            return false;
        }

        public string GetName()
        {
            return "";
        }
        
    }
}
