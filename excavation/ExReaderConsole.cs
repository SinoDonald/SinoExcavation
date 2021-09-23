using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

using Autodesk.Revit.UI;


namespace ExReaderConsole
{
    class ExReader
    {

        string exlocatiion;
        Excel.Application xlApp;
        Excel.Workbooks xlWorkbooks;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Excel.Range xlRange;

        int rowCount;
        int colCount;
        int startRow;
        int endRow;

        //DE&Circle part
        public string section;
        public List<Tuple<double, double>> excaRange = new List<Tuple<double, double>>();   // X Y
        public List<Tuple<int, double>> excaLevel = new List<Tuple<int, double>>();   // 階數 深度
        public List<Tuple<string, double, int, string, int>> supLevel = new List<Tuple<string, double, int, string, int>>();
        public List<Tuple<string, int, string>> beamLevel = new List<Tuple<string, int, string>>();
        public List<Tuple<double, double, double, double, string, double>> column = new List<Tuple<double, double, double, double, string, double>>();
        public List<Tuple<string, double, double>> sideWall = new List<Tuple<string, double, double>>();
        public List<Tuple<string, double, double, string>> back = new List<Tuple<string, double, double, string>>();
        public List<Tuple<string, double, double, double, double>> circle_floor = new List<Tuple<string, double, double, double, double>>();
        public List<Tuple<double, double, string>> timber_lagging = new List<Tuple<double, double, string>>();
        public List<Tuple<double, double, double, double>> soldier_pile = new List<Tuple<double, double, double, double>>();
        public List<Tuple<double, double, double, double, double, double, double>> rail_soldier_pile = new List<Tuple<double, double, double, double, double, double, double>>();
        public List<Tuple<string, double, double>> unit_text = new List<Tuple<string, double, double>>();
        public List<double> centralCol = new List<double>();
        public double wall_width;
        public double wall_high;
        public double Diameter;
        public string First_single = "";
        public string First_double = "";

        public double protection_width;
        public List<Tuple<double, double, string, double, string, double>> vertical_r_rebar = new List<Tuple<double, double, string, double, string, double>>();
        public List<Tuple<double, double, string, double, string, double>> vertical_e_rebar = new List<Tuple<double, double, string, double, string, double>>();
        public List<Tuple<double, double, string, double>> horizontal_rebar = new List<Tuple<double, double, string, double>>();
        public List<Tuple<double, double>> shear_rebar_depth = new List<Tuple<double, double>>();
        public List<Tuple<string, string, double, string, string, double>> shear_rebar = new List<Tuple<string, string, double, string, string, double>>();

        public List<double> monitor_double = new List<double>();
        public List<string> monitor_string = new List<string>();

        //MH part
        public List<List<string>> MHdata = new List<List<string>>();

        public ExReader()
        {

        }

        public void SetData(string file, int page)
        {
            exlocatiion = Path.Combine(String.Format(@"{0}", file));

            xlApp = new Excel.Application();
            xlWorkbooks = xlApp.Workbooks;
            xlWorkbook = xlWorkbooks.Open(exlocatiion);

            xlWorksheet = xlWorkbook.Sheets[page];
            xlRange = xlWorksheet.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
        }

        void SetPage(int page)
        {

            xlWorksheet = xlWorkbook.Sheets[page];
            xlRange = xlWorksheet.UsedRange;
            rowCount = xlRange.Rows.Count;
            colCount = xlRange.Columns.Count;
        }

        public void PassCircle()
        {
            var pos = this.FindAddress("擋土壁厚度");
            wall_width = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;

            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();

            pos = this.FindAddress("擋土壁長度");
            wall_high = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;
            //pos = this.FindAddress("環形擋土直徑");
            //Diameter = xlRange.Cells[pos.Item1 + 1, pos.Item2].Value2;
            pos = this.FindAddress("開挖範圍");
            int i = 1;
            try
            {
                do
                {
                    var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                    excaRange.Add(data);
                    i++;
                } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            }
            catch { centralCol.Add(4); }
            pos = this.FindAddress("開挖階數");
            i = 1;
            do
            {
                var data = Tuple.Create((int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                excaLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);


            pos = this.FindAddress("支撐階數");
            i = 1;
            do
            {
                var data = Tuple.Create((string)(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2.ToString()), xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2.ToString(),
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2);
                supLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);


            pos = this.FindAddress("樓板回築&回填");
            i = 1;
            do
            {
                var data = Tuple.Create((string)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2);
                circle_floor.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            SetPage(1);
            xlWorkbook.Close(0);
            xlWorkbooks.Close();

        }
        public void PassWallData()
        {
            var pos = this.FindAddress("擋土壁厚度");
            wall_width = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;

            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();
            pos = this.FindAddress("擋土壁長度");
            wall_high = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;

            int i = 0;
            if (excaRange.Count == 0)
            {
                try
                {

                    pos = this.FindAddress("開挖範圍");
                    i = 1;
                    do
                    {
                        var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                        excaRange.Add(data);
                        i++;
                    } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
                }
                catch { centralCol.Add(4); }
            }

            pos = this.FindAddress("開挖階數");
            i = 1;
            do
            {
                var data = Tuple.Create((int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                excaLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            SetPage(1);
        }

        public void PassColumnData()
        {

            var pos = this.FindAddress("中間柱");
            int i = 1;

            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 6].Value2);
                column.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();

            if (excaRange.Count == 0)
            {
                try
                {
                    pos = this.FindAddress("開挖範圍");
                    i = 1;
                    do
                    {
                        var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                        excaRange.Add(data);
                        i++;
                    } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
                    centralCol.Add(4);

                }
                catch { centralCol.Add(4); }
            }
        }
        public void PassSoldierPile()
        {
            var pos = this.FindAddress(column[0].Item5);
            var data = Tuple.Create(xlRange.Cells[pos.Item1, pos.Item2 + 8].Value2, xlRange.Cells[pos.Item1, pos.Item2 + 9].Value2,
                xlRange.Cells[pos.Item1, pos.Item2 + 10].Value2, xlRange.Cells[pos.Item1, pos.Item2 + 11].Value2);
            soldier_pile.Add(data);
        }
        public void PassRailSoldierPile()
        {
            var pos = this.FindAddress("JRS˙JIS37Kg A");
            var data = Tuple.Create(xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1, pos.Item2 + 2].Value2,
                xlRange.Cells[pos.Item1, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1, pos.Item2 + 4].Value2,
                xlRange.Cells[pos.Item1, pos.Item2 + 5].Value2, xlRange.Cells[pos.Item1, pos.Item2 + 6].Value2,
                xlRange.Cells[pos.Item1, pos.Item2 + 7].Value2);
            rail_soldier_pile.Add(data);
        }
        public void PassTimberLagging()
        {
            var pos = this.FindAddress("長(m)");
            var data = Tuple.Create(xlRange.Cells[pos.Item1 + 1, pos.Item2 + 1].Value2,
                xlRange.Cells[pos.Item1 + 1, pos.Item2 + 2].Value2,
                xlRange.Cells[pos.Item1 + 1, pos.Item2 + 3].Value2);
            timber_lagging.Add(data);
        }
        public void PassUnitText()
        {
            int i = 0;
            var pos = Tuple.Create(1, 1);
            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, 
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                unit_text.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2].Value2 != null);
        }
        public void PassFrameData()
        {
            var pos = this.FindAddress("支撐階數");
            int i = 1;
            do
            {
                var data = Tuple.Create((string)(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2.ToString()), xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2.ToString(),
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2);
                supLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            //Find the first single and double
            foreach (Tuple<string, double, int, string, int> tuple in supLevel)
            {
                if (tuple.Item3 == 1 && First_single == "")
                {
                    First_single = tuple.Item1;
                }
                else if (tuple.Item3 == 2 && First_double == "")
                {
                    First_double = tuple.Item1;
                }
            }

            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();
        }
        public void PassBeamData()
        {
            var pos = this.FindAddress("圍囹階數");
            int i = 1;
            do
            {
                var data = Tuple.Create((string)(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2.ToString()), (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2.ToString());
                beamLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();
            try
            {
                pos = this.FindAddress("開挖範圍");
                i = 1;
                do
                {
                    var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                    excaRange.Add(data);
                    i++;
                } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            }
            catch { centralCol.Add(4); };
        }
        public void PassSideData()
        {
            var pos = this.FindAddress("支撐階數");
            int i = 1;
            do
            {
                var data = Tuple.Create((string)(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2.ToString()), xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2.ToString(),
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2);
                supLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            if (this.FindAddress("側牆") != null)
            {
                pos = this.FindAddress("側牆");
                i = 1;
                do
                {
                    var data = Tuple.Create((string)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                        (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2);
                    sideWall.Add(data);
                    i++;
                } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            }


            if (this.FindAddress("樓板回築&回填") != null)
            {
                pos = this.FindAddress("樓板回築&回填");
                i = 1;
                do
                {
                    var data = Tuple.Create((string)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                        (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, (string)xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2.ToString());
                    back.Add(data);
                    i++;
                } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            }

            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();
        }

        public void PassDE()
        {
            var pos = this.FindAddress("擋土壁厚度");
            wall_width = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;
            pos = this.FindAddress("擋土壁長度");
            wall_high = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;

            pos = this.FindAddress("開挖範圍");
            int i = 1;
            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                excaRange.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            pos = this.FindAddress("開挖階數");
            i = 1;
            do
            {
                var data = Tuple.Create((int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                excaLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);


            pos = this.FindAddress("支撐階數");
            i = 1;
            do
            {
                var data = Tuple.Create((string)(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2.ToString()), xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2.ToString(),
                    (int)xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2);
                supLevel.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            pos = this.FindAddress("中間柱");
            i = 1;
            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2.Tostring(), xlRange.Cells[pos.Item1 + i, pos.Item2 + 6].Value2);
                column.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            /*
            if (this.FindAddress("側牆") != null)
            {
                pos = this.FindAddress("側牆");
                i = 1;
                do
                {
                    var data = Tuple.Create((string)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                        (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2);
                    sideWall.Add(data);
                    i++;
                } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            }
            
            
            if (this.FindAddress("樓板回築&回填") != null)
            {
                pos = this.FindAddress("樓板回築&回填");
                i = 1;
                do
                {
                    var data = Tuple.Create((string)xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                        (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, (double)xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2);
                    back.Add(data);
                    i++;
                } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            }*/
            /*
            SetPage(2);
            pos = this.FindAddress("中間樁");
            for(int j = 1; j != 7; j++)
            {
                double d = xlRange.Cells[pos.Item1 + 1, pos.Item2 + j].Value2;
                centralCol.Add(d);
            }*/
            //0608_test
            centralCol.Add(4);
            centralCol.Add(4);
            centralCol.Add(4);
            centralCol.Add(4);
            centralCol.Add(4);
            SetPage(1);
            xlWorkbook.Close(0);
            xlWorkbooks.Close();

        }


        // 鋼筋配筋
        public void PassRebarData()
        {
            var pos = this.FindAddress("擋土壁厚度");
            wall_width = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;
            pos = this.FindAddress("分析斷面");
            section = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2.ToString();
            pos = this.FindAddress("擋土壁長度");
            wall_high = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;
            pos = this.FindAddress("保護層厚度");
            protection_width = xlRange.Cells[pos.Item1, pos.Item2 + 1].Value2;

            // 垂直筋擋土側
            int i = 1;
              
            pos = this.FindAddress("擋土側");
            i = 2;
            do 
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 6].Value2);
                vertical_r_rebar.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            
            // 垂直筋開挖側
            pos = this.FindAddress("開挖側");
            i = 2;
            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 6].Value2);
                vertical_e_rebar.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            //水平筋
            pos = this.FindAddress("擋土牆水平筋設計");
            i = 2;
            do
            {
                var data = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2);
                horizontal_rebar.Add(data);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);

            
            //剪力筋
            pos = this.FindAddress("擋土牆剪力筋設計");
            i = 2;
            do
            {
                var data1 = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 2].Value2);
                shear_rebar_depth.Add(data1);

                var data2 = Tuple.Create(xlRange.Cells[pos.Item1 + i, pos.Item2 + 3].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 4].Value2, 
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 5].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 6].Value2,
                    xlRange.Cells[pos.Item1 + i, pos.Item2 + 7].Value2, xlRange.Cells[pos.Item1 + i, pos.Item2 + 8].Value2);
                shear_rebar.Add(data2);

                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2 + 1].Value2 != null);
            
            SetPage(1);
        }

        public List<double> PassMonitorDouble(string column_name, int i)
        {
            monitor_double.Clear();

            var pos = this.FindAddress(column_name);
            do
            {
                monitor_double.Add(xlRange.Cells[pos.Item1 + i, pos.Item2].Value2);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2].Value2 != null);

            return monitor_double;
        }

        public List<string> PassMonitortString(string column_name, int i = 3)
        {
            monitor_string.Clear();

            var pos = this.FindAddress(column_name);
            do
            {
                monitor_string.Add(xlRange.Cells[pos.Item1 + i, pos.Item2].Value2);
                i++;
            } while (xlRange.Cells[pos.Item1 + i, pos.Item2].Value2 != null);

            return monitor_string;
        }



        public Tuple<int, int> FindAddress(string name)
        {
            Excel.Range address;
            address = xlRange.Find(name, MatchCase: true);
            if (address == null)
            {
                return null;
            }
            else
            {
                var pos = Tuple.Create(address.Row, address.Column);
                return pos;
            }

        }


        public void CloseEx()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            /*Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlWorkbooks);*/
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

        }


    }


}
