using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace SinoExcavation_2020
{
    public partial class Form3 : Form
    {
        public Form3(DataTable dt)//, List<double> alert, int left_boundary, int right_boundary)
        {
            InitializeComponent();
            chart1.DataSource = dt;

            chart1.Series["Series1"].XValueMember = "distance";
            chart1.Series["Series1"].YValueMembers = "settlement";
            chart1.Series["Series1"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.ChartAreas[0].AxisY.LabelStyle.Format = "";
            chart1.ChartAreas[0].AxisX.Title = "distance";
            chart1.ChartAreas[0].AxisY.Title = "settlement";
            /* 
            double left = chart1.ChartAreas[0].AxisX.ValueToPixelPosition(left_boundary);
            double right = chart1.ChartAreas[0].AxisX.ValueToPixelPosition(right_boundary);
            double a = chart1.ChartAreas[0].AxisX.ValueToPixelPosition(alert[0]);
            double b = chart1.ChartAreas[0].AxisX.ValueToPixelPosition(alert[1]);
            double c = chart1.ChartAreas[0].AxisX.ValueToPixelPosition(alert[2]);
            double d = chart1.ChartAreas[0].AxisX.ValueToPixelPosition(alert[3]);
            */
            StripLine sl1 = new StripLine();
            sl1.Interval = 0;
            sl1.StripWidth = 100;
            sl1.IntervalOffset = 0;
            sl1.BackColor = Color.FromArgb(64, Color.LightSalmon);
            chart1.ChartAreas[0].AxisY.StripLines.Add(sl1);

            //https://stackoverflow.com/questions/49775533/line-chart-with-different-colors-on-a-different-interval-at-x-axis-fill-in-same
        }
    }
}
