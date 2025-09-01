using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Runtime.InteropServices;

namespace SinoExcavation_2025
{
    public partial class Form4 : System.Windows.Forms.Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
       (
           int nLeftRect,     // x-coordinate of upper-left corner
           int nTopRect,      // y-coordinate of upper-left corner
           int nRightRect,    // x-coordinate of lower-right corner
           int nBottomRect,   // y-coordinate of lower-right corner
           int nWidthEllipse, // height of ellipse
           int nHeightEllipse // width of ellipse

       );

        public IList<string> dect
        {
            get;
            set;
        }
        public form2 Read_form
        {
            get;
            set;
        }
        //在此階段，我們將各個建置的指令利用ExternalEvent的類型建立起來
        //要使用的時候再在各個按鍵去觸發他即可

        ExternalEvent externalEvent_CreateWall;
        CreateWall handler_createWall = new CreateWall();

        ExternalEvent externalEvent_CreateWallCADmode;
        CreateWallCADmode handler_createWallCADmode = new CreateWallCADmode();

        ExternalEvent externalEvent_CreateSheetPile;
        sheet_pile_NoCAD handler_sheet_pile_NoCAD = new sheet_pile_NoCAD();

        ExternalEvent externalEvent_CreateSoldierPile;
        CreateSoldierPile handler_CreateSoldierPile = new CreateSoldierPile();

        ExternalEvent externalEvent_CreateRebar;
        CreateRebar handler_CreateRebar = new CreateRebar();

        ExternalEvent externalEvent_CreateMonitor;
        CreateMonitor handler_CreateMonitor = new CreateMonitor();

        ExternalEvent externalEvent_VisualizeMonitor;
        VisualizeMonitor handler_VisualizeMonitor = new VisualizeMonitor();

        ExternalEvent externalEvent_CreateUnit;
        CreateUnit handler_CreateUnit = new CreateUnit();

        ExternalEvent externalEvent_CreateColumn;
        CreateColumn handler_createColumn = new CreateColumn();

        ExternalEvent externalEvent_MoveColumn;
        MoveColumn handler_moveColumn = new MoveColumn();

        ExternalEvent externalEvent_CreateFrame;
        CreateFrame handler_createFrame = new CreateFrame();

        ExternalEvent externalEvent_CreateBeam;
        CreateBeam handler_createBeam = new CreateBeam();

        ExternalEvent ExternalEvent_SingleSlopeFrame;
        SingleSlopeFrame handler_singleslopeframe = new SingleSlopeFrame();

        ExternalEvent externalEvent_CreateSlope;
        CreateSlope handler_createSlope = new CreateSlope();

        ExternalEvent externalEvent_CreateSide;
        CreateSide handler_createSide = new CreateSide();

        ExternalEvent externalEvent_circle_excavation;
        Circle_excavation handler_circle_excavation = new Circle_excavation();

        ExternalEvent externalEvent_make_sheet;
        MakeSheet handler_make_sheet = new MakeSheet();

        ExternalEvent externalEvent_PutDetect;
        PutDetect handler_putdetect = new PutDetect();


        ExternalEvent externalEvent_DetecLevel;
        DetecLevel handler_deteclevel = new DetecLevel();

        ExternalEvent ExternalEvent_counting;
        Counting handler_counting = new Counting();

        ExternalEvent ExternalEvent_output_excel;
        OutputExcel handler_output_excel = new OutputExcel();

        ExternalEvent ExternalEvent_DrawJack;
        DrawJack handler_drawjack = new DrawJack();

        ExternalEvent ExternalEvent_DrawChannelSteel;
        DrawChannelSteel handler_draw_channelsteel = new DrawChannelSteel();

        ExternalEvent ExternalEvent_CreateTriSupportofBeam;
        CreateTriSupportofBeam handler_create_trisupport_beam = new CreateTriSupportofBeam();

        ExternalEvent ExternalEvent_CreateTriSupportOfFrame;
        CreateTriSupportOfFrame handler_create_trisupport_frame = new CreateTriSupportOfFrame();

        ExternalEvent externalEvent_CreateBentPile;
        CreateBentPile handler_CreateBentPile = new CreateBentPile();


        ExternalEvent ExternalEvent_CreateU;
        CreateU handler_create_U = new CreateU();

        public Form4(UIDocument uIDocument, form2 form2)
        {
            //初始化視窗
            InitializeComponent();

            //讀取圖層視窗
            Read_form = form2;

            //限制視窗外觀
            this.FormBorderStyle = FormBorderStyle.None;
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            //將事件與指令結合，好讓我們之後利用ExternalEvent的指令觸發
            externalEvent_CreateWall = ExternalEvent.Create(handler_createWall);
            externalEvent_CreateWallCADmode = ExternalEvent.Create(handler_createWallCADmode);
            externalEvent_CreateSheetPile = ExternalEvent.Create(handler_sheet_pile_NoCAD);
            externalEvent_CreateSoldierPile = ExternalEvent.Create(handler_CreateSoldierPile);
            externalEvent_CreateRebar = ExternalEvent.Create(handler_CreateRebar);
            externalEvent_CreateMonitor = ExternalEvent.Create(handler_CreateMonitor);
            externalEvent_VisualizeMonitor = ExternalEvent.Create(handler_VisualizeMonitor);
            externalEvent_CreateUnit = ExternalEvent.Create(handler_CreateUnit);
            externalEvent_CreateColumn = ExternalEvent.Create(handler_createColumn);
            externalEvent_MoveColumn = ExternalEvent.Create(handler_moveColumn);
            externalEvent_CreateFrame = ExternalEvent.Create(handler_createFrame);
            ExternalEvent_SingleSlopeFrame = ExternalEvent.Create(handler_singleslopeframe);
            externalEvent_CreateBeam = ExternalEvent.Create(handler_createBeam);
            externalEvent_CreateSlope = ExternalEvent.Create(handler_createSlope);
            externalEvent_CreateSide = ExternalEvent.Create(handler_createSide);
            externalEvent_circle_excavation = ExternalEvent.Create(handler_circle_excavation);
            externalEvent_make_sheet = ExternalEvent.Create(handler_make_sheet);
            externalEvent_PutDetect = ExternalEvent.Create(handler_putdetect);
            externalEvent_DetecLevel = ExternalEvent.Create(handler_deteclevel);
            ExternalEvent_counting = ExternalEvent.Create(handler_counting);
            ExternalEvent_output_excel = ExternalEvent.Create(handler_output_excel);
            ExternalEvent_DrawJack = ExternalEvent.Create(handler_drawjack);
            ExternalEvent_DrawChannelSteel = ExternalEvent.Create(handler_draw_channelsteel);
            ExternalEvent_CreateTriSupportofBeam = ExternalEvent.Create(handler_create_trisupport_beam);
            ExternalEvent_CreateTriSupportOfFrame = ExternalEvent.Create(handler_create_trisupport_frame);
            externalEvent_CreateBentPile = ExternalEvent.Create(handler_CreateBentPile);
            ExternalEvent_CreateU = ExternalEvent.Create(handler_create_U);


            //剖面數量
            comboBox7.Items.Add(1);
            comboBox7.Items.Add(2);
            comboBox7.Items.Add(4);
            comboBox7.Items.Add(6);

            //偏移量&中間樁偏移
            cir_shift_x.Text = "0";
            cir_shift_y.Text = "0";
            shift_x.Text = "0";
            shift_y.Text = "0";
            textBox2.Text = "0";
            textBox3.Text = "0";
        }

        private void Form1_Load(object sender, EventArgs e)
        //介面內容包裝
        {// 找出字體大小,並算出比例
            float dpiX, dpiY;
            Graphics graphics = this.CreateGraphics();
            dpiX = graphics.DpiX;
            dpiY = graphics.DpiY;
            int intPercent = (dpiX == 96) ? 100 : (dpiX == 120) ? 125 : 150;

            // 針對字體變更Form的大小
            this.Height = this.Height * intPercent / 100;
            this.Width = this.Width * intPercent / 100;
            this.Size = new System.Drawing.Size(this.header.Size.Width, this.header.Size.Height + this.homeleftpanel.Size.Height + this.panel3.Height);

            //將基礎介面之大小縮為(600,385)
            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));
            //將基礎介面之四個角落進行半徑為20之圓滑處理

            int start_position = leftpanel.Location.X + leftpanel.Size.Width;
            int start_position_y = panel3.Location.Y + panel3.Height;
            tabControl1.SizeMode = TabSizeMode.Fixed;
            tabControl1.Appearance = TabAppearance.FlatButtons;
            tabControl1.ItemSize = new Size(0, 1);

            main.Location = new System.Drawing.Point(start_position + 5, start_position_y);
            //上傳檔案初始介面位置移到(start_position,125)
            output1.Location = new System.Drawing.Point(start_position, start_position_y);
            //建立深開挖基礎介面位置移到(start_position,102) 對初始Tab頁籤進行遮擋
            output2.Location = new System.Drawing.Point(start_position + 5, start_position_y);
            //建立井式基礎介面位置移到(start_position,125)
            output3.Location = new System.Drawing.Point(start_position, start_position_y - 23);
            //圖資產出介面位置移到(start_position,125)




        }

        private void button1_Click(object sender, EventArgs e)
        {
            //讀取斷面資料
            upload_textbox.Text = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            file_names = openFileDialog.FileNames;
            upload_textbox.Text = String.Format("已選取 {0} 個開挖面", file_names.Count());

            button1.Show();
            button10.Show();

        }



        private void SetReturnDegree(double degree)
        {
            //CAD模式的連續壁角度委派方法
            textBox7.Text = degree.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //判斷是否有載入CAD，使用不同建置方法
            if (textBox1.Text.Contains("dwg") && comboBox10.Text == "連續壁")
            {
                
                //建立連續壁CADmode
                handler_createWallCADmode.files_path = new List<string>();
                handler_createWallCADmode.files_path = file_names;
                handler_createWallCADmode.xy_shift = new List<double>();
                handler_createWallCADmode.xy_shift.Add(double.Parse(shift_x.Text));
                handler_createWallCADmode.xy_shift.Add(double.Parse(shift_y.Text));
                externalEvent_CreateWallCADmode.Raise();
                handler_createWallCADmode.ReturnDegreeCallback += new CreateWallCADmode.ReturnDegree(this.SetReturnDegree);//委派方法給這個事件
            }
            else
            {
                if(comboBox10.Text == "連續壁")
                {
                    //建立連續壁
                    handler_createWall.files_path = new List<string>();
                    handler_createWall.files_path = file_names;
                    handler_createWall.xy_shift = new List<double>();
                    handler_createWall.xy_shift.Add(double.Parse(shift_x.Text));
                    handler_createWall.xy_shift.Add(double.Parse(shift_y.Text));
                    externalEvent_CreateWall.Raise();
                    textBox7.Text = "0";
                }
                else if (comboBox10.Text == "排樁")
                {
                    //建立排樁
                    handler_CreateBentPile.files_path = new List<string>();
                    handler_CreateBentPile.files_path = file_names;
                    handler_CreateBentPile.xy_shift = new List<double>();
                    handler_CreateBentPile.xy_shift.Add(double.Parse(shift_x.Text));
                    handler_CreateBentPile.xy_shift.Add(double.Parse(shift_y.Text));
                    externalEvent_CreateBentPile.Raise();
                    textBox7.Text = "0";
                }
                else if(comboBox10.Text == "鋼板樁")
                {
                    //建立鋼板樁
                    handler_sheet_pile_NoCAD.files_path = new List<string>();
                    handler_sheet_pile_NoCAD.files_path = file_names;
                    handler_sheet_pile_NoCAD.xy_shift = new List<double>();
                    handler_sheet_pile_NoCAD.xy_shift.Add(double.Parse(shift_x.Text));
                    handler_sheet_pile_NoCAD.xy_shift.Add(double.Parse(shift_y.Text));
                    externalEvent_CreateSheetPile.Raise();
                    textBox7.Text = "0";
                }
                else if (comboBox10.Text == "型鋼樁" || comboBox10.Text == "鋼軌樁")
                {
                    //建立鋼板樁
                    handler_CreateSoldierPile.files_path = new List<string>();
                    handler_CreateSoldierPile.files_path = file_names;
                    handler_CreateSoldierPile.xy_shift = new List<double>();
                    handler_CreateSoldierPile.xy_shift.Add(double.Parse(shift_x.Text));
                    handler_CreateSoldierPile.xy_shift.Add(double.Parse(shift_y.Text));
                    handler_CreateSoldierPile.type = comboBox10.Text;
                    externalEvent_CreateSoldierPile.Raise();
                    textBox7.Text = "0";
                }
                else if (comboBox10.Text == "單元分割")
                {
                    
                    //建立連續壁CADmode
                    handler_CreateUnit.files_path = new List<string>();
                    handler_CreateUnit.files_path = file_names;
                    handler_CreateUnit.xy_shift = new List<double>();
                    handler_CreateUnit.xy_shift.Add(double.Parse(shift_x.Text));
                    handler_CreateUnit.xy_shift.Add(double.Parse(shift_y.Text));
                    externalEvent_CreateUnit.Raise();
                    
                }


            }
        }

        public IList<string> file_names
        {
            get;
            set;
        }

        public IList<string> pic_frame
        {
            get;
            set;
        }


        private void button3_Click(object sender, EventArgs e)
        {
            //建立中間樁
            handler_createColumn.ShiftData = new List<double>();
            for (int i = 0; i < 5; i++) { handler_createColumn.ShiftData.Add(0); }
            handler_createColumn.files_path = new List<string>();
            handler_createColumn.files_path = file_names;
            handler_createColumn.xy_shift = new List<double>();
            handler_createColumn.xy_shift.Add(double.Parse(shift_x.Text));
            handler_createColumn.xy_shift.Add(double.Parse(shift_y.Text));
            externalEvent_CreateColumn.Raise();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //中間樁間距之修正斷面選擇
            textBox4.Text = (comboBox1.SelectedItem as dynamic).Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //建置中間樁按鈕
            handler_createColumn.ShiftData = new List<double>();
            handler_createColumn.ShiftData.Add(double.Parse(textBox4.Text));
            handler_createColumn.ShiftData.Add(double.Parse(textBox5.Text));
            handler_createColumn.ShiftData.Add(double.Parse(textBox6.Text));
            handler_createColumn.files_path = new List<string>();
            handler_createColumn.files_path.Add(file_names[comboBox1.SelectedIndex]);
            externalEvent_CreateColumn.Raise();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //修正中間樁
            if (comboBox1.SelectedItem != null)
            {

                handler_moveColumn.SectionName = comboBox1.Text;
                handler_moveColumn.ShiftData = new List<double>();
                handler_moveColumn.ShiftData.Add(double.Parse(textBox2.Text));
                handler_moveColumn.ShiftData.Add(double.Parse(textBox3.Text));
                externalEvent_MoveColumn.Raise();
            }
            else
            {
                MessageBox.Show("請選擇欲修正斷面");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //選擇方向向後建置支撐
            if (comboBox2.SelectedItem != null)
            {
                handler_createFrame.files_path = new List<string>();
                handler_createFrame.files_path.Add(file_names[comboBox2.SelectedIndex]);
                handler_createFrame.draw_dir = new bool[2];

                if (comboBox3.Text == "x向")
                {
                    handler_createFrame.draw_dir[0] = true;
                    handler_createFrame.draw_dir[1] = false;
                }
                if (comboBox3.Text == "y向")
                {
                    handler_createFrame.draw_dir[0] = false;
                    handler_createFrame.draw_dir[1] = true;
                }
                if (comboBox3.Text == "雙向")
                {
                    handler_createFrame.draw_dir[0] = true;
                    handler_createFrame.draw_dir[1] = true;
                }

                handler_createFrame.draw_channel_steel = checkBox1.Checked;

                externalEvent_CreateFrame.Raise();

            }
            else
            {
                MessageBox.Show("請選擇欲修正斷面");
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (comboBox9.SelectedItem != null)
            {
                handler_createBeam.files_path = new List<string>();
                handler_createBeam.files_path.Add(file_names[comboBox9.SelectedIndex]);
                handler_createBeam.elevation_decide = new bool[2];
                if (comboBox8.Text == "搭接")
                {
                    handler_createBeam.elevation_decide[0] = true;
                    handler_createBeam.elevation_decide[1] = true;
                }
                if (comboBox8.Text == "x向重疊")
                {
                    handler_createBeam.elevation_decide[0] = false;
                    handler_createBeam.elevation_decide[1] = false;
                }
                if (comboBox8.Text == "y向重疊")
                {
                    handler_createBeam.elevation_decide[0] = false;
                    handler_createBeam.elevation_decide[1] = true;
                }

                externalEvent_CreateBeam.Raise();
            }
            else
            {
                MessageBox.Show("請選擇欲修正斷面");
            }
            step4.Show();
            step5.Show();
            step6.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //選擇方向及單雙向後開始繪製斜撐
            bool sel_combo_xy = new bool();
            bool sel_single_double = new bool();
            if (comboBox6.Text == "x向")
            {
                sel_combo_xy = true;
            }
            else if (comboBox6.Text == "y向")
            {
                sel_combo_xy = false;
            }
            if (comboBox5.Text == "單排")
            {
                sel_single_double = true;
            }
            else if (comboBox5.Text == "雙排")
            {
                sel_single_double = false;
            }
            handler_singleslopeframe.files_path = new List<string>();
            handler_singleslopeframe.files_path.Add(file_names[comboBox4.SelectedIndex]);
            handler_singleslopeframe.sel_combo_xy = sel_combo_xy;
            handler_singleslopeframe.sel_single_double = sel_single_double;
            handler_singleslopeframe.frame_or_slope = "斜撐";
            ExternalEvent_SingleSlopeFrame.Raise();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            //選擇方向及單雙向後開始繪製支撐
            bool sel_combo_xy = new bool();
            bool sel_single_double = new bool();
            if (comboBox6.Text == "x向")
            {
                sel_combo_xy = true;
            }
            else if (comboBox6.Text == "y向")
            {
                sel_combo_xy = false;
            }
            if (comboBox5.Text == "單排")
            {
                sel_single_double = true;
            }
            else if (comboBox5.Text == "雙排")
            {
                sel_single_double = false;
            }
            handler_singleslopeframe.files_path = new List<string>();
            handler_singleslopeframe.files_path.Add(file_names[comboBox4.SelectedIndex]);
            handler_singleslopeframe.sel_combo_xy = sel_combo_xy;
            handler_singleslopeframe.sel_single_double = sel_single_double;
            handler_singleslopeframe.frame_or_slope = "支撐";
            ExternalEvent_SingleSlopeFrame.Raise();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //建置回填按鈕
            handler_createSide.files_path = new List<string>();
            handler_createSide.files_path = file_names;
            externalEvent_CreateSide.Raise();
        }

        //最小化按鍵
        private void mini_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        //關閉按鍵
        private void leavebutton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void uploadbutton_Click(object sender, EventArgs e)
        {

            leftpanel.Height = uploadbutton.Height;
            leftpanel.Top = uploadbutton.Top;
            output1.Hide();
            output2.Hide();
            output3.Hide();
            main.Show();

        }

        private void outputbutton_Click(object sender, EventArgs e)
        {
            leftpanel.Height = outputbutton.Height;
            leftpanel.Top = outputbutton.Top;
            if (button1.BackColor == System.Drawing.Color.DeepSkyBlue)
            {
                output1.Show();
                output3.Hide();
                main.Hide();
            }

            else if (button10.BackColor == System.Drawing.Color.DeepSkyBlue)
            {
                output1.Hide();
                output2.Show();
                output3.Hide();
                main.Hide();
            }
            else
            {
                output1.Show();
                output3.Hide();
                main.Hide();
                //MessageBox.Show("請選擇建置形式");
            }


        }

        //上傳完檔案後所點擊之深開挖按鈕
        private void button1_Click_1(object sender, EventArgs e)
        {
            //深開挖基礎按鈕
            handler_output_excel.excavationType = "深開挖基礎";

            button1.BackColor = System.Drawing.Color.DeepSkyBlue;
            button10.BackColor = System.Drawing.Color.DodgerBlue;
            comboBox1.Items.Clear();
            comboBox1.DisplayMember = "Text";
            comboBox1.ValueMember = "Value";
            comboBox2.Items.Clear();
            comboBox2.DisplayMember = "Text";
            comboBox2.ValueMember = "Value";
            comboBox4.Items.Clear();
            comboBox4.DisplayMember = "Text";
            comboBox4.ValueMember = "Value";
            comboBox9.Items.Clear();
            comboBox9.DisplayMember = "Text";
            comboBox9.ValueMember = "Value";
            foreach (string file in file_names)
            {
                ExReader exReader = new ExReader();
                exReader.SetData(file, 1);
                exReader.PassColumnData();
                exReader.PassCircle();
                exReader.CloseEx();
                comboBox1.Items.Add(new { Text = exReader.section, Value = exReader.centralCol[0] });
                comboBox2.Items.Add(new { Text = exReader.section, Value = exReader.centralCol[0] });
                comboBox4.Items.Add(new { Text = exReader.section, Value = exReader.centralCol[0] });
                comboBox9.Items.Add(new { Text = exReader.section, Value = exReader.centralCol[0] });
            }

        }

        //上傳完檔案後所點擊之井式基礎按鈕
        private void button10_Click(object sender, EventArgs e)
        {
            handler_output_excel.excavationType = "井式基礎";
            button10.BackColor = System.Drawing.Color.DeepSkyBlue;
            button1.BackColor = System.Drawing.Color.DodgerBlue;

        }


        private void step1_Click(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.LemonChiffon;
            step2.BackColor = System.Drawing.Color.OldLace;
            step3.BackColor = System.Drawing.Color.OldLace;
            step4.BackColor = System.Drawing.Color.OldLace;
            step5.BackColor = System.Drawing.Color.OldLace;
            step6.BackColor = System.Drawing.Color.OldLace;
            button21.BackColor = System.Drawing.Color.OldLace;
            tabControl1.SelectTab(0);
        }


        private void button21_Click_1(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.OldLace;
            step2.BackColor = System.Drawing.Color.OldLace;
            step3.BackColor = System.Drawing.Color.OldLace;
            step4.BackColor = System.Drawing.Color.OldLace;
            step5.BackColor = System.Drawing.Color.OldLace;
            step6.BackColor = System.Drawing.Color.OldLace;
            button21.BackColor = System.Drawing.Color.LemonChiffon;
            tabControl1.SelectTab(6);
        }

        
        private void step2_Click(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.OldLace;
            step2.BackColor = System.Drawing.Color.LemonChiffon;
            step3.BackColor = System.Drawing.Color.OldLace;
            step4.BackColor = System.Drawing.Color.OldLace;
            step5.BackColor = System.Drawing.Color.OldLace;
            step6.BackColor = System.Drawing.Color.OldLace;
            button21.BackColor = System.Drawing.Color.OldLace;
            tabControl1.SelectTab(1);
        }


        private void step3_Click(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.OldLace;
            step2.BackColor = System.Drawing.Color.OldLace;
            step3.BackColor = System.Drawing.Color.LemonChiffon;
            step4.BackColor = System.Drawing.Color.OldLace;
            step5.BackColor = System.Drawing.Color.OldLace;
            step6.BackColor = System.Drawing.Color.OldLace;
            button21.BackColor = System.Drawing.Color.OldLace;
            tabControl1.SelectTab(2);
        }

        private void step4_Click(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.OldLace;
            step2.BackColor = System.Drawing.Color.OldLace;
            step3.BackColor = System.Drawing.Color.OldLace;
            step4.BackColor = System.Drawing.Color.LemonChiffon;
            step5.BackColor = System.Drawing.Color.OldLace;
            step6.BackColor = System.Drawing.Color.OldLace;
            button21.BackColor = System.Drawing.Color.OldLace;
            tabControl1.SelectTab(3);
        }

        private void step5_Click(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.OldLace;
            step2.BackColor = System.Drawing.Color.OldLace;
            step3.BackColor = System.Drawing.Color.OldLace;
            step4.BackColor = System.Drawing.Color.OldLace;
            step5.BackColor = System.Drawing.Color.LemonChiffon;
            step6.BackColor = System.Drawing.Color.OldLace;
            button21.BackColor = System.Drawing.Color.OldLace;
            tabControl1.SelectTab(4);
        }

        private void step6_Click(object sender, EventArgs e)
        {
            step1.BackColor = System.Drawing.Color.OldLace;
            step2.BackColor = System.Drawing.Color.OldLace;
            step3.BackColor = System.Drawing.Color.OldLace;
            step4.BackColor = System.Drawing.Color.OldLace;
            step5.BackColor = System.Drawing.Color.OldLace;
            step6.BackColor = System.Drawing.Color.LemonChiffon;
            button21.BackColor = System.Drawing.Color.OldLace;
            tabControl1.SelectTab(5);
        }

        //圖資產出按鈕
        private void pictureout_Click(object sender, EventArgs e)
        {
            leftpanel.Height = pictureout.Height;
            leftpanel.Top = pictureout.Top;
            output1.Hide();
            output2.Hide();
            main.Hide();
            output3.Show();
            picture_button.Show();
            output_excel_button.Show();
        }



        private void button6_Click_1(object sender, EventArgs e)
        {
            //建置斜撐，依照選取之方向和單雙向來進行
            if (comboBox2.SelectedItem != null)
            {
                handler_createSlope.files_path = new List<string>();
                handler_createSlope.files_path.Add(file_names[comboBox2.SelectedIndex]);
                handler_createSlope.draw_direc = new bool[2];

                if (comboBox3.Text == "x向")
                {
                    handler_createSlope.draw_direc[0] = true;
                    handler_createSlope.draw_direc[1] = false;
                }
                if (comboBox3.Text == "y向")
                {
                    handler_createSlope.draw_direc[0] = false;
                    handler_createSlope.draw_direc[1] = true;
                }
                if (comboBox3.Text == "雙向")
                {
                    handler_createSlope.draw_direc[0] = true;
                    handler_createSlope.draw_direc[1] = true;
                }
                externalEvent_CreateSlope.Raise();
            }
            else
            {
                MessageBox.Show("請選擇欲修正斷面");
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            //建置井式基礎
            handler_circle_excavation.files_path = new List<string>();
            handler_circle_excavation.files_path = file_names;
            handler_circle_excavation.xy_shift = new List<double>();
            handler_circle_excavation.xy_shift.Add(double.Parse(cir_shift_x.Text));
            handler_circle_excavation.xy_shift.Add(double.Parse(cir_shift_y.Text));
            externalEvent_circle_excavation.Raise();
        }


        private void button7_Click_3(object sender, EventArgs e)
        {
            //讀取圖框資料
            upload_textBoxFrame.Text = "";
            OpenFileDialog openFileDialogFrame = new OpenFileDialog();
            openFileDialogFrame.Multiselect = true;
            openFileDialogFrame.ShowDialog();
            handler_make_sheet.openFileDialog = openFileDialogFrame;
            pic_frame = openFileDialogFrame.FileNames;
            upload_textBoxFrame.Text = String.Format("已選取 {0} 個圖框", pic_frame.Count());

        }



        private void picture_button_Click(object sender, EventArgs e)
        {
            //開始出圖
            handler_make_sheet.sectionLineNumber = Convert.ToInt32(comboBox7.Text);
            handler_make_sheet.sectionName = Read_form.selected;//載入form2目前所選取的圖面
            externalEvent_make_sheet.Raise();
        }




        private void detect_button_Click(object sender, EventArgs e)
        {
            handler_deteclevel.FloorName = new List<string>();
            externalEvent_DetecLevel.Raise();
            button11.BackColor = System.Drawing.Color.OldLace;//剖面
            //button9.BackColor = System.Drawing.Color.OldLace;//模型倒出
            button12.BackColor = System.Drawing.Color.OldLace;//數量計算
            step7.BackColor = System.Drawing.Color.LemonChiffon;
            tabControl2.SelectTab(3);
        }
        private void Floor_dropdown(object sender, EventArgs e)
        {
            //監測儀器的樓層顯示
            select_floor_comboBox.Items.Clear();
            foreach (string floor in handler_deteclevel.FloorName)
            {
                select_floor_comboBox.Items.Add(floor);
            }
        }


        private void place_detector_Click(object sender, EventArgs e)
        {
            //放置監測儀器
            handler_putdetect.detectObject = select_detector_comboBox.Text;
            handler_putdetect.Levelname = select_floor_comboBox.Text;
            externalEvent_PutDetect.Raise();
        }

        private void CountBtm_Click(object sender, EventArgs e)
        {
            //數量計算
            handler_counting.files_path = new List<string>();
            handler_counting.files_path = file_names;

            ExternalEvent_counting.Raise();
        }



        private void output_excel_button_Click(object sender, EventArgs e)
        {
            //模型倒出
            ExternalEvent_output_excel.Raise();
        }


        private void output_excel_read_button_Click(object sender, EventArgs e)
        {
            //讀取圖框資料
            output_excel_textBox.Text = "";
            OpenFileDialog openFileDialogFrame = new OpenFileDialog();
            openFileDialogFrame.Multiselect = true;
            openFileDialogFrame.ShowDialog();
            handler_output_excel.FilePath = openFileDialogFrame.FileName;
            output_excel_textBox.Text = openFileDialogFrame.FileName;
        }


        private void button11_Click(object sender, EventArgs e)
        {
            button11.BackColor = System.Drawing.Color.LemonChiffon;//剖面
            //button9.BackColor = System.Drawing.Color.OldLace;//模型倒出
            button12.BackColor = System.Drawing.Color.OldLace;//數量計算
            step7.BackColor = System.Drawing.Color.OldLace;
            tabControl2.SelectTab(0);
        }


        /*private void button9_Click_1(object sender, EventArgs e)
        {
            button11.BackColor = System.Drawing.Color.OldLace;//剖面
            //button9.BackColor = System.Drawing.Color.LightSkyBlue;//模型倒出
            button12.BackColor = System.Drawing.Color.OldLace;//數量計算
            tabControl2.SelectTab(1);

        }*/

        private void button12_Click(object sender, EventArgs e)
        {
            button11.BackColor = System.Drawing.Color.OldLace;//剖面
            //button9.BackColor = System.Drawing.Color.OldLace;//模型倒出
            button12.BackColor = System.Drawing.Color.LemonChiffon;//數量計算
            step7.BackColor = System.Drawing.Color.OldLace;
            tabControl2.SelectTab(2);
            button12.Show();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //讀取CAD開挖範圍資料
            textBox1.Text = "";
            OpenFileDialog openFileDialogFrame = new OpenFileDialog();
            openFileDialogFrame.ShowDialog();
            handler_createWallCADmode.openFileDialog = openFileDialogFrame;
            handler_CreateSoldierPile.openFileDialog = openFileDialogFrame;
            handler_sheet_pile_NoCAD.openFileDialog = openFileDialogFrame;
            handler_CreateBentPile.openFileDialog = openFileDialogFrame;

            textBox1.Text = openFileDialogFrame.SafeFileName;
        }


        private void button14_Click(object sender, EventArgs e)
        {
            //顯示圖面選取視窗
            Read_form.Visible = true;
        }

        private void Save_loc_btn_Click(object sender, EventArgs e)
        {
            //指定數量計算之儲存路徑
            Save_loc_textbox.Text = "";
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = "";
            dialog.ShowDialog();
            handler_counting.path = dialog.FileName;
            Save_loc_textbox.Text = dialog.FileName;
        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            handler_VisualizeMonitor.files_path = new List<string>();
            handler_VisualizeMonitor.files_path = file_names;
            handler_VisualizeMonitor.monitor_path = textBox14.Text;
            handler_VisualizeMonitor.excel_path = textBox8.Text;

            externalEvent_VisualizeMonitor.Raise();
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            handler_CreateRebar.files_path = new List<string>();
            handler_CreateRebar.files_path = file_names;
            externalEvent_CreateRebar.Raise();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            //讀取斷面資料
            textBox8.Text = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            file_names = openFileDialog.FileNames;
            textBox8.Text = file_names[0];
            
        }


        private void button17_Click(object sender, EventArgs e)
        {
            //選擇方向及單雙向後開始繪製支撐
            bool sel_combo_xy = new bool();
            bool sel_single_double = new bool();
            if (comboBox6.Text == "x向")
            {
                sel_combo_xy = true;
            }
            else if (comboBox6.Text == "y向")
            {
                sel_combo_xy = false;
            }
            if (comboBox5.Text == "單排")
            {
                sel_single_double = true;
            }
            else if (comboBox5.Text == "雙排")
            {
                sel_single_double = false;
            }
            handler_drawjack.files_path = new List<string>();
            handler_drawjack.files_path.Add(file_names[comboBox4.SelectedIndex]);
            handler_drawjack.sel_combo_xy = sel_combo_xy;
            handler_drawjack.sel_single_double = sel_single_double;
            handler_drawjack.jack_array = new List<string>();
            handler_drawjack.jack_array.Add(textBox9.Text);
            handler_drawjack.jack_array.Add(textBox11.Text);
            ExternalEvent_DrawJack.Raise();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            //選擇方向及單雙向後開始繪製支撐
            bool sel_combo_xy = new bool();
            bool sel_single_double = new bool();
            if (comboBox6.Text == "x向")
            {
                sel_combo_xy = true;
            }
            else if (comboBox6.Text == "y向")
            {
                sel_combo_xy = false;
            }
            if (comboBox5.Text == "單排")
            {
                sel_single_double = true;
            }
            else if (comboBox5.Text == "雙排")
            {
                sel_single_double = false;
            }
            handler_draw_channelsteel.files_path = new List<string>();
            handler_draw_channelsteel.files_path.Add(file_names[comboBox4.SelectedIndex]);
            handler_draw_channelsteel.sel_combo_xy = sel_combo_xy;
            handler_draw_channelsteel.sel_single_double = sel_single_double;
            handler_draw_channelsteel.frame_or_slope = "Jack";
            ExternalEvent_DrawChannelSteel.Raise();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (comboBox9.SelectedItem != null)
            {
                handler_create_trisupport_beam.files_path = new List<string>();
                handler_create_trisupport_beam.files_path.Add(file_names[comboBox9.SelectedIndex]);
                handler_create_trisupport_beam.elevation_decide = new bool[2];
                if (comboBox8.Text == "搭接")
                {
                    handler_create_trisupport_beam.elevation_decide[0] = true;
                    handler_create_trisupport_beam.elevation_decide[1] = true;
                }
                if (comboBox8.Text == "x向重疊")
                {
                    handler_create_trisupport_beam.elevation_decide[0] = false;
                    handler_create_trisupport_beam.elevation_decide[1] = false;
                }
                if (comboBox8.Text == "y向重疊")
                {
                    handler_create_trisupport_beam.elevation_decide[0] = false;
                    handler_create_trisupport_beam.elevation_decide[1] = true;
                }

                ExternalEvent_CreateTriSupportofBeam.Raise();
            }
            else
            {
                MessageBox.Show("請選擇欲修正斷面");
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            //選擇方向向後建置支撐
            if (comboBox2.SelectedItem != null)
            {
                handler_create_trisupport_frame.files_path = new List<string>();
                handler_create_trisupport_frame.files_path.Add(file_names[comboBox2.SelectedIndex]);
                handler_create_trisupport_frame.draw_dir = new bool[2];

                if (comboBox3.Text == "x向")
                {
                    handler_create_trisupport_frame.draw_dir[0] = true;
                    handler_create_trisupport_frame.draw_dir[1] = false;
                }
                if (comboBox3.Text == "y向")
                {
                    handler_create_trisupport_frame.draw_dir[0] = false;
                    handler_create_trisupport_frame.draw_dir[1] = true;
                }
                if (comboBox3.Text == "雙向")
                {
                    handler_create_trisupport_frame.draw_dir[0] = true;
                    handler_create_trisupport_frame.draw_dir[1] = true;
                }

                handler_create_trisupport_frame.draw_channel_steel = checkBox1.Checked;

                ExternalEvent_CreateTriSupportOfFrame.Raise();

            }
            else
            {
                MessageBox.Show("請選擇欲修正斷面");
            }
        }

        private void topic_Click(object sender, EventArgs e)
        {

        }

        private void button21_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void supportpanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void upload_textbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void gradientPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        public IList<string> file_names_rebar
        {
            get;
            set;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            //讀取單元分割資料
            textBox10.Text = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
         
            openFileDialog.ShowDialog();
            file_names_rebar = openFileDialog.FileNames;
            textBox10.Text = String.Format("已選取");

            button1.Show();
            button10.Show();
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void button23_Click_1(object sender, EventArgs e)
        {
            handler_CreateRebar.files_path = new List<string>();
            handler_CreateRebar.files_path = file_names_rebar;
            externalEvent_CreateRebar.Raise();
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button26_Click(object sender, EventArgs e)
        {
            handler_create_U.files_path = new List<string>();
            handler_create_U.files_path = file_names;
            ExternalEvent_CreateU.Raise();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            textBox14.Text = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            file_names = openFileDialog.FileNames;
            textBox14.Text = file_names[0];
        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            textBox12.Text = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            IList<string> unit_dwg_file_name = openFileDialog.FileNames;
            textBox12.Text = unit_dwg_file_name[0];
        }

        private void button25_Click(object sender, EventArgs e)
        {
            textBox13.Text = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.ShowDialog();
            IList<string> unit_id_position_excel_path = openFileDialog.FileNames;
            textBox13.Text = unit_id_position_excel_path[0];
        }

        private void button24_Click(object sender, EventArgs e)
        {
            handler_CreateUnit.files_path = new List<string>();
            handler_CreateUnit.files_path = file_names;
            handler_CreateUnit.xy_shift = new List<double>();
            handler_CreateUnit.xy_shift.Add(double.Parse(shift_x.Text));
            handler_CreateUnit.xy_shift.Add(double.Parse(shift_y.Text));

            handler_CreateUnit.unit_dwg_file_name = textBox12.Text;
            handler_CreateUnit.unit_id_position_excel_path = textBox13.Text;

            externalEvent_CreateUnit.Raise();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            //指定數量計算之儲存路徑
            Save_loc_textbox.Text = "";
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.FileName = "";
            dialog.ShowDialog();
            handler_counting.path = dialog.FileName;
            Save_loc_textbox.Text = dialog.FileName;
        }
    }
}
