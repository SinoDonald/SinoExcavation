using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excavation
{
    public partial class form2 : Form
    {
        ExternalEvent externalEvent_ReadPlane;
        ReadPlane handler_readPlane = new ReadPlane();
        public form2()
        {
            InitializeComponent();
            externalEvent_ReadPlane = ExternalEvent.Create(handler_readPlane);
            checkedListBox1.Items.Clear();
        }
        public IList<string> selected
        {
            get;
            set;
        }


        private void SetReturnPlane(IList<string> PlanesName)
        {
            //此為委派的事件所需執行的方法
            checkedListBox1.Items.Clear();
            foreach(string s in PlanesName)
            {
                checkedListBox1.Items.Add(s);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //重新載入現有圖面
            externalEvent_ReadPlane.Raise();
            //委派SetReturnPlane這個方法到handler_readPlane事件
            handler_readPlane.ReturnPlanes += new ReadPlane.ReturnPlane(this.SetReturnPlane);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //隱藏圖面選取視窗，為了以防被關閉
            this.Visible = false;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //寫入目前所選取的圖面
            selected = new List<string>();
            for (int i = 0; i < ((CheckedListBox)sender).CheckedItems.Count; i++)
            {
                selected.Add(((CheckedListBox)sender).CheckedItems[i].ToString());
            }
        }
    }
}
