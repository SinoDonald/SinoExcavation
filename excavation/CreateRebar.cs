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
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

                    //TaskDialog.Show("1", dex.protection_width.ToString());
                    //TaskDialog.Show("1", dex.shear_rebar[0].Item4.ToString());

                    //開始建立鋼筋
                    //先檢查讀入資料有沒有錯
                    //垂直筋擋土側

                    
                    foreach (var data in dex.vertical_r_rebar)
                    {
                        TaskDialog.Show("垂直筋擋土側", data.Item1.ToString() + '/' + data.Item2.ToString() + '/' + data.Item3.ToString() + '/' +
                            data.Item4.ToString() + '/' + data.Item5.ToString() + '/' + data.Item6.ToString());
                    }
                    //垂直筋開挖側
                    foreach (var data in dex.vertical_e_rebar)
                    {
                        TaskDialog.Show("垂直筋開挖側", data.Item1.ToString() + '/' + data.Item2.ToString() + '/' + data.Item3.ToString() + '/' +
                            data.Item4.ToString() + '/' + data.Item5.ToString() + '/' + data.Item6.ToString());
                    }
                    //水平筋
                    foreach (var data in dex.horizontal_rebar)
                    {
                        TaskDialog.Show("水平筋", data.Item1.ToString() + '/' + data.Item2.ToString() + '/' + data.Item3.ToString() +
                            data.Item4.ToString());
                    }
                    
                    //剪力筋
                    
                    foreach (var data in dex.shear_rebar_depth)
                    {
                        TaskDialog.Show("剪力筋", data.Item1.ToString() + '/' + data.Item2.ToString());
                    }
                    foreach (var data in dex.shear_rebar)
                    {
                        TaskDialog.Show("剪力筋", data.Item1.ToString() + '/' + data.Item2.ToString() + '/' + data.Item3.ToString() + '/' +
                            data.Item4.ToString() + '/' + data.Item5.ToString() + '/' + data.Item6.ToString());
                    }

                }
                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); break; }
            }

            
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}