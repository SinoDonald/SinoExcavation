using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;
using System.Diagnostics;

namespace SinoExcavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class CreateTriSupportOfFrame : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        public IList<bool> draw_dir
        {
            get;
            set;
        }

        public bool draw_channel_steel
        {
            get;
            set;
        }
        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            foreach (string file_path in files_path)
            {
                //API CODE START

                try
                {
                    //讀取資料
                    ExReader dex = new ExReader();
                    dex.SetData(file_path, 1);
                    try
                    {
                        dex.PassFrameData();
                        dex.PassBeamData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message); }



                    ICollection<Level> levels = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
                    Level levdeep = levels.Where(x => x.Name.Contains("斷面" + dex.section) && !(x.Name.Contains("擋土壁深度"))).OrderBy(x => x.Elevation).ToList()[0];


                    //取得中間樁元件  (先篩選名字 因為會有其他instance沒有特定參數)
                    ICollection<FamilyInstance> columns_instance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                    where x.Name.Contains("中間樁")
                                                                    where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                                                    select x).ToList();


                    IList<XYZ> columns_xyz = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                              where x.Name.Contains("中間樁")
                                              where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                              select (x.Location as LocationPoint).Point).ToList();



                    double columns_dis = double.Parse(columns_instance.First().get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString()) * 1000 / 304.8;//透過標註參數讀取中間樁間距
                    double columns_H = double.Parse(columns_instance.First().Symbol.Name.Split('H')[1].Split('x')[0].ToString());//透過中間樁品類讀取中間樁H|||
                    double columns_B = double.Parse(columns_instance.First().Symbol.Name.Split('H')[1].Split('x')[1].ToString());//透過中間樁品類讀取中間樁B---

                    // 判斷中間樁方向 交換H&B

                    bool isColumnRotate = false;   // 0 

                    if ((columns_instance.First().Location as LocationPoint).Rotation != 0)
                    {
                        isColumnRotate = true;  // 90
                    }

                    
                    //三角托架元件
                    FamilySymbol trisupport = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                                     where x.Name == "三角架"
                                                     select x).First();

                    Transaction trans_2 = new Transaction(doc);
                    trans_2.Start("交易開始");
                    //string erroemessage = "";
                    for (int lev = 0; lev != dex.supLevel.Count(); lev++)
                    {
                        int temp;
                        if (int.TryParse(dex.supLevel[lev].Item1, out temp) == true)
                        {
                            
                            //讀取支撐HB
                            double Fframe_H = double.Parse(dex.supLevel[lev].Item4.Split('x')[0].Remove(0, 1));
                            double Fframe_B = double.Parse(dex.supLevel[lev].Item4.Split('x')[1]);
                                    

                            //三角托架

                            double additonal_depth = 0;    //  柱轉向90 : 0  ;  柱轉向0 :  (Fframe_H  or 槽鋼(200) )   

                            if (!isColumnRotate)
                            {
                                if (draw_channel_steel == true)
                                {
                                    additonal_depth = 200;
                                }
                                else
                                {
                                    additonal_depth = Fframe_H;
                                }
                            }

                            double triSupport_depth = (dex.supLevel[lev].Item2 * 1000 + additonal_depth) / -304.8; //  槽鋼深度暫用定值  case 1 : y向支+x向槽
                            //double triSupport_depth = (dex.supLevel[lev].Item2 * 1000 + Fframe_H) / -304.8;
                            //else triSupport_depth = (dex.supLevel[lev].Item2 * 1000 + beam_B) / -304.8;

                            foreach (XYZ column_pos in columns_xyz)
                            {


                                if (isColumnRotate)  // 中間柱 H  放左右  (放右 旋轉 鏡像)
                                {
                                    FamilyInstance trisupport_instance = doc.Create.NewFamilyInstance(new XYZ(column_pos.X + columns_H / 2 / 304.8, column_pos.Y , triSupport_depth)
                                                                                    , trisupport, StructuralType.NonStructural);

                                    XYZ pos = new XYZ(column_pos.X + columns_H / 2 / 304.8, column_pos.Y, 0);
                                    Line axis = Line.CreateBound(pos, pos + XYZ.BasisZ);

                                    ElementTransformUtils.RotateElement(doc, trisupport_instance.Id, axis, Math.PI * (0.5));


                                    if (dex.supLevel[lev].Item3 == 2)   // 隻數2 放對稱
                                    {

                                        

                                        ElementTransformUtils.MirrorElement(doc, trisupport_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, column_pos));
                                    }
                                }
                                else   // 中間柱 I  放上下
                                {
                                    FamilyInstance trisupport_instance = doc.Create.NewFamilyInstance(new XYZ(column_pos.X, column_pos.Y - columns_H / 2 / 304.8, triSupport_depth)
                                                                                    , trisupport, StructuralType.NonStructural);

                                    if (dex.supLevel[lev].Item3 == 2 || draw_channel_steel)   // 槽鋼或隻數2 放對稱
                                    {
                                        ElementTransformUtils.MirrorElement(doc, trisupport_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, column_pos));
                                    }
                                }


                            }

                        }
                    }
                    trans_2.Commit();
                }


                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message); break; }
            }
            TaskDialog.Show("done", "三角托架建置完畢");
        }
        
       
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
