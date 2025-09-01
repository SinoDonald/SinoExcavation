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

    class CreateU : IExternalEventHandler
    {
        public IList<string> files_path
        {
            get;
            set;
        }
        
        /*public IList<bool> draw_dir
        {
            get;
            set;
        }*/

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
                    

                    //取得中間樁元件
                    ICollection<FamilyInstance> columns_instance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                    where x.Name.Contains("中間樁")
                                                                    where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                                                    select x).ToList();
                    
                    //取得中間樁位置xyz
                    IList<XYZ> columns_xyz = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                              where x.Name.Contains("中間樁")
                                              where x.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString().Split(':')[0] == "斷面" + dex.section
                                              select (x.Location as LocationPoint).Point).ToList();
                    
                    double columns_dis = double.Parse(columns_instance.First().get_Parameter(BuiltInParameter.DOOR_NUMBER).AsString()) * 1000 / 304.8;//透過標註參數讀取中間樁間距
                    double columns_H = double.Parse(columns_instance.First().Symbol.Name.Split('H')[1].Split('x')[0].ToString());//透過中間樁品類讀取中間樁H|||
                    double columns_B = double.Parse(columns_instance.First().Symbol.Name.Split('H')[1].Split('x')[1].ToString());//透過中間樁品類讀取中間樁B---
                    //TaskDialog.Show("TEST", "中間樁高 = "+columns_H.ToString()+"\n中間樁寬 = "+columns_B.ToString());
                    
                    // 判斷中間樁方向 交換H&B
                    /*
                    bool isColumnRotate = false;   // 0 

                    if ((columns_instance.First().Location as LocationPoint).Rotation != 0)
                    {
                        isColumnRotate = true;  // 90
                    }
                    */

                    // 讀取U形螺栓元件
                    FamilySymbol ubolt48SUB = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "48SUB").First();
                    FamilySymbol ubolt48LUB = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "48LUB").First();
                    FamilySymbol ubolt = ubolt48LUB; //暫定

                    // 讀取固定角鐵元件
                    FamilySymbol fixedIron = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == "固定角鐵").First();

                    Transaction trans1 = new Transaction(doc);
                    trans1.Start("交易開始");
                    for (int lev = 0; lev != dex.supLevel.Count(); lev++) //支撐階數
                    {
                        int temp;
                        if (int.TryParse(dex.supLevel[lev].Item1, out temp) == true) //判斷是不是int
                        {
                            //讀取支撐H(高)B(寬)
                            double Fframe_H = double.Parse(dex.supLevel[lev].Item4.Split('x')[0].Remove(0, 1));
                            double Fframe_B = double.Parse(dex.supLevel[lev].Item4.Split('x')[1]);
                            //TaskDialog.Show("TEST", "支撐高 = " + Fframe_H.ToString() + "\n支撐寬 = " + Fframe_B.ToString());

                            
                            double ubolt_depth = (dex.supLevel[lev].Item2 * 1000 + Fframe_H) / -304.8 ;
                            double fixedIron_depth = (dex.supLevel[lev].Item2 * 1000 - Fframe_H) / -304.8;

                            foreach (XYZ column_pos in columns_xyz)
                            {
                                ubolt.Activate();
                                fixedIron.Activate();

                                double ubolt_B = 420/304.8;

                                FamilyInstance ubolt_instance1 = doc.Create.NewFamilyInstance(new XYZ(column_pos.X - columns_B/2/304.8, 
                                                                                                      column_pos.Y - columns_H/2/304.8 - columns_B/2/304.8 - Fframe_B/304.8,
                                                                                                      ubolt_depth),
                                                                                                      ubolt, StructuralType.NonStructural);
                                FamilyInstance ubolt_instance2 = doc.Create.NewFamilyInstance(new XYZ(column_pos.X - columns_B/2/304.8,
                                                                                                      column_pos.Y - columns_H/2/304.8 - columns_B/2/304.8,    
                                                                                                      ubolt_depth),
                                                                                                      ubolt, StructuralType.NonStructural);

                                FamilyInstance fixedIron_instance1 = doc.Create.NewFamilyInstance(new XYZ(column_pos.X + columns_B/2/304.8 + Fframe_B/2/304.8, //X 是對的
                                                                                                          column_pos.Y - columns_H/2/304.8, //Y
                                                                                                          fixedIron_depth),
                                                                                                          fixedIron, StructuralType.NonStructural);
                                FamilyInstance fixedIron_instance2 = doc.Create.NewFamilyInstance(new XYZ(column_pos.X + columns_B/2/304.8 + Fframe_B/2/304.8, //X 是對的
                                                                                                          column_pos.Y - columns_H/2/304.8 - ubolt_B,  //Y
                                                                                                          fixedIron_depth),
                                                                                                          fixedIron, StructuralType.NonStructural);

                                XYZ pos = new XYZ(column_pos.X , column_pos.Y - columns_H /2/304.8, 0);
                                Line axis = Line.CreateBound(pos, pos + XYZ.BasisZ);
                                ElementTransformUtils.RotateElement(document, ubolt_instance1.Id, axis,  Math.PI/2);
                                ElementTransformUtils.RotateElement(document, ubolt_instance2.Id, axis, Math.PI / 2);
                                
                                if (dex.supLevel[lev].Item3 == 2)   // 隻數2 放對稱
                                {
                                    // U bolt
                                    ElementTransformUtils.MirrorElement(doc, ubolt_instance1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, column_pos));
                                    ElementTransformUtils.MirrorElement(doc, ubolt_instance1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, column_pos));
                                    XYZ loc = new XYZ(-(columns_B+2*Fframe_B)/304.8, (columns_B + Fframe_B) / 304.8, 0);
                                    ElementTransformUtils.CopyElement(doc, ubolt_instance1.Id, loc);

                                    ElementTransformUtils.MirrorElement(doc, ubolt_instance2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, column_pos));
                                    ElementTransformUtils.MirrorElement(doc, ubolt_instance2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, column_pos));
                                    XYZ loc2 = new XYZ(-(columns_B)/304.8, (columns_B + Fframe_B) / 304.8, 0);
                                    ElementTransformUtils.CopyElement(doc, ubolt_instance2.Id, loc2);

                                    // fixed iron
                                    ElementTransformUtils.MirrorElement(doc, fixedIron_instance1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, column_pos));
                                    ElementTransformUtils.MirrorElement(doc, fixedIron_instance1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, column_pos));
                                    XYZ loc3 = new XYZ(-(columns_B + Fframe_B)/304.8, (columns_H + Fframe_H) / 304.8, 0);
                                    ElementTransformUtils.CopyElement(doc, fixedIron_instance1.Id, loc3);

                                    ElementTransformUtils.MirrorElement(doc, fixedIron_instance2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, column_pos));
                                    ElementTransformUtils.MirrorElement(doc, fixedIron_instance2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, column_pos));
                                    XYZ loc4 = new XYZ(-(columns_B+Fframe_B)/304.8, (columns_H + Fframe_H) / 304.8, 0);
                                    ElementTransformUtils.CopyElement(doc, fixedIron_instance2.Id, loc4);
                                }                           
                            }                          
                        }
                    }
                    trans1.Commit();
                }
                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.ToString()); break; }
            }
            TaskDialog.Show("done", "U型螺栓建置完畢");
        }


        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}