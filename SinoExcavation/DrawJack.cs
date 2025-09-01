using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;

namespace SinoExcavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]


    class DrawJack : IExternalEventHandler
    {
        //用哪一個斷面處理
        public IList<string> files_path
        {
            get;
            set;
        }
        //建置在哪一個高程，利用這個布林子判斷
        public bool sel_combo_xy
        {
            get;
            set;
        }
        //建置在單向階層或雙向階層的布林子
        public bool sel_single_double
        {
            get;
            set;
        }
        //千斤頂陣列(右,下,間距)
        public IList<string> jack_array
        {
            get;
            set;
        }


        public void Execute(UIApplication app)
        {
            Autodesk.Revit.DB.Document document = app.ActiveUIDocument.Document;
            UIDocument uidoc = new UIDocument(document);

            Document doc = uidoc.Document;
            //每一個斷面資料分開處理
            foreach (string file_path in files_path)
            {

                Transaction transaction = new Transaction(doc);

                ViewFamilyType viewFamilyType = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().Where(x => x.Name == "樓板平面圖").First();

                View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().First();

                Level baseLevel = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().First();
               

                ExReader cop = new ExReader();
                cop.SetData(file_path, 1);
                cop.PassFrameData();
                cop.CloseEx();
                
                //藉由判斷階層為單雙向的支撐，來決定此階層需不需要我們繪製的斜撐，並將高程記錄下來
                //以便等等運算使用
                IList<double> ob_3 = new List<double>();  //深度
                IList<string> ob_4 = new List<string>();  //型號
                for (int i = 0; i != cop.supLevel.Count(); i++)
                {
                    //判斷單向或雙向
                    double ss;
                    if (sel_single_double == true)
                    { ss = 1; }
                    else { ss = 2; }

                    if (cop.supLevel[i].Item3 == ss)
                    {
                        ob_3.Add(cop.supLevel[i].Item2);
                        ob_4.Add(cop.supLevel[i].Item4);
                    }
                }
                //建置平面視圖
                transaction.Start("create view");
                //取得型鋼第一個的厚度以作偏移
                double a = double.Parse(ob_4.First().Split('x').First().Remove(0, 1));

                double LL_eleva = new double();
                if (sel_combo_xy == true)  //X向
                {
                    LL_eleva = ob_3.First() * (-1) * 1000 / 304.8;
                }
                else //Y向，由於Y向斜撐較低，所以須作偏移
                {
                    LL_eleva = (ob_3.First() * (-1) + a * 0.001) * 1000 / 304.8;
                }
                //創建此高程及高程視圖以便使用者繪製
                Level LL = Level.Create(doc, LL_eleva);
                ViewPlan viewPlan = ViewPlan.Create(doc, viewFamilyType.Id, LL.Id);

                //繼承視圖範圍及修改
                PlanViewRange planViewRange = viewPlan.GetViewRange();

                planViewRange.SetOffset(PlanViewPlane.ViewDepthPlane, -2 * 1000 / 304.8);

                planViewRange.SetOffset(PlanViewPlane.BottomClipPlane, -1 * 1000 / 304.8);

                viewPlan.SetViewRange(planViewRange);

                transaction.Commit();

                //更改當前視圖
                uidoc.ActiveView = viewPlan;
                //蒐集目前元件
                IList<ElementId> old_familyInstance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                       where x.Name == "千斤頂"
                                                       select x.Id).ToList();

                transaction.Start("create view");
                //取得千斤頂元件
                FamilySymbol jackSymbol = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                                           where x.Name == "千斤頂"
                                           select x).ToList().First();
               
                TaskDialog.Show("繪製千斤頂", "Start" );

                if (!jackSymbol.IsActive)
                {
                    jackSymbol.Activate();
                    doc.Regenerate();
                }
                transaction.Commit();
                
                //進入繪製此元件的步驟
                try
                {
                    uidoc.PromptForFamilyInstancePlacement(jackSymbol);
                }
                catch { };

                //結束繪製之後跳出提示視窗
                TaskDialog.Show("繪製千斤頂", "畫完了");

                //蒐集新元件
                IList<ElementId> familyInstance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                       where x.Name == "千斤頂"
                                                       select x.Id).ToList();

                try
                {
                    transaction.Start("Copy & Cut");
                    foreach (ElementId famI in familyInstance)
                    {
                        //若元件沒包含在舊元件內
                        //則此元件即為剛剛被建置出來之元件
                        //再做以下處理
                        if (old_familyInstance.Contains(famI) == false)
                        {
                            
                            doc.GetElement(famI).get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(baseLevel.Id);

                            CopyPaste(famI, sel_combo_xy, sel_single_double, cop, doc, int.Parse(jack_array[0]), double.Parse(jack_array[1]));
                            
                        }
                    }
                }
                catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message + e.StackTrace); }
                

                
                //蒐集新元件(複製後)
                familyInstance = (from x in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                  where x.Name == "千斤頂"
                                  select x.Id).ToList();
                
                //空心切割
                try
                {
                    
                    foreach (ElementId famI in familyInstance)
                    {
                        //若元件沒包含在舊元件內
                        //則此元件即為剛剛被建置出來之元件
                        //再做以下處理
                        if (old_familyInstance.Contains(famI) == false)
                        {
                            
                            doc.GetElement(famI).get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM).Set(baseLevel.Id);

                            CutJack(famI, doc);
                                                        
                        }
                    }
                }
                catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message + e.StackTrace); }
                transaction.Commit();



                //結束更改並刪除試圖
                uidoc.ActiveView = view3D;
                transaction.Start("delete view");
                doc.Delete(LL.Id);
                transaction.Commit();

            }

        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
        public void CopyPaste(ElementId id, bool XY, bool onetwo, ExReader cop, Document doc, int num, double spaceing)
        {
            //判斷單雙向
            int i;
            if (onetwo == true)
            { i = 1; }
            else { i = 2; }
            //將階數、型號及深度放入Tuple中
            IList<Tuple<string, double, string>> one_slope = new List<Tuple<string, double, string>>();
            for (int love = 0; love != cop.supLevel.Count(); love++)
            {
                if (cop.supLevel[love].Item3 == i)
                {
                    Tuple<string, double, string> addin = new Tuple<string, double, string>(cop.supLevel[love].Item1, cop.supLevel[love].Item2, cop.supLevel[love].Item4);
                    one_slope.Add(addin);
                }
            }
            double first_elev = one_slope.First().Item2 * (-1000);
            doc.GetElement(id).get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).SetValueString(first_elev.ToString());

            for (int one = 0; one != one_slope.Count(); one++)
            {
                //利用所需到達的深度-第一個階層的深度來計算偏移量
                //double ori_high = one_slope.First().Item2;
                double sup_deep = one_slope[one].Item2;
                double plus;
                                
                string thick = one_slope[one].Item3.Split('x').First().Remove(0, 1);
                string support_B = one_slope[one].Item3.Split('x')[1];  //透過中間樁品類讀取中間樁B---

                plus = double.Parse(thick) / 2;
                string direction = "Y";

                if (XY == true) //Y向的話，要再扣掉一個型鋼的厚度
                {
                    plus *= -1;
                    direction = "X";
                }

                double elev = sup_deep * (-1000) + plus;

                //算出偏移量
                //double ne_eleva = (old_math * (-1) + ori_high) * 1000 + plus;
                //取得要複製之元件
                Element element = doc.GetElement(id);
                
                XYZ loc = new XYZ(0, 0, 0);

                
                //複製元件 (階層)
                if (one > 0)
                {
                    ICollection<ElementId> ass = ElementTransformUtils.CopyElement(doc, id, loc);

                    element = doc.GetElement(ass.First());
                }

                //參數值 H, B, Z偏移, XY方向
                element.LookupParameter("H").SetValueString(thick);
                element.LookupParameter("B").SetValueString(support_B);
                                
                element.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).SetValueString((elev).ToString());
                element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(direction);

                //複製元件 (陣列) X 向上  Y 向右
                if (direction == "X")
                {
                    for (int j = 1; j < num; j++)
                    {
                        ElementTransformUtils.CopyElement(doc, element.Id, new XYZ(0, spaceing, 0) * j * 1000 / 304.8);
                    }
                }
                else
                {
                    for (int j = 1; j < num; j++)
                    {
                        ElementTransformUtils.CopyElement(doc, element.Id, new XYZ(spaceing, 0, 0) * j * 1000 / 304.8);
                    }
                }
                                
                //寫備註
                //new_ele.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + "-" + one_slope[one].Item1.ToString() + "-" + frame_or_slope);

                
            }

        }

        public void CutJack(ElementId id, Document doc)
        {

            
            Element jack_instance = doc.GetElement(id);

            XYZ jack_xyz = (jack_instance.Location as LocationPoint).Point;

            
            string direction = jack_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString();



            // 法1 FindNearest

            //View3D view3D = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().First();


            //ReferenceIntersector refIntersector = new ReferenceIntersector(new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming), FindReferenceTarget.Element, view3D);

            //FamilyInstance jack_instance = doc.GetElement(id) as FamilyInstance;

            //jack_xyz = Double.Parse(doc.GetElement(id).get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).AsValueString()) / 304.8;

            //double jack_B = double.Parse((jack_instance as FamilyInstance).Symbol.LookupParameter("B").AsValueString());

            //try
            //{

            //    if (direction == "X")
            //    {
            //        var refID = refIntersector.FindNearest(jack_xyz - new XYZ(0, jack_B / 2 / 304.8, 0), XYZ.BasisY).GetReference().ElementId;
            //        var ele = doc.GetElement(refID);
            //        InstanceVoidCutUtils.AddInstanceVoidCut(doc, ele, jack_instance);

            //        Line axis = Line.CreateBound(jack_xyz, jack_xyz + XYZ.BasisZ);
            //        XYZ slope = ((ele.Location as LocationCurve).Curve as Line).Direction;
            //        double angle;
            //        if (slope.Y >= 0) angle = new XYZ(1, 0, 0).AngleTo(slope);
            //        else angle = new XYZ(-1, 0, 0).AngleTo(slope) + Math.PI;
            //        ElementTransformUtils.RotateElement(doc, id, axis, angle);
            //    }
            //    else
            //    {
            //        var refID = refIntersector.FindNearest(jack_xyz - new XYZ(jack_B / 2 / 304.8, 0 , 0), XYZ.BasisX).GetReference().ElementId;
            //        var ele = doc.GetElement(refID);
            //        InstanceVoidCutUtils.AddInstanceVoidCut(doc, ele, jack_instance);
            //    }


            //}
            //catch (Exception e)
            //{
            //    TaskDialog.Show("Error", e.Message);

            //}

            //法2 BoundingBoxIntersectsFilter

            try
            {
                //用 BoundingBox 尋找相交元件
                BoundingBoxIntersectsFilter boxIntersectsFilter = new BoundingBoxIntersectsFilter(new Outline(jack_instance.get_BoundingBox(doc.ActiveView).Min, jack_instance.get_BoundingBox(doc.ActiveView).Max));
                BoundingBoxIsInsideFilter boxIsInsideFilter = new BoundingBoxIsInsideFilter(new Outline(jack_instance.get_BoundingBox(doc.ActiveView).Min, jack_instance.get_BoundingBox(doc.ActiveView).Max));

                //收集過濾 (category, familyinstance, boundingbox)
                IList<FamilyInstance> intersectInstance = (from x in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming)
                                                           .OfClass(typeof(FamilyInstance))
                                                           .WherePasses(new LogicalOrFilter(boxIntersectsFilter, boxIsInsideFilter)).Cast<FamilyInstance>()
                                                           select x).ToList();

                //由於支撐用Z偏移, Boundingbox的z會被拉大, 最後會濾出同XY不同Z之型鋼, 須再多加判斷式

                foreach (var familyInstance in intersectInstance)
                {
                   
                    if (familyInstance.GetTransform().Origin.Z <= jack_instance.get_BoundingBox(doc.ActiveView).Max.Z && familyInstance.GetTransform().Origin.Z >= jack_instance.get_BoundingBox(doc.ActiveView).Min.Z)
                    {
                        //空心切割實體                      
                        InstanceVoidCutUtils.AddInstanceVoidCut(doc, familyInstance, jack_instance);

                        //旋轉千斤頂至與被切割體同方向
                        Line axis = Line.CreateBound(jack_xyz, jack_xyz + XYZ.BasisZ);
                        XYZ slope = ((familyInstance.Location as LocationCurve).Curve as Line).Direction;
                        double angle;
                        if (slope.Y >= 0)
                        {
                            angle = new XYZ(1, 0, 0).AngleTo(slope);
                        }
                        else
                        {
                            angle = new XYZ(-1, 0, 0).AngleTo(slope) + Math.PI; 
                        }
                        
                        ElementTransformUtils.RotateElement(doc, id, axis, angle);

                        //對齊中心距(只限xy向)  因為不同尺寸之型鋼中心線並不在同一線上
                        if (slope.Y == 0)
                        {                            
                            ElementTransformUtils.MoveElement(doc, id, new XYZ(0, familyInstance.GetTransform().Origin.Y - jack_xyz.Y, 0));
                        }
                        else if (slope.X == 0)
                        {
                            ElementTransformUtils.MoveElement(doc, id, new XYZ(familyInstance.GetTransform().Origin.X - jack_xyz.X, 0, 0));
                        }

                        break;
                    }

                }

                
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.Message);
            }

           

        }

    }
}