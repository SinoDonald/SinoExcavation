using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;

namespace SinoExcavation_2020
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    class SingleSlopeFrame : IExternalEventHandler
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
        //是建置斜撐或支撐
        public string frame_or_slope
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

                ExReader cop = new ExReader();
                cop.SetData(file_path, 1);
                cop.PassFrameData();
                cop.PassBeamData();
                cop.CloseEx();

                //藉由判斷階層為單雙向的支撐，來決定此階層需不需要我們繪製的斜撐，並將高程記錄下來
                //以便等等運算使用
                IList<double> ob_3 = new List<double>(); //深度
                IList<string> ob_4 = new List<string>(); //型號

                //圍囹
                IList<string> beam_name = new List<string>(); //型號

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
                        beam_name.Add(cop.beamLevel[i].Item3);
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
                LL.Name = "繪製支撐平面";

                //繼承視圖範圍及修改
                PlanViewRange planViewRange = viewPlan.GetViewRange();

                planViewRange.SetOffset(PlanViewPlane.ViewDepthPlane, -2 * 1000 / 304.8);

                planViewRange.SetOffset(PlanViewPlane.BottomClipPlane, -1 * 1000 / 304.8);

                viewPlan.SetViewRange(planViewRange);

                transaction.Commit();

                //更改當前視圖
                uidoc.ActiveView = viewPlan;

                //蒐集目前元件
                //斜撐
                IList<ElementId> old_familyInstance = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name == ob_4.First()).ToList().Select(x => x.Id).ToList();
                //圍囹
                IList<ElementId> old_familyInstance_beam = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name == beam_name.First()).ToList().Select(x => x.Id).ToList();

                //建立大斜撐
                FamilySymbol familySymbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == ob_4.First()).First();
                //建立圍囹
                FamilySymbol familySymbol_beam = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == beam_name.First()).First();


                //進入繪製此元件的步驟
                try
                {
                    uidoc.PromptForFamilyInstancePlacement(familySymbol);
                }
                catch { };

                //蒐集新元件-斜撐
                IList<ElementId> familyInstance = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name == ob_4.First()).ToList().Select(x => x.Id).ToList();

                //補圍囹
                TaskDialog.Show("Test", "開始補圍囹，若不需補則按esc離開");
                try
                {
                    uidoc.PromptForFamilyInstancePlacement(familySymbol_beam);
                }
                catch { };

                //蒐集新元件-圍囹
                IList<ElementId> familyInstance_beam = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>().Where(x => x.Name == beam_name.First()).ToList().Select(x => x.Id).ToList();

                //檢查斜撐和圍囹型號是否一樣，避免重複蒐集
                foreach (ElementId famI in familyInstance)
                {
                    if (familyInstance_beam.Contains(famI))
                    {
                        familyInstance_beam.Remove(famI);
                    }
                }

                //結束繪製之後跳出提示視窗
                TaskDialog.Show("Test", "畫完了");

                try
                {
                    transaction.Start("寫備註&複製貼上");
                    foreach (ElementId famI in familyInstance)
                    {
                        //若元件沒包含在舊元件內
                        //則此元件即為剛剛被建置出來之元件
                        //再做以下處理
                        if (old_familyInstance.Contains(famI) == false)
                        {
                            //取消接合
                            StructuralFramingUtils.DisallowJoinAtEnd(doc.GetElement(famI) as FamilyInstance, 0);
                            StructuralFramingUtils.DisallowJoinAtEnd(doc.GetElement(famI) as FamilyInstance, 1);

                            if (sel_single_double == true) //單向
                            {
                                doc.GetElement(famI).get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + String.Format("-{0}-", cop.First_single) + frame_or_slope);
                            }
                            else  //雙向
                            {
                                doc.GetElement(famI).get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + String.Format("-{0}-", cop.First_double) + frame_or_slope);
                            }
                            //複製貼上至所需階層，此方法的說明在下方
                            CopyPaste(famI, sel_combo_xy, sel_single_double, cop, doc);
                        }
                    }
                }

                catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message + e.StackTrace); }
                transaction.Commit();
                //TaskDialog.Show("Test", "複製完斜撐");

                //圍囹
                try
                {
                    transaction.Start("圍囹寫備註&複製貼上");
                    foreach (ElementId famI in familyInstance_beam)
                    {
                        //若元件沒包含在舊元件內
                        //則此元件即為剛剛被建置出來之元件
                        //再做以下處理
                        if (old_familyInstance_beam.Contains(famI) == false)
                        {
                            //取消接合
                            StructuralFramingUtils.DisallowJoinAtEnd(doc.GetElement(famI) as FamilyInstance, 0);
                            StructuralFramingUtils.DisallowJoinAtEnd(doc.GetElement(famI) as FamilyInstance, 1);


                            if (sel_single_double == true) //單向
                            {
                                doc.GetElement(famI).get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + String.Format("-{0}-", cop.First_single) + "圍囹");
                            }
                            else  //雙向
                            {
                                doc.GetElement(famI).get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + String.Format("-{0}-", cop.First_double) + "圍囹");
                            }
                            int count = 0;
                            //複製貼上至所需階層，此方法的說明在下方
                            CopyPaste_beam(famI, sel_combo_xy, sel_single_double, cop, doc, count);
                        }
                    }
                }

                catch (Exception e) { TaskDialog.Show("error", Environment.NewLine + e.Message + e.StackTrace); }
                transaction.Commit();

                //TaskDialog.Show("Test", "複製完圍囹");

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

        IList<ElementId> slopes = new List<ElementId>();

        public void CopyPaste(ElementId id, bool XY, bool onetwo, ExReader cop, Document doc)
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

            for (int one = 1; one != one_slope.Count(); one++)
            {
                //利用所需到達的深度-第一個階層的深度來計算偏移量
                double ori_high = one_slope.First().Item2;
                double old_math = one_slope[one].Item2;
                double plus = 0;

                if (XY == false) //Y向的話，要再扣掉一個型鋼的厚度
                {
                    string thick = one_slope.First().Item3.Split('x').First().Remove(0, 1);
                    string ne_thick = one_slope[one].Item3.Split('x').First().Remove(0, 1);
                    plus = double.Parse(ne_thick) - double.Parse(thick);
                }
                //算出偏移量
                double ne_eleva = (old_math * (-1) + ori_high) * 1000 + plus;
                //取得要複製之元件
                Element element = doc.GetElement(id);

                XYZ loc = new XYZ(0, 0, 0);
                //複製元件
                ICollection<ElementId> ass = ElementTransformUtils.CopyElement(doc, id, loc);

                Element new_ele = doc.GetElement(ass.First());
                FamilyInstance new_fami = new_ele as FamilyInstance;
                //指定symbol類型
                new_fami.Symbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == one_slope[one].Item3).First();
                //偏移元件
                new_ele.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString(ne_eleva.ToString());
                //寫備註
                new_ele.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + "-" + one_slope[one].Item1.ToString() + "-" + frame_or_slope);
                //取消接合

                StructuralFramingUtils.DisallowJoinAtEnd(element as FamilyInstance, 0);
                StructuralFramingUtils.DisallowJoinAtEnd(element as FamilyInstance, 1);
                StructuralFramingUtils.DisallowJoinAtEnd(new_ele as FamilyInstance, 0);
                StructuralFramingUtils.DisallowJoinAtEnd(new_ele as FamilyInstance, 1);


                slopes.Add(ass.First());
            }



        }

        public void CopyPaste_beam(ElementId id, bool XY, bool onetwo, ExReader cop, Document doc, int count)
        {
            //判斷單雙向
            int i;
            if (onetwo == true)
            { i = 1; }
            else { i = 2; }
            //將階數、型號及深度放入Tuple中
            IList<Tuple<string, double, string>> one_beam = new List<Tuple<string, double, string>>();

            for (int love = 0; love != cop.supLevel.Count(); love++)
            {
                if (cop.supLevel[love].Item3 == i)
                {
                    Tuple<string, double, string> addin = new Tuple<string, double, string>(cop.supLevel[love].Item1, cop.supLevel[love].Item2, cop.beamLevel[love].Item3);
                    one_beam.Add(addin);
                }
            }

            for (int one = 1; one != one_beam.Count(); one++)
            {
                //利用所需到達的深度-第一個階層的深度來計算偏移量
                double ori_high = one_beam.First().Item2;
                double old_math = one_beam[one].Item2;
                double plus = 0;

                if (XY == false) //Y向的話，要再扣掉一個型鋼的厚度
                {
                    string thick = one_beam.First().Item3.Split('x').First().Remove(0, 1);
                    string ne_thick = one_beam[one].Item3.Split('x').First().Remove(0, 1);
                    plus = double.Parse(ne_thick) - double.Parse(thick);
                }
                //算出偏移量
                double ne_eleva = (old_math * (-1) + ori_high) * 1000 + plus;
                //取得要複製之元件
                Element element = doc.GetElement(id);
                LocationPoint fam_location = element.Location as LocationPoint;
                // TaskDialog.Show("test", fam_location.Point.X.ToString() + "/n" + fam_location.Point.Y.ToString() + "/n" + fam_location.Point.Z.ToString());

                double beam_B = double.Parse(cop.beamLevel[one].Item3.Split('x')[1]);
                element.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-beam_B / 2).ToString());
                //TaskDialog.Show("test",beam_B.ToString());

                XYZ loc = new XYZ(0, 0, 0);


                //複製元件
                ICollection<ElementId> ass = ElementTransformUtils.CopyElement(doc, id, loc);

                Element new_ele = doc.GetElement(ass.First());
                FamilyInstance new_fami = new_ele as FamilyInstance;
                //指定symbol類型
                new_fami.Symbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().Where(x => x.Name == one_beam[one].Item3).First();
                //偏移元件
                new_ele.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString(ne_eleva.ToString());
                //寫備註
                new_ele.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(cop.section + "-" + one_beam[one].Item1.ToString() + "-" + "圍囹");
                //取消接合
                StructuralFramingUtils.DisallowJoinAtEnd(element as FamilyInstance, 0);
                StructuralFramingUtils.DisallowJoinAtEnd(element as FamilyInstance, 1);
                StructuralFramingUtils.DisallowJoinAtEnd(new_ele as FamilyInstance, 0);
                StructuralFramingUtils.DisallowJoinAtEnd(new_ele as FamilyInstance, 1);


            }

        }
    }
}