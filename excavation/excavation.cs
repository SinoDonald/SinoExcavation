using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using ExReaderConsole;

namespace excavation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Transaction trans = new Transaction(doc);
            trans.Start("交易開始");

            ExReader dex = new ExReader();
            dex.SetData("LGDE.xlsx", 1);
            dex.PassDE();
            dex.CloseEx();

            //訂定土體範圍            
            List<XYZ> topxyz = new List<XYZ>();
            XYZ t1 = new XYZ(-300, -300, 0);
            XYZ t2 = new XYZ(300, -300, 0);
            XYZ t3 = new XYZ(300, 300, 0);
            XYZ t4 = new XYZ(-300, 300, 0);
            topxyz.Add(t1 * 1000 / 304.8);
            topxyz.Add(t2 * 1000 / 304.8);
            topxyz.Add(t3 * 1000 / 304.8);
            topxyz.Add(t4 * 1000 / 304.8);
            TopographySurface.Create(doc, topxyz);

            //開挖深度所需參數
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(BuildingPadType));
            BuildingPadType bdtp = collector.FirstElement() as BuildingPadType;

            //開挖各階之深度輸入
            List<double> height = new List<double>();
            foreach (var data in dex.excaLevel)
                height.Add(data.Item2 * -1);


            //建立開挖階數            
            Level[] levlist = new Level[height.Count()];
            for (int i = 0; i != height.Count(); i++)
            {
                levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                levlist[i].Name = "開挖階數" + (i + 1).ToString();

            }

            //訂定開挖範圍
            IList<CurveLoop> profileloops = new List<CurveLoop>();
            IList<Curve> wall_profileloops = new List<Curve>();

            //須回到原點           

            XYZ[] points = new XYZ[dex.excaRange.Count()];

            for (int i = 0; i != dex.excaRange.Count(); i++)
                points[i] = new XYZ(dex.excaRange[i].Item1, dex.excaRange[i].Item2, 0) * 1000 / 304.8;

            //XYZ[] points = new XYZ[5];
            //points[0] = new XYZ(0, 0, 0) * 1000 / 304.8;
            //points[1] = new XYZ(160, 0, 0) * 1000 / 304.8;
            //points[2] = new XYZ(160, 36, 0) * 1000 / 304.8;
            //points[3] = new XYZ(0, 36, 0) * 1000 / 304.8;
            //points[4] = points[0];

            CurveLoop profileloop = new CurveLoop();
            for (int i = 0; i < points.Count() - 1; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                wall_profileloops.Add(line);
                profileloop.Append(line);
            }
            profileloops.Add(profileloop);
            Level levdeep = null;

            //建立開挖深度
            ICollection<Level> level_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().ToList();
            foreach (Level lev in level_familyinstance)
            {

                if (lev.Name == levlist[levlist.Count() - 1].Name)
                {
                    BuildingPad b = BuildingPad.Create(doc, bdtp.Id, lev.Id, profileloops);
                    levdeep = lev;
                }
            }

            IList<Curve> inner_wall_curves = new List<Curve>();

            //建立連續壁
            double wall_W = dex.wall_width * 1000; //連續壁厚度
            List<Wall> inner_wall = new List<Wall>();
            ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
            foreach (WallType walltype in walltype_familyinstance)
            {
                if (walltype.Name == "連續壁")
                {
                    for (int i = 0; i < points.Count<XYZ>() - 1; i++)
                    {
                        Curve c = wall_profileloops[i].CreateOffset(wall_W / 2 / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線
                        Wall w = Wall.Create(doc, c, walltype.Id, levdeep.Id, height.Min() * -1 * 1000 / 304.8, 0, false, false);
                        w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("連續壁");
                        inner_wall.Add(w);
                    }
                }
            }
            trans.Commit();

            //建立中間樁
            //取得連續壁內座標點

            //取得連續壁內座標點
            XYZ[] innerwall_points = new XYZ[inner_wall.Count<Wall>()];
            for (int i = 0; i < (inner_wall.Count<Wall>()); i++)
            {

                innerwall_points[i] = (inner_wall[i].Location as LocationCurve).Curve.Tessellate()[0];
                /*if (i == 0)
                {
                    LocationCurve c0 = inner_wall[inner_wall.Count<Wall>() - 1].Location as LocationCurve;
                    LocationCurve c1 = inner_wall[i].Location as LocationCurve;
                    innerwall_points[i] = new XYZ(points[i].X - 2 * (points[i].X - c0.Curve.GetEndPoint(1).X), points[0].Y - 2 * (points[i].Y - c1.Curve.GetEndPoint(0).Y), 0);
                }
                else if ((i) % 2 == 0)
                {
                    LocationCurve c0 = inner_wall[i - 1].Location as LocationCurve;
                    LocationCurve c1 = inner_wall[i].Location as LocationCurve;
                    innerwall_points[i] = new XYZ(points[i].X - 2 * (points[i].X - c0.Curve.GetEndPoint(1).X), points[i].Y - 2 * (points[i].Y - c1.Curve.GetEndPoint(0).Y), 0);
                }
                else
                {
                    LocationCurve c0 = inner_wall[i - 1].Location as LocationCurve;
                    LocationCurve c1 = inner_wall[i].Location as LocationCurve;
                    innerwall_points[i] = new XYZ(points[i].X - 2 * (points[i].X - c1.Curve.GetEndPoint(0).X), points[i].Y - 2 * (points[i].Y - c0.Curve.GetEndPoint(1).Y), 0);
                }*/
            }
            /*
            //取得所有XY數值
            List<double> Xs = new List<double>();
            List<double> Ys = new List<double>();
            for (int i = 0; i < (innerwall_points.Count<XYZ>()); i++)
            {
                Xs.Add(innerwall_points[i].X);
                Ys.Add(innerwall_points[i].Y);
            }


            double columns_dis = dex.centralCol[2] * 1000 / 304.8;//中間樁間距

            //開始建立中間樁
            ICollection<FamilySymbol> columns_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            foreach (FamilySymbol column_type in columns_familyinstance)
            {
                if (column_type.Name == "中間樁")
                {
                    for (int j = 0; j < (Ys.Max() / columns_dis - 1); j++)
                    {
                        for (int i = 0; i < (Xs.Max() / columns_dis - 1); i++)
                        {
                            XYZ column_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                            if (IsInPolygon(column_location, innerwall_points) == true)
                            {
                                FamilyInstance column_instance = doc.Create.NewFamilyInstance(column_location, column_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Column);
                                column_instance.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM).SetValueString("0");//給定中間樁長度
                                column_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set("中間樁");
                            }
                        }
                    }
                }
            }

            


            
            //折點判斷演算
            XYZ symb1 = new XYZ();
            XYZ symb2 = new XYZ();
            List<XYZ> ranoutpoint = new List<XYZ>();
            List<double> tan_value = new List<double>();
            for (int i = 0; i < wall_profileloops.Count; i++)
            {
                symb1 = wall_profileloops[i].GetEndPoint(1) - wall_profileloops[i].GetEndPoint(0);
                if (i  == wall_profileloops.Count<Curve>()-1)
                {
                    symb2 = wall_profileloops[0].GetEndPoint(1) - wall_profileloops[0].GetEndPoint(0);
                }
                else
                {
                    symb2 = wall_profileloops[i + 1].GetEndPoint(1) - wall_profileloops[i + 1].GetEndPoint(0);
                }
                XYZ vector = symb2.CrossProduct(symb1);
                if (vector.Z > 0 || symb1.X != 0 && symb1.Y != 0)
                {
                    
                    ranoutpoint.Add(wall_profileloops[i].GetEndPoint(1));
                }
                
                double theda = Math.Acos(symb1.Normalize().DotProduct(symb2.Normalize()))/ 2;
                if(vector.Z  < 0)
                    tan_value.Add(Math.Tan(theda)*-1);
                else
                    tan_value.Add(Math.Tan(theda));
                
            }
            tan_value.Insert(0, tan_value.Last());
            tan_value.RemoveAt(tan_value.Count-1);
            

            for (int lev = 0; lev!= dex.supLevel.Count(); lev++)
            {
                //建立圍囹
                //開始建立圍囹
                ICollection<FamilySymbol> beam_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
                foreach (FamilySymbol beam_type in beam_familyinstance)
                {
                    if (beam_type.Name == "H100x100")
                    {
                        double beam_H = double.Parse(beam_type.LookupParameter("H").AsValueString());
                        double beam_B = double.Parse(beam_type.LookupParameter("B").AsValueString());
                        for (int i = 0; i < points.Count<XYZ>() - 1; i++)
                        {
                            Curve c = wall_profileloops[i].CreateOffset((wall_W + beam_H) / 304.8, new XYZ(0, 0, -1));
                            FamilyInstance beam = doc.Create.NewFamilyInstance(c, beam_type, levlist[0], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                            beam.LookupParameter("斷面旋轉").SetValueString("90");
                            beam.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-圍囹");
                            
                            //判斷圍囹是否遇到折點，延伸長度不一
                            
                            if (calculator(wall_profileloops[i].GetEndPoint(0), ranoutpoint))
                            {
                                beam.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((+wall_W * tan_value[i]).ToString());
                            }
                            else
                            {
                                beam.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-wall_W).ToString());
                            }
                            if (calculator(wall_profileloops[i].GetEndPoint(1), ranoutpoint))
                            {
                                beam.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((+wall_W * tan_value[i+1]).ToString());   //起點歪的狀況自補
                            }
                            else
                            {
                                beam.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-wall_W).ToString());
                            }

                            //判斷圍囹之垂直深度，斜率零為負，反之為正
                            if ((c.GetEndPoint(0).Y - c.GetEndPoint(1).Y) / (c.GetEndPoint(0).X - c.GetEndPoint(1).X) == 0)
                            {
                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 - beam_B / 2).ToString());//2000為支撐階數深度，表1中
                            }
                            else
                            {
                                beam.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * 1000 + beam_B / 2).ToString());
                            }
                        }
                    }
                }*/

            /*
            //建立支撐
            XYZ frame_startpoint = null;
            XYZ frame_endpoint = null;

            //開始建立支撐
            ICollection<FamilySymbol> frame_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            foreach (FamilySymbol frame_type in frame_familyinstance)
            {
                if (frame_type.Name == "H100x100")
                {
                    double frame_H = double.Parse(frame_type.LookupParameter("H").AsValueString());
                    //X向支撐
                    for (int j = 0; j < (Ys.Max() / columns_dis - 1); j++)
                    {
                        for (int i = 0; i < (Xs.Max() / columns_dis - 1); i++)
                        {
                            XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                            frame_startpoint = intersection(frame_location, innerwall_points, points, true)[0];
                            if (IsInPolygon(frame_location, innerwall_points) == true)
                            {

                                try
                                {
                                    frame_endpoint = intersection(frame_location, innerwall_points, points, true)[1];
                                    Line line = Line.CreateBound(frame_startpoint, frame_endpoint);
                                    FamilyInstance frame_instance = doc.Create.NewFamilyInstance(line, frame_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    frame_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-支撐");
                                    //處理偏移與延伸問題
                                    frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * -1000).ToString());//2000為支撐階數深度，表1中
                                    frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-frame_H).ToString());
                                    frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-frame_H).ToString());
                                    frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-frame_H).ToString());

                                    //取消接合
                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                    //若為雙向支撐，則鏡射支撐             
                                    if (dex.supLevel[lev].Item3 == 2)
                                    {
                                        ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                    }
                                    break;
                                }
                                catch { }
                            }
                        }
                    }

                    //Y向支撐
                    for (int i = 0; i < (Xs.Max() / columns_dis - 1); i++)
                    {
                        for (int j = 0; j < (Ys.Max() / columns_dis - 1); j++)
                        {
                            XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (1 + j) * columns_dis, 0);
                            frame_startpoint = intersection(frame_location, innerwall_points, points, false)[0];                            
                            if (IsInPolygon(frame_location, innerwall_points) == true)
                            {
                                try
                                {
                                    frame_endpoint = intersection(frame_location, innerwall_points, points, false)[1];
                                    Line line = Line.CreateBound(frame_startpoint, frame_endpoint);
                                    FamilyInstance frame_instance = doc.Create.NewFamilyInstance(line, frame_type, levdeep, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                    frame_instance.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-支撐");

                                    //處理偏移與延伸問題
                                    frame_instance.get_Parameter(BuiltInParameter.Z_OFFSET_VALUE).SetValueString((dex.supLevel[lev].Item2 * -1000 + frame_H).ToString());//2000為支撐階數深度，表1中
                                    frame_instance.get_Parameter(BuiltInParameter.Y_OFFSET_VALUE).SetValueString((-frame_H).ToString());
                                    frame_instance.get_Parameter(BuiltInParameter.START_EXTENSION).SetValueString((-frame_H).ToString());
                                    frame_instance.get_Parameter(BuiltInParameter.END_EXTENSION).SetValueString((-frame_H).ToString());

                                    //取消接合
                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 0);
                                    StructuralFramingUtils.DisallowJoinAtEnd(frame_instance, 1);

                                    //若為雙向支撐，則鏡射支撐
                                    if (dex.supLevel[lev].Item3 == 2)
                                    {
                                        ElementTransformUtils.MirrorElement(doc, frame_instance.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                    }
                                    break;
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            */
            /*

            //建立斜撐
            ICollection<FamilySymbol> slopframe_symbol = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            foreach (FamilySymbol slopframe_type in slopframe_symbol)
            {
                if (slopframe_type.Name == "斜撐")
                {
                    //X向斜撐
                    for (int j = 0; j < (Ys.Max() / columns_dis - 1); j++)
                    {
                        for (int i = 0; i < (Xs.Max() / columns_dis - 1); i++)
                        {
                            XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (j + 1) * columns_dis, 0);
                            frame_startpoint = intersection(frame_location, innerwall_points, points, true)[0];
                            if (IsInPolygon(frame_location, innerwall_points) == true)
                            {
                                frame_endpoint = intersection(frame_location, innerwall_points, points, true)[1];
                                FamilyInstance slopframe_1 = doc.Create.NewFamilyInstance(frame_startpoint, slopframe_type, levlist[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                FamilyInstance slopframe_2 = doc.Create.NewFamilyInstance(frame_endpoint, slopframe_type, levlist[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                slopframe_1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-斜撐");
                                slopframe_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-斜撐");
                                slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(levlist[lev].Elevation-dex.supLevel[lev].Item2*-1000/ 304.8 + slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2));

                                slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(levlist[lev].Elevation - dex.supLevel[lev].Item2 * -1000/ 304.8 + slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2));

                                //旋轉斜撐元件
                                Line rotate_line_s = Line.CreateBound(frame_startpoint, frame_startpoint + new XYZ(0, 0, 1));
                                Line rotate_line_e = Line.CreateBound(frame_endpoint, frame_endpoint + new XYZ(0, 0, 1));
                                slopframe_1.Location.Rotate(rotate_line_s, 1.5 * Math.PI);
                                slopframe_2.Location.Rotate(rotate_line_e, 0.5 * Math.PI);

                                //鏡射斜撐元件
                                if (dex.supLevel[lev].Item3 == 2)//雙排
                                {
                                    ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint));
                                    ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_endpoint));
                                }
                                else//單排
                                {
                                    ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_startpoint.Add(new XYZ(0, -(slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2 + slopframe_1.LookupParameter("中間樁直徑").AsDouble()/2), 0))));
                                    ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisY, frame_endpoint));
                                    slopframe_2.Location.Move((new XYZ(0,-(slopframe_1.LookupParameter("支撐厚度").AsDouble() + slopframe_1.LookupParameter("中間樁直徑").AsDouble()), 0)));

                                }

                                break;
                            }
                        }
                    }

                    //Y向斜撐
                    for (int i = 0; i < (Xs.Max() / columns_dis - 1); i++)
                    {
                        for (int j = 0; j < (Ys.Max() / columns_dis - 1); j++)
                        {
                            XYZ frame_location = new XYZ(Xs.Min() + (i + 1) * columns_dis, Ys.Min() + (1 + j) * columns_dis, 0);
                            frame_startpoint = intersection(frame_location, innerwall_points, points, false)[0];
                            if (IsInPolygon(frame_location, innerwall_points) == true)
                            {
                                frame_endpoint = intersection(frame_location, innerwall_points, points, false)[1];
                                FamilyInstance slopframe_1 = doc.Create.NewFamilyInstance(frame_startpoint, slopframe_type, levlist[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                FamilyInstance slopframe_2 = doc.Create.NewFamilyInstance(frame_endpoint, slopframe_type, levlist[lev], Autodesk.Revit.DB.Structure.StructuralType.Beam);
                                slopframe_1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-斜撐");
                                slopframe_2.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set((lev + 1).ToString() + "-斜撐");
                                slopframe_1.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(levlist[lev].Elevation  - dex.supLevel[lev].Item2 * -1000/304.8 - slopframe_1.LookupParameter("支撐厚度").AsDouble()/2));

                                slopframe_2.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM).Set(Math.Abs(levlist[lev].Elevation  - dex.supLevel[lev].Item2 * -1000/304.8 - slopframe_2.LookupParameter("支撐厚度").AsDouble()/2));

                                //旋轉斜撐元件
                                Line rotate_line = Line.CreateBound(frame_endpoint, frame_endpoint + new XYZ(0, 0, 1));
                                slopframe_2.Location.Rotate(rotate_line, Math.PI);

                                //鏡射斜撐元件
                                if (dex.supLevel[lev].Item3 == 2)//雙排
                                {
                                    ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint));
                                    ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_endpoint));
                                }
                                else//單排
                                {
                                    ElementTransformUtils.MirrorElement(doc, slopframe_1.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_startpoint.Add(new XYZ(slopframe_1.LookupParameter("支撐厚度").AsDouble() / 2 + slopframe_1.LookupParameter("中間樁直徑").AsDouble()/2,0, 0))));
                                    ElementTransformUtils.MirrorElement(doc, slopframe_2.Id, Plane.CreateByNormalAndOrigin(XYZ.BasisX, frame_endpoint));
                                    slopframe_2.Location.Move((new XYZ((slopframe_1.LookupParameter("支撐厚度").AsDouble() + slopframe_1.LookupParameter("中間樁直徑").AsDouble()), 0, 0)));

                                }
                                break;
                            }
                        }
                    }
                }
            }*/
        
            

            return Result.Succeeded;
        }
        public static bool calculator(XYZ checkPoint, List<XYZ> ranoutpoints)//判斷點與點是否為相同點，因為xyz類型無法做equals
        {
            bool a = false;
            foreach (XYZ ranoutpoint in ranoutpoints)
            {
                if (ranoutpoint.X - checkPoint.X < 1 && ranoutpoint.X - checkPoint.X > -1)
                {
                    if (ranoutpoint.Y - checkPoint.Y < 1 && ranoutpoint.Y - checkPoint.Y > -1)
                    {
                        if (ranoutpoint.Z - checkPoint.Z < 1 && ranoutpoint.Z - checkPoint.Z > -1)
                        {
                            a = true;
                        }
                    }
                }
            }
            return a;
        }
        public static bool IsInPolygon(XYZ checkPoint, XYZ[] polygonPoints)//判斷點是否位於開挖範圍內
        {
            bool inside = false;
            int pointCount = polygonPoints.Count<XYZ>();
            XYZ p1, p2;
            for (int i = 0, j = pointCount - 1; i < pointCount; j = i, i++)
            {
                p1 = polygonPoints[i];
                p2 = polygonPoints[j];
                if (checkPoint.Y < p2.Y)
                {
                    if (p1.Y <= checkPoint.Y)
                    {
                        if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) > (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                        {

                            inside = (!inside);
                        }
                    }
                }
                else if (checkPoint.Y < p1.Y)
                {

                    if ((checkPoint.Y - p1.Y) * (p2.X - p1.X) < (checkPoint.X - p1.X) * (p2.Y - p1.Y))
                    {
                        inside = (!inside);
                    }
                }
            }
            return inside;
        }
        public static XYZ[] intersection(XYZ checkPoint, XYZ[] polygonPoints, XYZ[] checkpolygonPoints, bool x_or_y)//計算支撐與連續壁交點，true為計算X向，false為計算Y向
        {
            List<XYZ> intersection_points = new List<XYZ>();
            List<double> x = new List<double>();
            List<double> y = new List<double>();
            if (x_or_y == true)
            {
                for (int i = 1; i < polygonPoints.Count<XYZ>(); i++)
                {
                    double cheak_b = checkPoint.Y;
                    if (i + 1 == polygonPoints.Count<XYZ>() && (polygonPoints[i].X == polygonPoints[0].X))
                    {
                        XYZ intersection_point = new XYZ(polygonPoints[0].X, checkPoint.Y, 0);
                        if (IsInPolygon(intersection_point, checkpolygonPoints))
                        {
                            intersection_points.Add(intersection_point);
                        }
                    }
                    else if ((polygonPoints[i].X == polygonPoints[i - 1].X))
                    {
                        XYZ intersection_point = new XYZ(polygonPoints[i].X, checkPoint.Y, 0);
                        if (IsInPolygon(intersection_point, checkpolygonPoints))
                        {
                            intersection_points.Add(intersection_point);
                        }
                    }
                }
                foreach (XYZ point in intersection_points)
                {
                    x.Add(point.X);
                    y.Add(point.Y);
                }
            }

            else//若為False則計算Y向
            {
                for (int i = 1; i < polygonPoints.Count<XYZ>(); i++)
                {
                    double cheak_b = checkPoint.Y;
                    if (i + 1 == polygonPoints.Count<XYZ>())
                    {
                        XYZ intersection_point = new XYZ(checkPoint.X, polygonPoints[0].Y, 0);
                        if (IsInPolygon(intersection_point, checkpolygonPoints))
                        {
                            intersection_points.Add(intersection_point);
                        }
                    }
                    else if (((polygonPoints[i].X - polygonPoints[i - 1].X) / (polygonPoints[i].Y - polygonPoints[i - 1].Y)) == 0)
                    {
                        XYZ intersection_point = new XYZ(checkPoint.X, polygonPoints[i].Y, 0);
                        if (IsInPolygon(intersection_point, checkpolygonPoints))
                        {
                            intersection_points.Add(intersection_point);
                        }
                    }
                }
                foreach (XYZ point in intersection_points)
                {
                    x.Add(point.X);
                    y.Add(point.Y);
                }
            }
            XYZ[] intersection = new XYZ[2];
            intersection[0] = new XYZ(x.Min(), y.Min(), 0);
            intersection[1] = new XYZ(x.Max(), y.Max(), 0);
            return intersection;
        }
    }
}

