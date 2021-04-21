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

    class CreateWall : IExternalEventHandler
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
            
            //開始建立連續壁
            foreach(string file_path in files_path)
            {
                //API CODE START
                try
                {
                    //讀取資料
                    ExReader dex = new ExReader();
                    dex.SetData(file_path, 1);
                    try
                    {
                        dex.PassWallData();
                        dex.CloseEx();
                    }
                    catch (Exception e) { dex.CloseEx(); TaskDialog.Show("Error", e.Message); }

                    //開始建立連續壁
                    Transaction trans = new Transaction(doc);
                    trans.Start("交易開始");
                    //開挖各階之深度輸入
                    List<double> height = new List<double>();
                    foreach (var data in dex.excaLevel)
                        height.Add(data.Item2 * -1);

                    //偏移量
                    double xshift = xy_shift[0];
                    double yshift = xy_shift[1];

                    //建立開挖階數            
                    Level[] levlist = new Level[height.Count()];
                    for (int i = 0; i != height.Count(); i++)
                    {
                        levlist[i] = Level.Create(doc, height[i] * 1000 / 304.8);
                        levlist[i].Name = String.Format("斷面{0}-開挖階數" + (i + 1).ToString(), dex.section);
                    }
                    Level wall_level = Level.Create(doc, dex.wall_high * 1000 * -1 / 304.8);
                    wall_level.Name = String.Format("斷面{0}-擋土壁深度", dex.section);

                    //訂定開挖範圍
                    IList<CurveLoop> profileloops = new List<CurveLoop>();

                    //須回到原點
                    XYZ[] points = new XYZ[dex.excaRange.Count()];

                    //依照xlsx內給定的xy座標建立開挖範圍
                    //讀取座標建立points
                    for (int i = 0; i != dex.excaRange.Count(); i++)
                        points[i] = new XYZ(dex.excaRange[i].Item1 - xshift, dex.excaRange[i].Item2 - yshift, 0) * 1000 / 304.8;

                    //利用points建立線，用來作為建立牆的參數
                    CurveLoop profileloop = new CurveLoop();
                    IList<Curve> wall_profileloops = new List<Curve>();
                    for (int i = 0; i < points.Count() - 1; i++)
                    {
                        Line line = Line.CreateBound(points[i], points[i + 1]);
                        wall_profileloops.Add(line);
                        profileloop.Append(line);
                    }
                    profileloops.Add(profileloop);


                    //建立連續壁
                    IList<Curve> inner_wall_curves = new List<Curve>();
                    double wall_W = dex.wall_width * 1000; //連續壁厚度
                    WallType wallType = null;
                    List<Wall> inner_wall = new List<Wall>();
                    ICollection<WallType> walltype_familyinstance = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().ToList();
                    //檢查擋土壁
                    if (walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (dex.wall_width * 1000).ToString() + "mm").ToList().Count != 0)
                        wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁" && x.Name.Split('-')[1] == (dex.wall_width * 1000).ToString() + "mm").ToList().First();
                    //建立擋土壁新類型
                    if (wallType == null)
                    {
                        wallType = walltype_familyinstance.Where(x => x.Name.Split('-')[0] == "連續壁").ToList().First();
                        WallType new_wallFamSym = wallType.Duplicate("連續壁" + dex.wall_width * 1000 + "mm") as WallType;
                        CompoundStructure ly = new_wallFamSym.GetCompoundStructure();
                        ly.SetLayerWidth(0, dex.wall_width * 1000 / 304.8);
                        new_wallFamSym.SetCompoundStructure(ly);
                        wallType = new_wallFamSym;
                    }
                    //建立實體
                    for (int i = 0; i < points.Count<XYZ>() - 1; i++)
                    {
                        Curve c = wall_profileloops[i].CreateOffset(wall_W / 2 / 304.8, new XYZ(0, 0, -1));//此步驟為偏移擋土牆厚度1/2距離，作為建置參考線
                        Wall w = Wall.Create(doc, c, wallType.Id, wall_level.Id, dex.wall_high * 1000 / 304.8, 0, false, false);
                        w.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(dex.section + "-" + "連續壁");
                        inner_wall.Add(w);
                    }
                    trans.Commit();
                    //完成建立連續壁
                }
                catch (Exception e) { TaskDialog.Show("error", new StackTrace(e, true).GetFrame(0).GetFileLineNumber() + Environment.NewLine + e.Message + e.StackTrace); break; }

            }

            TaskDialog.Show("Done", "連續壁建置完畢");
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
