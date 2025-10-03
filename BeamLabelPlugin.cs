using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(BeamLabelPlugin.BeamLabelCommands))]

namespace BeamLabelPlugin
{
    public class BeamLabelCommands
    {
        [CommandMethod("BEAMLABEL")]
        public void InsertBeamLabels()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // 提示用戶選擇編號資料檔案
                PromptStringOptions pso = new PromptStringOptions("\n請輸入編號資料檔案路徑 (或按Enter使用預設路徑): ");
                pso.AllowSpaces = true;
                PromptResult pr = ed.GetString(pso);
                
                string filePath = pr.StringResult;
                if (string.IsNullOrEmpty(filePath))
                {
                    filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ETABS_梁編號_分頁.xlsx");
                }

                if (!File.Exists(filePath))
                {
                    ed.WriteMessage($"\n錯誤：找不到檔案 {filePath}");
                    return;
                }

                // 讀取編號資料
                List<BeamLabelData> beamLabels = ReadBeamLabelData(filePath);
                
                if (beamLabels.Count == 0)
                {
                    ed.WriteMessage("\n警告：沒有找到有效的編號資料");
                    return;
                }

                ed.WriteMessage($"\n找到 {beamLabels.Count} 個梁編號，開始插入...");

                // 開始事務
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 獲取當前圖層
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId layerId = lt["0"]; // 使用預設圖層

                    // 創建文字樣式
                    TextStyleTable ts = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                    ObjectId textStyleId = ts["Standard"];

                    foreach (BeamLabelData beam in beamLabels)
                    {
                        // 創建文字物件
                        DBText text = new DBText();
                        text.Position = new Point3d(beam.X, beam.Y, 0);
                        text.TextString = beam.Label;
                        text.Height = beam.TextHeight;
                        text.LayerId = layerId;
                        text.TextStyleId = textStyleId;
                        text.ColorIndex = beam.ColorIndex;

                        // 根據梁的方向設定文字旋轉角度
                        if (beam.IsVertical)
                        {
                            text.Rotation = Math.PI / 2; // 90度
                        }
                        else if (beam.Angle != 0)
                        {
                            text.Rotation = beam.Angle;
                        }

                        // 將文字加入圖面
                        BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                        btr.AppendEntity(text);
                        tr.AddNewlyCreatedDBObject(text, true);
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\n成功插入 {beamLabels.Count} 個梁編號！");
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n錯誤：{ex.Message}");
            }
        }

        [CommandMethod("BEAMLABELSETUP")]
        public void SetupBeamLabels()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 創建專用圖層
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);
                    
                    // 大梁圖層
                    if (!lt.Has("BEAM_MAIN"))
                    {
                        LayerTableRecord ltr = new LayerTableRecord();
                        ltr.Name = "BEAM_MAIN";
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1); // 紅色
                        ObjectId layerId = lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }

                    // 小梁圖層
                    if (!lt.Has("BEAM_SECONDARY"))
                    {
                        LayerTableRecord ltr = new LayerTableRecord();
                        ltr.Name = "BEAM_SECONDARY";
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 3); // 綠色
                        ObjectId layerId = lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }

                    // 特殊梁圖層
                    if (!lt.Has("BEAM_SPECIAL"))
                    {
                        LayerTableRecord ltr = new LayerTableRecord();
                        ltr.Name = "BEAM_SPECIAL";
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 6); // 洋紅色
                        ObjectId layerId = lt.Add(ltr);
                        tr.AddNewlyCreatedDBObject(ltr, true);
                    }

                    tr.Commit();
                }

                ed.WriteMessage("\n梁編號圖層設定完成！");
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\n錯誤：{ex.Message}");
            }
        }

        private List<BeamLabelData> ReadBeamLabelData(string filePath)
        {
            List<BeamLabelData> beamLabels = new List<BeamLabelData>();

            try
            {
                // 這裡需要根據您的Excel格式來解析資料
                // 假設Excel格式包含：ETABS編號, 編號, 樓層, X座標, Y座標
                
                // 使用Excel讀取庫（如EPPlus或ClosedXML）
                // 這裡提供一個簡化的範例
                
                // 實際實作時，您需要：
                // 1. 讀取Excel檔案
                // 2. 解析座標資料
                // 3. 判斷梁的類型（大梁、小梁、特殊梁）
                // 4. 設定對應的樣式參數

                // 範例資料（實際使用時需要從Excel讀取）
                beamLabels.Add(new BeamLabelData
                {
                    Label = "B1-1",
                    X = 0,
                    Y = 0,
                    TextHeight = 2.5,
                    ColorIndex = 1,
                    IsVertical = false,
                    Angle = 0
                });

            }
            catch (Exception ex)
            {
                throw new Exception($"讀取編號資料時發生錯誤：{ex.Message}");
            }

            return beamLabels;
        }
    }

    public class BeamLabelData
    {
        public string Label { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double TextHeight { get; set; }
        public short ColorIndex { get; set; }
        public bool IsVertical { get; set; }
        public double Angle { get; set; }
    }
}
