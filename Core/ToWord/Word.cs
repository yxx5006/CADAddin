using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Exception = System.Exception;
using MSWord = Microsoft.Office.Interop.Word;

namespace ToOffice
{
    public class Word
    {
        private static object _nothing = Missing.Value;
        private static Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        private static Document acDoc = Application.DocumentManager.MdiActiveDocument;
        private static Database acCurDb = acDoc.Database;
        private static List<string> IMGLayers = new List<string>();

        /// <summary>
        /// 初始化Word页面
        /// </summary>
        /// <param name="wordApp">传入WordApp</param>
        /// <param name="wordDoc">传入WordDoc</param>
        private static void SetPage(MSWord.Application wordApp, MSWord.Document wordDoc)
        {
            try
            {
                wordDoc.PageSetup.Orientation = MSWord.WdOrientation.wdOrientLandscape;
                wordDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(0.8f);
                wordDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(0.5f);
                wordDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(2.5f);
                wordDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.5f);
                wordDoc.PageSetup.PageHeight = wordApp.CentimetersToPoints(21);
                wordDoc.PageSetup.PageWidth = wordApp.CentimetersToPoints(29.7f);
                wordDoc.PageSetup.HeaderDistance = wordApp.CentimetersToPoints(1);
                wordDoc.PageSetup.FooterDistance = wordApp.CentimetersToPoints(0.5f);
            }
            catch (Exception)
            {

                ed.WriteMessage("\n设置Word页面发生未知错误，程序中断。如需帮助请联系作者。");
                wordApp.Quit(ref _nothing, ref _nothing, ref _nothing);
            }
        }

        /// <summary>
        /// 设置页眉
        /// </summary>
        /// <param name="wordApp">传入WordApp</param>
        /// <param name="wordDoc">传入WordDoc</param>
        private static void AddPageHeaderFooter(MSWord.Application wordApp, MSWord.Document wordDoc)
        {

            try
            {
                Bitmap png = Pic.pic;
                png.Save(Path.GetTempPath() + "\\pic.png");
                if (wordApp.ActiveWindow.ActivePane.View.Type == MSWord.WdViewType.wdNormalView ||
                    wordApp.ActiveWindow.ActivePane.View.Type == MSWord.WdViewType.wdOutlineView)
                {
                    wordApp.ActiveWindow.ActivePane.View.Type = MSWord.WdViewType.wdPrintView;
                }
                wordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekCurrentPageHeader;
                wordApp.Selection.HeaderFooter.LinkToPrevious = false;
                MSWord.InlineShape il =
                    wordApp.ActiveWindow.ActivePane.Selection.InlineShapes.AddPicture(Path.GetTempPath() + "\\pic.png",
                        ref _nothing, ref _nothing, ref _nothing);
                il.ScaleHeight = 134f;
                il.ScaleWidth = 238f;
                wordApp.Selection.HeaderFooter.Range.ParagraphFormat.Alignment =
                    MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                wordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
            }
            catch (Exception)
            {

                ed.WriteMessage("\n设置Word页面发生未知错误，程序中断。如需帮助请联系作者。");
                wordApp.Quit(ref _nothing, ref _nothing, ref _nothing);
            }
        }

        /// <summary>
        /// 创建表格并作为后续的模板
        /// </summary>
        /// <param name="wordApp">传入WordApp</param>
        /// <param name="wordDoc">传入WordDoc</param>
        private static void CreatFristTable(MSWord.Application wordApp, MSWord.Document wordDoc)
        {
            try
            {
                MSWord.Table table = wordDoc.Tables.Add(wordApp.Selection.Range, 9, 10, ref _nothing, ref _nothing);
                table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineWidth = MSWord.WdLineWidth.wdLineWidth150pt;

                table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                table.Borders.InsideLineWidth = MSWord.WdLineWidth.wdLineWidth050pt;

                table.Range.Cells.Height = 17.142f;
                table.Range.Cells.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;
                table.Range.Font.Name = "黑体";
                table.Range.Font.Size = 10.5f;
                table.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                table.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                table.Cell(2, 9).Range.Font.Size = 14;
                table.Cell(2, 9).Range.Font.ColorIndex = MSWord.WdColorIndex.wdRed;

                table.Cell(4, 1).Range.Cells.Height = 34.284f;
                table.Cell(4, 1).Range.Cells.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;
                table.Cell(5, 1).Range.Cells.Height = 171.42f;
                table.Cell(5, 1).Range.Cells.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;
                table.Cell(6, 1).Range.Cells.Height = 171.42f;
                table.Cell(6, 1).Range.Cells.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;

                table.Cell(1, 2).Merge(table.Cell(1, 10));
                table.Cell(3, 2).Merge(table.Cell(3, 6));
                table.Cell(4, 2).Merge(table.Cell(4, 10));
                table.Cell(5, 1).Merge(table.Cell(5, 10));
                table.Cell(6, 1).Merge(table.Cell(6, 10));
                table.Cell(7, 2).Merge(table.Cell(7, 8));
                table.Cell(8, 2).Merge(table.Cell(8, 8));
                table.Cell(9, 2).Merge(table.Cell(9, 3));
                table.Cell(9, 3).Merge(table.Cell(9, 4));
                table.Cell(9, 4).Merge(table.Cell(9, 6));
                table.Cell(9, 5).Merge(table.Cell(9, 6));

                object unit = MSWord.WdUnits.wdLine;
                object count = 1;
                object extend = MSWord.WdMovementType.wdExtend;
                table.Cell(2, 9).Select();
                wordApp.Selection.MoveDown(ref unit, ref count, ref extend);
                wordApp.Selection.Cells.Merge();
                table.Cell(2, 10).Select();
                wordApp.Selection.MoveDown(ref unit, ref count, ref extend);
                wordApp.Selection.Cells.Merge();
                table.Cell(5, 1).Select();
                wordApp.Selection.MoveDown(ref unit, ref count, ref extend);
                wordApp.Selection.Cells.Merge();

                table.Cell(3, 2).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                table.Cell(4, 2).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                table.Cell(6, 2).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                table.Cell(7, 2).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;

                table.Cell(1, 1).Range.Text = "项目名称";
                table.Cell(2, 1).Range.Text = "记录信息";
                table.Cell(2, 3).Range.Text = "记录人";
                table.Cell(2, 5).Range.Text = "记录日期";
                table.Cell(2, 7).Range.Text = "问题等级";
                table.Cell(2, 9).Range.Text = "编号";
                table.Cell(2, 9).Range.Font.ColorIndex = MSWord.WdColorIndex.wdRed;
                table.Cell(2, 10).Range.Font.ColorIndex = MSWord.WdColorIndex.wdRed;
                table.Cell(2, 9).Range.Font.Bold = 1;
                table.Cell(3, 1).Range.Text = "图号图名";
                table.Cell(3, 3).Range.Text = "问题分类";
                table.Cell(4, 1).Range.Text = "问题描述";
                table.Cell(6, 1).Range.Text = "图纸定位";
                table.Cell(6, 2).Range.Cells.SetWidth(504.2605f, MSWord.WdRulerStyle.wdAdjustFirstColumn);
                table.Cell(6, 3).Range.Text = "对应问题编号";
                table.Cell(7, 1).Range.Text = "答复意见";
                table.Cell(7, 2).Range.Cells.SetWidth(504.2605f, MSWord.WdRulerStyle.wdAdjustFirstColumn);
                table.Cell(7, 3).Range.Text = "答复人";
                table.Cell(8, 1).Range.Text = "解决情况";
                table.Cell(8, 2).Range.Text = "□未回复";
                table.Cell(8, 3).Range.Text = "□已回复（未附变更单）";
                table.Cell(8, 4)
                    .Range.Cells.SetWidth(504.2605f - table.Cell(8, 2).Width - table.Cell(8, 3).Width,
                        MSWord.WdRulerStyle.wdAdjustFirstColumn);
                table.Cell(8, 4).Range.Text = "□已回复（附变更单）";
                table.Cell(8, 5).Range.Text = "□已解决";

            }
            catch (Exception)
            {
                ed.WriteMessage("\n创建表格失败，程序异常中断。如需帮助请联系作者。");
                wordApp.Quit(ref _nothing, ref _nothing, ref _nothing);
            }
        }

        /// <summary>
        /// 获得图层上所有实体
        /// </summary>
        /// <param name="layerName">图层名</param>
        /// <returns></returns>
        private static List<MText> GetLayerDbtxtByLayerName(List<string> layerNames)
        {

            List<MText> txts = new List<MText>();
            if (layerNames.Count != 0)
            {
                using (OpenCloseTransaction trans = acCurDb.TransactionManager.StartOpenCloseTransaction())
                {
                    foreach (string layerName in layerNames)
                    {
                        try
                        {
                            TypedValue[] typedValue = new TypedValue[]
                            {
                                new TypedValue((int)DxfCode.LayerName, layerName),

                                //new TypedValue((int)DxfCode.Start, "MTEXT"),

                            };
                            SelectionFilter filter = new SelectionFilter(typedValue);
                            PromptSelectionResult result = ed.SelectAll(filter);
                            //if (result.Value == null)
                            //{
                            //    return txts;
                            //}
                            SelectionSet set = result.Value;
                            ObjectId[] ids = set.GetObjectIds();
                            foreach (ObjectId entid in ids)
                            {
                                //txts.Add(trans.GetObject(entid, OpenMode.ForRead) as DBText);
                                txts.Add(trans.GetObject(entid, OpenMode.ForRead) as MText);
                            }
                        }
                        catch (Exception)
                        {

                            ed.WriteMessage("\n获取文本信息出错。");
                        }
                    }
                    trans.Commit();
                }

            }
            var temp = from mtext in txts
                       orderby mtext.Location.Y descending, mtext.Location.X
                       select mtext;
            txts = temp.ToList();
            return txts;
        }
        /// <summary>
        /// 获得IMG内部图层
        /// </summary>
        /// <returns>List<string> 图层集合</returns>
        public static List<string> GetIMGLayers()
        {
            List<string> Alllayers = new List<string>();
            using (OpenCloseTransaction trans = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {
                LayerTable lt = trans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                Alllayers.AddRange(from ObjectId id in lt select trans.GetObject(id, OpenMode.ForRead) as LayerTableRecord into ltr select ltr.Name);
                string strexp = @"^0-BIM-";
                Regex regex = new Regex(strexp);
                IMGLayers = Alllayers.Where(layer => regex.IsMatch(layer)).ToList();
                trans.Commit();
            }
            return IMGLayers;
        }

        private static void CreatWordApp(out MSWord.Application WordApp, out MSWord.Document WordDoc)
        {

            WordApp = new MSWord.Application();
            WordDoc = new MSWord.Document();

            WordDoc = WordApp.Documents.Add(ref _nothing, ref _nothing, ref _nothing, ref _nothing);
        }


        [CommandMethod("tw")]
        public static void ToWord()
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();
            ed.WriteMessage("\n获取问题记录，请稍后...");
            IMGLayers = GetIMGLayers();
            // List<DBText> ents = GetLayerDbtxtByLayerName(IMGLayers);
            List<MText> txts = GetLayerDbtxtByLayerName(IMGLayers);

            if (txts.Count == 0)
            {
                watch.Stop();
                ed.WriteMessage("\n未发现任何问题记录，请核对。");
            }
            else
            {

                MSWord.Application WordApp;
                MSWord.Document WordDoc;
                CreatWordApp(out WordApp, out WordDoc);
                object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;
                object filePath = Path.GetDirectoryName(acCurDb.Filename) + "\\" +
                                  Path.GetFileNameWithoutExtension(acCurDb.Filename) + "-问题记录.docx";
                object endkeyunit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
                object unit = Microsoft.Office.Interop.Word.WdUnits.wdLine;
                object count = 1;

                List<string> txtcontent = (from txt in txts where txt != null select txt.Contents).ToList();



                try
                {
                    ed.WriteMessage("\n设置IMG-BIM问题记录模板.请稍后...");
                    SetPage(WordApp, WordDoc);
                    AddPageHeaderFooter(WordApp, WordDoc);
                    CreatFristTable(WordApp, WordDoc);
                    ed.WriteMessage("\n模板创建完成，正在写入数据.请稍后...");
                }
                catch (Exception)
                {
                    ed.WriteMessage("\n创建Word过程发生错误，程序异常退出！");
                }
                WordDoc.Tables[1].Cell(1, 2).Range.Text = "测试工程名";
                ed.WriteMessage(txtcontent.Count.ToString());
                for (int i = 1; i < txtcontent.Count / 3; i++)
                {
                    WordDoc.Tables[1].Select();
                    WordApp.Selection.Copy();
                    WordApp.Selection.EndKey(ref endkeyunit, ref _nothing);
                    //WordApp.Selection.MoveUp(ref unit, ref count, ref _nothing);
                    object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
                    WordApp.Selection.InsertBreak(ref pBreak);
                    WordApp.Selection.Paste();

                }
                for (int i = 1; i < txtcontent.Count / 3 + 1; i++)
                {
                    for (int j = i * 3 - 3; j < i * 3; j++)
                    {
                        WordDoc.Tables[i].Cell(2, 4).Range.Text = "叶舒帆";
                        switch (j % 3)
                        {
                            case 0:
                                WordDoc.Tables[i].Cell(4, 2).Range.Text = txtcontent[j];
                                break;
                            case 1:
                                WordDoc.Tables[i].Cell(2, 10).Range.Text = txtcontent[j];
                                break;
                            case 2:
                                WordDoc.Tables[i].Cell(2, 6).Range.Text = txtcontent[j];
                                break;
                        }
                    }
                }

                WordDoc.SaveAs(ref filePath, ref format, ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing,
                    ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing,
                    ref _nothing, ref _nothing);
                WordDoc.Close(ref _nothing, ref _nothing, ref _nothing);
                WordApp.Quit(ref _nothing, ref _nothing, ref _nothing);
                watch.Stop();
                ed.WriteMessage("\n问题导出完成耗时：{0} ms，请到图纸目录下查看。", watch.ElapsedMilliseconds.ToString());
            }

        }
    }
}
