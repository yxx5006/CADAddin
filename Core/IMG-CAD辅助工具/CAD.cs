using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;
using MSWord = Microsoft.Office.Interop.Word;

namespace IMG_CAD辅助工具
{
    public static class Cad
    {
        public static List<string> ImgLayers { get; set; }//保存公司内部图层
        private static void GetImgLayers(Database acCurDb)
        {
            List<string> Alllayers = new List<string>();
            using (OpenCloseTransaction trans = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {
                LayerTable lt = trans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                Alllayers.AddRange(from ObjectId id in lt select trans.GetObject(id, OpenMode.ForRead) as LayerTableRecord into ltr select ltr.Name);
                string strexp = @"^0-BIM-";
                Regex regex = new Regex(strexp);
                ImgLayers = Alllayers.Where(layer => regex.IsMatch(layer)).ToList();
                trans.Commit();
            }
        }

        private static List<object> GetLayerObjectsByLayerName(List<string> ImgLayers)
        {
            List<object> objects = new List<object>();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            if (ImgLayers.Count != 0)
            {
                using (OpenCloseTransaction trans = acCurDb.TransactionManager.StartOpenCloseTransaction())
                {
                    foreach (string layerName in ImgLayers)
                    {
                        try
                        {
                            TypedValue[] typedValue = new TypedValue[]
                            {
                                new TypedValue((int)DxfCode.LayerName, layerName)
                            };
                            SelectionFilter filter = new SelectionFilter(typedValue);
                            PromptSelectionResult result = ed.SelectAll(filter);
                            SelectionSet set = result.Value;
                            ObjectId[] ids = set.GetObjectIds();
                            foreach (ObjectId entid in ids)
                            {
                                objects.Add(trans.GetObject(entid, OpenMode.ForRead));
                            }
                        }
                        catch (Exception)
                        {

                            ed.WriteMessage("\n获取信息出错。");
                        }
                    }
                    trans.Commit();
                }

            }
            return objects;
        }

        [CommandMethod("test")]
        public static void test()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            ProgressMeter pm = new ProgressMeter();
            MSWord.Application WordApp;
            MSWord.Document WordDoc;
            string filePath = Path.GetDirectoryName(acCurDb.Filename) + "\\" + Path.GetFileNameWithoutExtension(acCurDb.Filename) + "-问题记录.docx";
            Stopwatch sw = new Stopwatch();
            using (Transaction trans = acCurDb.TransactionManager.StartTransaction())
            {
                sw.Start();
                Word.CreatDocFile(filePath);
                //GetImgLayers(acCurDb);

                #region 键值对测试

                Dictionary<string, string> datas = new Dictionary<string, string>();
                datas.Add("{ProjectName}", "测试工程名称");
                datas.Add("{SubProjectName}", "测试子项名称");
                datas.Add("{RecName}", "姓名");
                datas.Add("{RecDate}", "2016/01/01");
                datas.Add("{Level}", "A");
                datas.Add("{Sort}", "3");
                datas.Add("{No}", "TJ002");
                datas.Add("{DrawingName}", "五层结构平面布置图");
                datas.Add("{Disc}", "阿斯兰的咖啡碱阿历克斯的风景阿里山扩阿斯兰的咖啡碱阿历克斯的风景阿里山扩大阿斯兰的咖啡碱阿历克斯的风景阿里山扩大大");
                datas.Add("{Location}", "景阿里山扩大大");
                datas.Add("{Reply}", "测试");
                datas.Add("{Response}", "姓名");

                #endregion

                Word.OpenDocFile(out WordApp, out WordDoc, filePath);
                pm.Start("创建问题页面");
                pm.SetLimit(100);
                Word.CopyTable(WordApp, WordDoc, 1);
                for (int i = 0; i < 5; i++)
                {

                    pm.MeterProgress();

                    Word.ReplaceToWord(WordApp, WordDoc, datas);
                    //Word.CopyTable(WordApp, WordDoc, 1);
                    System.Windows.Forms.Application.DoEvents();
                }
                pm.Stop();
                Word.CloseAndSaveDocFile(WordApp, WordDoc);
                sw.Stop();
                ed.WriteMessage(sw.Elapsed.ToString());
            }

        }

        [CommandMethod("test1")]
        public static void test1()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            GetImgLayers(acCurDb);
            List<object> objects = GetLayerObjectsByLayerName(ImgLayers);
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            foreach (object Object in objects)
            {
                if (Object is MText)
                {
                    ed.WriteMessage("\n当前是多行文字");
                }
                else if (Object is DBText)
                {
                    ed.WriteMessage("\n当前是多行文字");
                }
                else if (Object is Ole2Frame)
                {
                    ed.WriteMessage("\n当前是OLE对象");
                    Ole2Frame ole = (Ole2Frame)Object;

                    ed.WriteMessage("\n"+ole.LinkName);



                    //MSWord.Application WordApp;
                    //MSWord.Document WordDoc;
                    //string filePath = Path.GetDirectoryName(acCurDb.Filename) + "\\" + Path.GetFileNameWithoutExtension(acCurDb.Filename) + "-问题记录.docx";
                    //Word.CreatDocFile(filePath);
                    //Word.OpenDocFile(out WordApp, out WordDoc, filePath);
                    //Word.CopyTable(WordApp, WordDoc, 1);
                    //Ole2Frame ole = (Ole2Frame)Object;

                    //Clipboard.SetDataObject(ole,true);
                    //Clipboard.SetDataObject(ole);
                    //WordDoc.Tables[1].Cell(5, 1).Range.Paste();
                    //break;
                }
            }
        }
    }
}
