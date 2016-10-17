
using System.Diagnostics;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
namespace 钢筋文字修正
{
    public class ChangeText
    {
        [CommandMethod("CRT")]
        public static void ChangReinText()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = acDoc.Editor;
            ed.WriteMessage("\n异常钢筋符号快速替换 2016/10/14");
            Stopwatch sw=new Stopwatch();
            sw.Start();
            using (OpenCloseTransaction acTrams = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {
                TypedValue[] typedValue = new TypedValue[] { new TypedValue((int)DxfCode.Start, "TEXT") };
                SelectionFilter filter = new SelectionFilter(typedValue);
                PromptSelectionResult result = ed.SelectAll(filter);
                if (result.Status == PromptStatus.OK)
                {
                    SelectionSet set = result.Value;
                    foreach (SelectedObject text in set)
                    {
                        if (text != null)
                        {
                            DBText acText = acTrams.GetObject(text.ObjectId, OpenMode.ForWrite) as DBText;
                            if (acText != null)
                            {
                                if (acText.TextString.Contains(""))
                                {
                                    string[] strs = acText.TextString.Split('');
                                    acText.TextString = acText.TextString.Replace("", "%%132");
                                }

                            }
                        }
                    }
                }
                acTrams.Commit();
            }
            sw.Stop();
            ed.WriteMessage("\n替换操作，耗时{0}ms",sw.ElapsedMilliseconds);
        }


    }
}
