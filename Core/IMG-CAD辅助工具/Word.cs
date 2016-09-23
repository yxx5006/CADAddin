using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using MSWord = Microsoft.Office.Interop.Word;

namespace IMG_CAD辅助工具
{
    public static class Word
    {
        static int i = 0;
        private static object _nothing = Missing.Value;
        public static void CreatDocFile(string Path)
        {
            //MSWord.Application WordApp;
            //MSWord.Document WordDoc;
            //object file=temp
            byte[] bytes = global::IMG_CAD辅助工具.Resource.template;
            File.WriteAllBytes(Path, bytes);
        }

        public static void OpenDocFile(out MSWord.Application WordApp, out MSWord.Document WordDoc, string FileName)
        {
            object fileName = FileName;
            WordApp = new MSWord.Application();
            //WordDoc = new MSWord.Document();
            WordDoc = WordApp.Documents.Open(ref fileName, ref _nothing, ref _nothing, ref _nothing, ref _nothing,
                ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing, ref _nothing,
                ref _nothing, ref _nothing, ref _nothing, ref _nothing);
            WordApp.Visible = true;
        }

        public static void CloseAndSaveDocFile(MSWord.Application WordApp, MSWord.Document WordDoc)
        {
            WordDoc.Save();
            WordDoc.Close(ref _nothing, ref _nothing, ref _nothing);
            WordApp.Quit(ref _nothing, ref _nothing, ref _nothing);
        }

        public static void ReplaceToWord(MSWord.Application WordApp, MSWord.Document WordDoc, Dictionary<string, string> DocContent)
        {

            MSWord.Range headeRange = WordDoc.Sections[3].Headers[MSWord.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
            headeRange.Find.Replacement.ClearFormatting();
            headeRange.Find.ClearFormatting();
            headeRange.Find.Text = "{SubProjectName}";
            headeRange.Find.Replacement.Text = DocContent["{SubProjectName}"];
            object objReplace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            headeRange.Find.Execute(ref _nothing, ref _nothing, ref _nothing,
                                       ref _nothing, ref _nothing, ref _nothing,
                                       ref _nothing, ref _nothing, ref _nothing,
                                       ref _nothing, ref objReplace, ref _nothing,
                                       ref _nothing, ref _nothing, ref _nothing);

            foreach (var item in DocContent)
            {
                Thread.Sleep(500);
                WordApp.Selection.Find.Replacement.ClearFormatting();
                WordApp.Selection.Find.ClearFormatting();
                WordApp.Selection.Find.Text = item.Key;
                if (item.Key=="{RecDate}")
                {
                    WordApp.Selection.Find.Replacement.Text = item.Value;
                }
                else
                {
                    WordApp.Selection.Find.Replacement.Text = item.Value+i;
                }

                WordApp.Selection.Find.Execute(ref _nothing, ref _nothing, ref _nothing,
                                           ref _nothing, ref _nothing, ref _nothing,
                                           ref _nothing, ref _nothing, ref _nothing,
                                           ref _nothing, ref objReplace, ref _nothing,
                                           ref _nothing, ref _nothing, ref _nothing);
                i++;
            }

            object endkeyunit = Microsoft.Office.Interop.Word.WdUnits.wdStory;
            WordApp.Selection.EndKey(ref endkeyunit, ref _nothing);
            //WordApp.Selection.MoveUp(ref unit, ref count, ref _nothing);
            object pBreak = (int)Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak;
            WordApp.Selection.InsertBreak(ref pBreak);
            WordApp.Selection.Paste();
            object dummy = System.Reflection.Missing.Value;

            object what = MSWord.WdGoToItem.wdGoToLine;

            object which = MSWord.WdGoToDirection.wdGoToPrevious;

            object count = 1;

            WordApp.Selection.GoTo(ref what, ref which, ref count, ref dummy);

        }

        public static void CopyTable(MSWord.Application WordApp, MSWord.Document WordDoc, int Number)
        {
                WordDoc.Tables[WordDoc.Tables.Count].Select();
                WordApp.Selection.Copy();
        }



    }
}
