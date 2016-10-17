

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace IMG_CAD辅助工具
{
    public static class Test
    {
        private static Document AcDoc;
        private static Database AcCurDb;
        private static Editor Ed;
        private static List<Line> rLines = new List<Line>();//获得表格竖直直线
        private static List<Line> cLines = new List<Line>();//获得表格水平直线
        private static List<List<string>> contents = new List<List<string>>();//获得表格内容


        [CommandMethod("GetSelect")]
        public static void GetSelect()
        {

            ReadXML();

            bool isSuccess = ReadImgTable();
            if (isSuccess == false) return;
            string textString;
            textString = Cells(1, 1);

            Ed.WriteMessage(textString);
            ReadAllContents();
            //// Ed.WriteMessage(contents.Count.ToString());
            ////for (int i = 0; i < contents.Count; i++)
            ////{
            ////    StringBuilder sb = new StringBuilder();
            ////    for (int j = 0; j < contents[i].Count; j++)
            ////    {
            ////        sb.Append(contents[i][j] + " ");
            ////        //Ed.WriteMessage("第{0}行：{1}",i, contents[i][j]);

            ////    }
            ////    Ed.WriteMessage("第{0}行：{1}\n", i, sb);
            ////    sb.Clear();
            ////}


            cLines.Clear();
            rLines.Clear();
            contents.Clear();
        }

        private static string GetText(string mtext)
        {
            string content = mtext; //多行文本内容
            //将多行文本按“\\”进行分割
            string[] strs = content.Split(new string[] { @"\\" }, StringSplitOptions.None);
            //指定不区分大小写
            RegexOptions ignoreCase = RegexOptions.IgnoreCase;
            for (int i = 0; i < strs.Length; i++)
            {
                //删除段落缩进格式
                strs[i] = Regex.Replace(strs[i], @"\\pi(.[^;]*);", "", ignoreCase);
                //删除制表符格式
                strs[i] = Regex.Replace(strs[i], @"\\pt(.[^;]*);", "", ignoreCase);
                //删除堆迭格式
                strs[i] = Regex.Replace(strs[i], @"\\S(.[^;]*)(\^|#|\\)(.[^;]*);", @"$1$3", ignoreCase);
                strs[i] = Regex.Replace(strs[i], @"\\S(.[^;]*)(\^|#|\\);", "$1", ignoreCase);
                //删除字体、颜色、字高、字距、倾斜、字宽、对齐格式
                strs[i] = Regex.Replace(strs[i], @"(\\F|\\C|\\H|\\T|\\Q|\\W|\\A)(.[^;]*);", "", ignoreCase);
                //删除下划线、删除线格式
                strs[i] = Regex.Replace(strs[i], @"(\\L|\\O|\\l|\\o)", "", ignoreCase);
                //删除不间断空格格式
                strs[i] = Regex.Replace(strs[i], @"\\~", "", ignoreCase);
                //删除换行符格式
                strs[i] = Regex.Replace(strs[i], @"\\P", "\n", ignoreCase);
                //删除换行符格式(针对Shift+Enter格式)
                //strs[i] = Regex.Replace(strs[i], "\n", "", ignoreCase);
                //删除{}
                strs[i] = Regex.Replace(strs[i], @"({|})", "", ignoreCase);
            }
            return string.Join("\\", strs); //将文本中的特殊字符去掉后重新连接成一个字符串
        }

        private static string Cells(int rows, int columns)
        {

            Point3dCollection downPoints = new Point3dCollection();
            Point3dCollection upPoints = new Point3dCollection();
            string textString = string.Empty;
            PromptSelectionResult psr;
            using (OpenCloseTransaction acTrans = AcCurDb.TransactionManager.StartOpenCloseTransaction())
            {
                if (rows == 1 && columns == 1)
                {
                    cLines[rows].IntersectWith(rLines[rLines.Count - 1], Intersect.ExtendBoth, new Plane(), downPoints,
                        IntPtr.Zero, IntPtr.Zero);
                    cLines[rows - 1].IntersectWith(rLines[1], Intersect.ExtendBoth, new Plane(), upPoints, IntPtr.Zero,
                        IntPtr.Zero);
                }
                else
                {
                    cLines[rows].IntersectWith(rLines[columns], Intersect.ExtendBoth, new Plane(), downPoints,
                        IntPtr.Zero, IntPtr.Zero);
                    cLines[rows - 1].IntersectWith(rLines[columns - 1], Intersect.ExtendBoth, new Plane(), upPoints,
                        IntPtr.Zero, IntPtr.Zero);
                }

                psr = Ed.SelectCrossingWindow(downPoints[0], upPoints[0]);
                if (psr.Status == PromptStatus.OK)
                {
                    SelectionSet set = psr.Value;
                    foreach (SelectedObject text in set)
                    {
                        if (text != null)
                        {
                            MText acMText = acTrans.GetObject(text.ObjectId, OpenMode.ForWrite) as MText;
                            if (acMText == null)
                            {
                                DBText acText = acTrans.GetObject(text.ObjectId, OpenMode.ForRead) as DBText;
                                if (acText != null)
                                {
                                    textString = acText.TextString;
                                    return textString;
                                }
                            }
                            else
                            {
                                textString = GetText(acMText.Contents);
                                return textString;
                            }
                        }
                    }
                }
                acTrans.Commit();
            }
            return "null";
        }

        private static bool ReadImgTable()
        {
            AcDoc = Application.DocumentManager.MdiActiveDocument;
            AcCurDb = AcDoc.Database;
            Ed = AcDoc.Editor;
            PromptSelectionResult psr = Ed.GetSelection();
            if (psr.Status == PromptStatus.OK)
            {
                SelectionSet ss = psr.Value;
                if (ss == null) return false;
                var ids = ss.GetObjectIds();
                using (OpenCloseTransaction acTrans = AcCurDb.TransactionManager.StartOpenCloseTransaction())
                {
                    foreach (var id in ids)
                    {
                        var dbLine = acTrans.GetObject(id, OpenMode.ForRead) as Line;
                        if (dbLine == null) continue;
                        Vector2d ve = new Vector2d(dbLine.StartPoint.X - dbLine.EndPoint.X,
                            dbLine.StartPoint.Y - dbLine.EndPoint.Y);
                        if (ve.Angle == 0 || ve.Angle == Math.PI)
                        {
                            cLines.Add(dbLine);
                        }
                        else if (ve.Angle == Math.PI / 2 || ve.Angle == 1.5 * Math.PI)
                        {
                            rLines.Add(dbLine);
                        }
                    }
                    acTrans.Commit();
                }
                var temp = from line in cLines
                           orderby line.EndPoint.Y descending
                           select line;
                cLines = temp.ToList();

                temp = from line in rLines
                       orderby line.EndPoint.X
                       select line;
                rLines = temp.ToList();
                return true;
            }
            return false;
        }

        private static void ReadAllContents()
        {
            List<string> content = new List<string>();
            for (int i = 1; i < rLines.Count; i++)
            {
                content.Add(Cells(2, i));
            }
            contents.Add(content);

            for (int i = 1; i < cLines.Count - 2; i++)
            {
                List<string> tempList = new List<string>();
                for (int j = 1; j < rLines.Count; j++)
                {
                    tempList.Add(Cells(i + 2, j));
                }
                contents.Add(tempList);
            }
        }

        private static void ReadXML()
        {
            AcDoc = Application.DocumentManager.MdiActiveDocument;
            AcCurDb = AcDoc.Database;
            Ed = AcDoc.Editor;
            XmlDocument doc = new XmlDocument();
            doc.Load("IMGToolSettings.xml");    //加载Xml文件  
            XmlElement rootElem = doc.DocumentElement;   //获取根节点  
            XmlNodeList layerNodes = rootElem.GetElementsByTagName("Layer"); //获取Layer子节点集合  
            foreach (XmlNode node in layerNodes)
            {
                string strLayerName = ((XmlElement)node).GetAttribute("LayerName");   //获取LayerName属性值  
                //Console.WriteLine(strName);
                Ed.WriteMessage(strLayerName);
                XmlNodeList subNameNodes = ((XmlElement)node).GetElementsByTagName("Name");  //获取Name子XmlElement集合  
                if (subNameNodes.Count == 1)
                {
                    string strName = subNameNodes[0].InnerText;
                    //Console.WriteLine(s);
                    Ed.WriteMessage(strName);
                }
            }
        }
    }
}

