using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using ToOffice;
namespace ToOffice
{
    public class test
    {
        private static Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
        private static Document acDoc = Application.DocumentManager.MdiActiveDocument;
        private static Database acCurDb = acDoc.Database;
        private static List<object> GetLayerObjectsByLayerName(List<string> layerNames)
        {
            List<object> objects = new List<object>();

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
        public static void wrob()
        {
            List<string> IMGLayers = Word.GetIMGLayers();
            List<object> objects = GetLayerObjectsByLayerName(IMGLayers);
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
                }
            }
        }
    }
}