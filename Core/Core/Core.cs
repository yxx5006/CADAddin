using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Runtime;

namespace Core
{
    public class Core
    {
        [CommandMethod("count")]
        public static void GetAllObject()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (OpenCloseTransaction trans = acCurDb.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTable bt = trans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                int count = 0;
                acDoc.Editor.WriteMessage("\nModel space object:");

                foreach (var acObjId in btr)
                {
                    acDoc.Editor.WriteMessage("\n"+acObjId.ObjectClass.DxfName);
                    count += 1;
                }
                if (count == 0)
                {
                    acDoc.Editor.WriteMessage("\n0 Object Fount");
                }


            }
        }
    }
}
