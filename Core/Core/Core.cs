using System.Diagnostics;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace CADAddin
{
    public static class Core
    {
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

        public static ObjectId AddToModelSapce(this Database db, Entity ent)
        {
            ObjectId entId;
            using (OpenCloseTransaction trans=db.TransactionManager.StartOpenCloseTransaction())
            {
                BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr=trans.GetObject(bt[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;
                entId = btr.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent,true);
                trans.Commit();
            }
            return entId;
        }
    }
}
