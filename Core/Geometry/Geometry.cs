using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using CADAddin;

namespace Geometry
{
    public class Geometry
    {
        [CommandMethod("Ad")]
        public static void DrawLine()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database db = HostApplicationServices.WorkingDatabase;
            Stopwatch watch = new Stopwatch();
            watch.Start();
            for (int i = -5000; i < 5000; i += 1000)
            {
                Point3d startPoint = new Point3d(i, Math.Sqrt((Math.Pow(5000, 2) - Math.Pow(Math.Abs(i), 2))), 0);
                Point3d endPoint = new Point3d(-startPoint.X, -startPoint.Y, 0);
                Line line = new Line(startPoint, endPoint);
                //获得在此多段线上离起点距离为中点的点
                Point3d pts = line.GetPointAtDist(line.Length / 2);
                //得到这个点的切线方向
                Vector3d acVQ = line.GetFirstDerivative(pts);

                //旋转90度得到法线方向(顺)
                Vector3d acVF = acVQ.RotateBy(Math.PI / 2, Vector3d.ZAxis);

                //得到此向量上距离此点距离dust的点PT1
                Point3d txtPosition;
                double txtRotation;
                DBText text = new DBText();
                //得到此向量上距离此点距离dust的点PT1
                txtPosition = pts.Add(acVF.GetNormal() * 100);
                txtRotation = acVF.GetAngleTo(Vector3d.YAxis);

                text.Position = txtPosition;
                text.Height = 150;
                text.TextString = "测试文字";
                //text.Justify=AttachmentPoint.BottomLeft;
                text.Rotation = txtRotation;
                text.WidthFactor = 0.7;

                //text.HorizontalMode = TextHorizontalMode.TextLeft;
                line.ColorIndex = 1;
                db.AddToModelSapce(line);
                db.AddToModelSapce(text);
            }
            watch.Stop();
            ed.WriteMessage("\n命令执行完成：{0} ms", watch.ElapsedMilliseconds.ToString());
            watch.Reset();
        }
    }
}
