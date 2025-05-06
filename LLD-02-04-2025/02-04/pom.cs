namespace Vamshi;
class Vamshi{
    public static bool Data2(Entity UtilLines, Entity CenterLine){
        bool pos = false;
        using (Transaction tr = db.TransactionManager.StartTransaction())
        {
            Polyline plA = CenterLine as Polyline;
            Polyline plB = UtilLines as Polyline;
            Point3d midA = GetPolylineMidpoint(plA);
            Point3d midB = GetPolylineMidpoint(plB);
            if (midB.Y > midA.Y)
                return true;
            else if (midB.Y < midA.Y)
                return false;
        }
    }
}