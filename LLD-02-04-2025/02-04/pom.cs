public Point3dCollection PlaceRectangle(ObjectId refPolylineId, out ObjectId entId)
{
    Point3dCollection pts = new Point3dCollection();
    Polyline3d poly = new Polyline3d();
    Document doc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
    Database db = doc.Database;
    Editor ed = doc.Editor;
    ObjectId objId = ObjectId.Null;
    Matrix3d mat = ed.CurrentUserCoordinateSystem;

    // 1. Pick midpoint of width side
    PromptPointResult res = ed.GetPoint("\nPick center point of rectangle width side: ");
    if (res.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    Point3d basePt = res.Value.TransformBy(mat);

    // 2. Ask for width
    PromptDoubleOptions widthOpts = new PromptDoubleOptions("\nEnter rectangle width (across): ");
    widthOpts.AllowNegative = false;
    widthOpts.AllowZero = false;
    PromptDoubleResult widthRes = ed.GetDouble(widthOpts);
    if (widthRes.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    double width = widthRes.Value;

    // 3. Ask for length
    PromptDoubleOptions lengthOpts = new PromptDoubleOptions("\nEnter rectangle length (along direction): ");
    lengthOpts.AllowNegative = false;
    lengthOpts.AllowZero = false;
    PromptDoubleResult lengthRes = ed.GetDouble(lengthOpts);
    if (lengthRes.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    double length = lengthRes.Value;

    // 4. Get direction vector from polyline
    Vector3d unitDir = Vector3d.XAxis; // fallback
    using (Transaction tr = db.TransactionManager.StartTransaction())
    {
        Entity ent = tr.GetObject(refPolylineId, OpenMode.ForRead) as Entity;

        if (ent is Polyline pl && pl.NumberOfVertices > 1)
        {
            Point3d p10 = pl.GetPoint3dAt(0);
            Point3d p20 = pl.GetPoint3dAt(1);
            unitDir = (p20 - p10).GetNormal();
        }
        else if (ent is Polyline3d pl3d)
        {
            Point3d? firstPt = null, secondPt = null;
            int count = 0;

            foreach (ObjectId vId in pl3d)
            {
                Entity vEnt = tr.GetObject(vId, OpenMode.ForRead) as Entity;
                if (vEnt != null)
                {
                    var posProp = vEnt.GetType().GetProperty("Position");
                    if (posProp != null)
                    {
                        Point3d pt = (Point3d)posProp.GetValue(vEnt);
                        if (count == 0) firstPt = pt;
                        else if (count == 1) secondPt = pt;

                        count++;
                        if (count >= 2) break;
                    }
                }
            }

            if (firstPt.HasValue && secondPt.HasValue)
            {
                unitDir = (secondPt.Value - firstPt.Value).GetNormal();
            }
        }
        else
        {
            ed.WriteMessage("\nInvalid polyline reference.");
            entId = ObjectId.Null;
            return pts;
        }

        tr.Commit();
    }

    // 5. Compute perpendicular direction for width
    Vector3d perpDir = unitDir.CrossProduct(Vector3d.ZAxis).GetNormal();
    Vector3d halfWidthVec = perpDir * (width / 2);
    Vector3d lengthVec = unitDir * length;

    // 6. Compute rectangle corners clockwise from bottom-left
    Point3d p1 = basePt - halfWidthVec;        // bottom-left
    Point3d p2 = p1 + lengthVec;               // top-left
    Point3d p3 = p2 + (halfWidthVec * 2);      // top-right
    Point3d p4 = basePt + halfWidthVec;        // bottom-right

    pts.Add(p1);
    pts.Add(p2);
    pts.Add(p3);
    pts.Add(p4);

    // 7. Create rectangle polyline
    using (poly = new Polyline3d(Poly3dType.SimplePoly, pts, true))
    {
        using (BlockTableRecord btr = db.CurrentSpaceId.Open(OpenMode.ForWrite) as BlockTableRecord)
        {
            btr.AppendEntity(poly);
            db.TransactionManager.AddNewlyCreatedDBObject(poly, true);
            objId = poly.ObjectId;
        }
    }

    entId = objId;
    return pts;
}

public Point3dCollection PlaceRectangle(out ObjectId entId)
{
    Point3dCollection pts = new Point3dCollection();
    Polyline3d poly = new Polyline3d();
    Document doc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
    Database db = doc.Database;
    Editor ed = doc.Editor;
    ObjectId objId = ObjectId.Null;
    Matrix3d mat = ed.CurrentUserCoordinateSystem;

    // Step 1: Pick base point
    PromptPointResult res1 = ed.GetPoint("\nPick base point: ");
    if (res1.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    Point3d pickedPt1 = res1.Value.TransformBy(mat);
    Point3d basePt = GetNearestPointOnPolyline("BM_CENTERLINE", pickedPt1, db);

    // Step 2: Pick second point (for direction)
    PromptPointOptions ptOpts2 = new PromptPointOptions("\nPick another point for direction on centerline: ");
    ptOpts2.UseBasePoint = true;
    ptOpts2.BasePoint = basePt;
    PromptPointResult res2 = ed.GetPoint(ptOpts2);
    if (res2.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    Point3d pickedPt2 = res2.Value.TransformBy(mat);
    Point3d dirPt = GetNearestPointOnPolyline("BM_CENTERLINE", pickedPt2, db);

    // Step 3: Get width
    PromptDoubleOptions widthOpts = new PromptDoubleOptions("\nEnter width: ");
    widthOpts.AllowNegative = false;
    widthOpts.AllowZero = false;
    PromptDoubleResult widthRes = ed.GetDouble(widthOpts);
    if (widthRes.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    double width = widthRes.Value;

    // Step 4: Get length
    PromptDoubleOptions lenOpts = new PromptDoubleOptions("\nEnter length: ");
    lenOpts.AllowNegative = false;
    lenOpts.AllowZero = false;
    PromptDoubleResult lenRes = ed.GetDouble(lenOpts);
    if (lenRes.Status != PromptStatus.OK)
    {
        entId = ObjectId.Null;
        return pts;
    }
    double length = lenRes.Value;

    // Step 5: Compute rectangle orientation and corners
    Vector3d directionVec = (dirPt - basePt).GetNormal();  // length direction
    Vector3d widthVec = directionVec.CrossProduct(Vector3d.ZAxis).GetNormal(); // width direction (perpendicular)

    Vector3d halfLengthVec = directionVec * (length / 2.0);
    Vector3d halfWidthVec = widthVec * (width / 2.0);

    // Compute corners (clockwise)
    Point3d p1 = basePt - halfWidthVec - halfLengthVec;
    Point3d p2 = basePt - halfWidthVec + halfLengthVec;
    Point3d p3 = basePt + halfWidthVec + halfLengthVec;
    Point3d p4 = basePt + halfWidthVec - halfLengthVec;

    pts.Add(p1);
    pts.Add(p2);
    pts.Add(p3);
    pts.Add(p4);

    // Step 6: Create rectangle polyline

    using (poly = new Polyline3d(Poly3dType.SimplePoly, pts, true))
    {
        layer_methods.CrtTandChgLayer(doc.Editor, DictVals["GridLayer"]);
        poly.Layer = DictVals["GridLayer"];
        using (BlockTableRecord btr = db.CurrentSpaceId.Open(OpenMode.ForWrite) as BlockTableRecord)
        {
            btr.AppendEntity(poly);
            db.TransactionManager.AddNewlyCreatedDBObject(poly, true);
            objId = poly.ObjectId;
        }
    }

    entId = objId;
    return pts;
}