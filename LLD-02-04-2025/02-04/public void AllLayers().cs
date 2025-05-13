public void AllLayers()
        {
            try
            {
                ObjectId entId = new ObjectId();
                List<Entity> entlstRec = new List<Entity>();
                Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                iniPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\NubuildPermits\\NUB.ini";
                DictVals = ini_methods.GetIniKeyFieldNvalues("DimensionPlacement", iniPath);
                Entity ent = null;
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    Point3d pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                    while (pos1 != new Point3d())
                    {
                        pos1 = new Point3d();

                        using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                        {
                            double dblPnt = 0.0; Polyline plineEOP2 = new Polyline(); Polyline plineEOP = new Polyline();
                            Point3d EOPPnt2 = new Point3d(); Point3d EOPPnt = new Point3d();
                            Point3dCollection pntColl = PlaceRectangle(out entId);
                            pntColl.Add(pntColl[0]); bool isleft = false;
                            ent = acTr.GetObject(entId, OpenMode.ForRead) as Entity;
                            entlstRec.Add(ent);
                            if (pntColl != null && pntColl.Count > 0)
                            {
                                if (rbFenceAllLayers.Checked == true)
                                {
                                    SelectionSet ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*", DictVals["ReqLayers"] );// "Fn_FOC,Fn-FOC"); 
                                    Point3dCollection pntCol = new Point3dCollection(); bool isleft2 = false;
                                    Point3d ROWPnt = new Point3d(); Point3d ROWPnt2 = new Point3d(); Polyline plineROW = new Polyline(); Polyline plineROW2 = new Polyline();
                                    Point3d UEPnt = new Point3d(); Point3d UEPnt1 = new Point3d(); Polyline plineUE = new Polyline(); Polyline plineUE2 = new Polyline(); Polyline plineEOP4 = new Polyline();
                                    if (ssSrcLay != null && ssSrcLay.Count > 0)
                                    {
                                        IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                        Dictionary<ObjectId, string> Values = GetLayers(lstEnts);
                                        Entity entEOP1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"Fn_FOC"*/)).Select(x => x).FirstOrDefault();
                                        if (entEOP1 != null && ent != null)
                                        {
                                            ent.IntersectWith(entEOP1, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                            Point3d CrsPnt = pntCol.Count == 2 ? General_methods.GetDistanceBetweenPoints(pntCol[0], pntColl[0]) < General_methods.GetDistanceBetweenPoints(pntCol[1], pntColl[0]) ? pntCol[0] : pntCol[1] : pntCol[0];
                                            EOPPnt = CrsPnt;// pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                            double Ang = General_methods.GetAnglePntBetween3dPoints(pntCol[0], pntCol[1]);

                                            if ((Ang > 45 && Ang < 110) || (Ang > 210 && Ang < 300))
                                            {
                                                isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[1].Y > /*Math.Abs(*/pntCol[0].Y/*)*/ ? false : true : false;

                                            }
                                            else
                                            {
                                                isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[1].X > /*Math.Abs(*/pntCol[0].X/*)*/ ? false : true : false;
                                            }
                                            int eopLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("NB_BM_EOP"));
                                            if (eopLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more times
                                            {
                                                plineEOP = entEOP1 as Polyline;
                                                dblPnt = plineEOP.GetDistAtPoint(EOPPnt);
                                                dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                Entity entEOP2 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_BM_EOP"/*"FN_FOC"*/)).Select(x => x).LastOrDefault();
                                                plineEOP2 = entEOP2 as Polyline;
                                                EOPPnt2 = plineEOP2.GetClosestPointTo(EOPPnt, true);
                                                double dist = General_methods.GetDistanceBetweenPoints(EOPPnt, EOPPnt2);
                                                layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, EOPPnt, EOPPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                            }
                                            int rowLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("UG Utility - Gas"));
                                            if (rowLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more times
                                            {
                                                Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entROW1 is Polyline)
                                                {
                                                    plineROW = entROW1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    ROWPnt = plineROW.GetClosestPointTo(EOPPnt, true);
                                                    Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("UG Utility - Gas"/*"FN_ROW"*/)).Select(x => x).LastOrDefault();
                                                    plineROW2 = entROW2 as Polyline;
                                                    ROWPnt2 = plineROW2.GetClosestPointTo(ROWPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, ROWPnt2);  
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, ROWPnt, ROWPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            rowLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("UG Utility - Hydro"));
                                            if (rowLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more times
                                            {
                                                Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entROW1 is Polyline)
                                                {
                                                    plineROW = entROW1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    ROWPnt = plineROW.GetClosestPointTo(EOPPnt, true);
                                                    Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("UG Utility - Hydro"/*"FN_ROW"*/)).Select(x => x).LastOrDefault();
                                                    plineROW2 = entROW2 as Polyline;
                                                    ROWPnt2 = plineROW2.GetClosestPointTo(ROWPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, ROWPnt2);
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, ROWPnt, ROWPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            rowLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("UG Utility - Sanitary"));
                                            if (rowLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more times
                                            {
                                                Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entROW1 is Polyline)
                                                {
                                                    plineROW = entROW1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    ROWPnt = plineROW.GetClosestPointTo(EOPPnt, true);
                                                    Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("UG Utility - Sanitary"/*"FN_ROW"*/)).Select(x => x).LastOrDefault();
                                                    plineROW2 = entROW2 as Polyline;
                                                    ROWPnt2 = plineROW2.GetClosestPointTo(ROWPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, ROWPnt2);
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, ROWPnt, ROWPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            rowLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("UG Utility - Bell"));
                                            if (rowLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more PointF direction = new PointF((float)(ROWPnt2.X - ROWPnt.X), ROWPnt2.Y - ROWPnt.Y);
                                            {
                                                Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entROW1 is Polyline)
                                                {
                                                    plineROW = entROW1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    ROWPnt = plineROW.GetClosestPointTo(EOPPnt, true);
                                                    Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("UG Utility - Bell"/*"FN_ROW"*/)).Select(x => x).LastOrDefault();
                                                    plineROW2 = entROW2 as Polyline;
                                                    ROWPnt2 = plineROW2.GetClosestPointTo(ROWPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, ROWPnt2);
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, ROWPnt, ROWPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            rowLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("UG Utility - Storm"));
                                            if (rowLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more PointF direction = new PointF((float)(ROWPnt2.X - ROWPnt.X), ROWPnt2.Y - ROWPnt.Y);
                                            {
                                                Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entROW1 is Polyline)
                                                {
                                                    plineROW = entROW1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    ROWPnt = plineROW.GetClosestPointTo(EOPPnt, true);
                                                    Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("UG Utility - Storm"/*"FN_ROW"*/)).Select(x => x).LastOrDefault();
                                                    plineROW2 = entROW2 as Polyline;
                                                    ROWPnt2 = plineROW2.GetClosestPointTo(ROWPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, ROWPnt2);
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, ROWPnt, ROWPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            rowLayerCount = lstEnts.Count(x => x.Layer.ToString().Contains("UG Utility - Watermain"));
                                            if (rowLayerCount >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more PointF direction = new PointF((float)(ROWPnt2.X - ROWPnt.X), ROWPnt2.Y - ROWPnt.Y);
                                            {
                                                Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entROW1 is Polyline)
                                                {
                                                    plineROW = entROW1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    ROWPnt = plineROW.GetClosestPointTo(EOPPnt, true);
                                                    Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("UG Utility - Watermain"/*"FN_ROW"*/)).Select(x => x).LastOrDefault();
                                                    plineROW2 = entROW2 as Polyline;
                                                    ROWPnt2 = plineROW2.GetClosestPointTo(ROWPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, ROWPnt2);
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, ROWPnt, ROWPnt2, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            int ueLayerCount1 = lstEnts.Count(x => x.Layer.ToString().Contains("NB_BM_PROPERTY LINE"));

                                            if (ueLayerCount1 >= 1) // Only proceed if "UTILITY EASEMENT" appears two or more times
                                            {
                                                Entity entUE1 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_MICRODUCT"/*"Fn_FOC"*/)).Select(x => x).FirstOrDefault();
                                                if (entUE1 is Polyline)
                                                {
                                                    plineUE = entUE1 as Polyline;
                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                                    UEPnt = plineUE.GetClosestPointTo(EOPPnt, true);
                                                    Entity entUE2 = lstEnts.Where(x => x.Layer.ToString().Contains("NB_BM_PROPERTY LINE"/*"FN_FOC"*/)).Select(x => x).LastOrDefault();
                                                    plineUE2 = entUE2 as Polyline;
                                                    UEPnt1 = plineUE2.GetClosestPointTo(UEPnt, true);
                                                    double dist = General_methods.GetDistanceBetweenPoints(UEPnt, UEPnt1);
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, dist, UEPnt, UEPnt1, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }


                                            // Step 2: Process layers left of Fn-Center Line
                                            //ProcessFilteredUtilityLayers(acDoc, acTr, pntColl, EOPPnt, plineEOP);

                                        }
                                        else
                                        {
                                            // Check if the list contains "UTILITY EASEMENT" layer at least twice
                                            int ueLayerCount2 = lstEnts.Count(x => x.Layer.ToString().Contains("UTILITY EASEMENT"));

                                            if (ueLayerCount2 >= 2) // Only proceed if "UTILITY EASEMENT" appears two or more times
                                            {
                                                // Find the first entity with "UTILITY EASEMENT" layer
                                                Entity entUE1 = lstEnts.Where(x => x.Layer.ToString().Contains("UTILITY EASEMENT")).Select(x => x).FirstOrDefault();
                                                if (entUE1 is Polyline)
                                                {
                                                    plineUE = entUE1 as Polyline;
                                                    UEPnt = plineUE.GetClosestPointTo(pntColl[0], true);

                                                    // Find the last entity with "UTILITY EASEMENT" layer
                                                    Entity entUE2 = lstEnts.Where(x => x.Layer.ToString().Contains("UTILITY EASEMENT")).Select(x => x).LastOrDefault();
                                                    plineUE2 = entUE2 as Polyline;
                                                    UEPnt1 = plineUE2.GetClosestPointTo(UEPnt, true);

                                                    // Create and change to the desired layer and place dimensions
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs" /*DictVals["DimensionLayer"]*/);
                                                    ObjectId TxtId = PlaceDimension(acDoc, acTr, 2.0, UEPnt, UEPnt1, "DIMs"/*DictVals["DimensionLayer"]*/);
                                                }
                                            }
                                            //Step 2: Process layers left of Fn-Center Line
                                            //ProcessFilteredUtilityLayers(acDoc, acTr, pntColl, EOPPnt, plineEOP);

                                            List<Polyline> selectedRows = new List<Polyline>();
                                            List<Polyline> boreLayers1 = new List<Polyline>();

                                            foreach (Entity entity in lstEnts)
                                            {
                                                if (entity is Polyline pline)
                                                {
                                                    if (pline.Layer == "UTILITY EASEMENT" || pline.Layer == "Parcel" || pline.Layer == "Fn-Water" || pline.Layer == "Fn-Sanitary_Sewer" || pline.Layer == "Fn-Storm Drain" || pline.Layer == "ONG")
                                                    {
                                                        selectedRows.Add(pline);
                                                    }
                                                    if (pline.Layer == "Fn-P-DB" || pline.Layer == "Fn-TRENCH")

                                                    {
                                                        boreLayers1.Add(pline);
                                                    }
                                                }
                                            }

                                            // Proceed to work with all selected row layers
                                            if (selectedRows.Count > 0)
                                            {
                                                int i = 0;
                                                //bool isleft2 = false;
                                                double dblPnt2 = 0.0;
                                                foreach (Polyline currentBore in boreLayers1)
                                                {
                                                    // Loop through all selected row layers
                                                    foreach (Polyline selectedRow in selectedRows)
                                                    {

                                                        double offset = i * 8.0;
                                                        i++;
                                                        currentBore.IntersectWith(selectedRow, Intersect.OnBothOperands, pntColl, IntPtr.Zero, IntPtr.Zero);

                                                        if (pntColl.Count == 0)
                                                            pntColl.Add(General_methods.GetMidPntBetweenPoints(currentBore.StartPoint, currentBore.EndPoint));

                                                        double fenceAngle = General_methods.GetAnglePntBetween3dPoints(pntColl[0], pntColl[1]);
                                                        bool isLeft = false;

                                                        if (pntColl.Count >= 2)
                                                        {
                                                            if ((fenceAngle > 45 && fenceAngle < 110) || (fenceAngle > 210 && fenceAngle < 300))
                                                                isLeft = pntColl[1].Y > pntColl[0].Y ? false : true;
                                                            else
                                                                isLeft = pntColl[1].X > pntColl[0].X ? false : true;
                                                        }
                                                        currentBore.IntersectWith(selectedRow, Intersect.OnBothOperands, pntColl, IntPtr.Zero, IntPtr.Zero);
                                                        Point3d CrsPnt = pntColl.Count == 2 ? General_methods.GetDistanceBetweenPoints(pntColl[0], pntColl[0]) < General_methods.GetDistanceBetweenPoints(pntColl[1], pntColl[0]) ? pntColl[0] : pntColl[1] : pntColl[0];

                                                        double Ang = General_methods.GetAnglePntBetween3dPoints(pntColl[0], pntColl[1]);

                                                        if ((Ang > 45 && Ang < 110) || (Ang > 210 && Ang < 300))
                                                        {
                                                            isleft2 = pntColl.Count > 1 ? /*Math.Abs(*/pntColl[1].Y > /*Math.Abs(*/pntColl[0].Y/*)*/ ? false : true : false;

                                                        }
                                                        else
                                                        {
                                                            isleft2 = pntColl.Count > 1 ? /*Math.Abs(*/pntColl[1].X > /*Math.Abs(*/pntColl[0].X/*)*/ ? false : true : false;
                                                        }

                                                        // Adjust dimension placement based on user-defined fence rectangle with offset
                                                        Point3d BorePnt = currentBore.GetClosestPointTo(pntColl[0], false);
                                                        dblPnt2 = currentBore.GetDistAtPoint(BorePnt);
                                                        dblPnt2 = isleft2 == false ? dblPnt2 - 8.0 : dblPnt2 + 8.0;
                                                        BorePnt = currentBore.GetPointAtDist(dblPnt2);

                                                        Point3d rowPnt1 = selectedRow.GetClosestPointTo(BorePnt, true);

                                                        double distAtRowPnt1 = selectedRow.GetDistAtPoint(rowPnt1);

                                                        // Adjust Bore Point with offset
                                                        double adjustedDist = isLeft ? distAtRowPnt1 - offset : distAtRowPnt1 + offset;
                                                        Point3d rowPnt2 = selectedRow.GetPointAtDist(adjustedDist);

                                                        Point3d borePnt = currentBore.GetClosestPointTo(rowPnt2, true);
                                                        Point3d rowPnt = selectedRow.GetClosestPointTo(borePnt, true);

                                                        double finalDist = Math.Round(General_methods.GetDistanceBetweenPoints(borePnt, rowPnt), 2);

                                                        layer_methods.CrtTandChgLayer(acDoc.Editor, "DIMs");
                                                        ObjectId TxtId = PlaceDimension_Updated(acDoc, acTr, finalDist, borePnt, rowPnt, "DIMs");

                                                    }
                                                }
                                            }

                                        }
                                        acTr.Commit();
                                    }
                                    acDoc.Editor.Regen();
                                    MessageBox.Show("Please provide a new point to continue or ESC to exit");
                                    pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                                    if (pos1 == new Point3d(0,0,0))
                                    {
                                        MessageBox.Show("Dimensions placed successfully");
                                    }
                                }
                            }

                        }
                    }
                    foreach (Entity item in entlstRec)
                    {
                        General_methods.delEntity(item);
                    }
                }
            }

            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }