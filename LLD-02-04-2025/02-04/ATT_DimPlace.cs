using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections.Generic;
using Teigha.Geometry;
using IntelliCAD.EditorInput;
//using Teigha.Runtime;
using Teigha.DatabaseServices;
using IntelliCAD.ApplicationServices;
using System.Collections;
using System.Collections.Specialized;
//using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Xml;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using Teigha.Colors;
using System.Reflection;

namespace AT_TPermits
{
    public partial class DimPlace_AT_T_Permits : Form
    {
        Dictionary<string, string> DictVals = new Dictionary<string, string>(); string iniPath = string.Empty;
        public DimPlace_AT_T_Permits()
        {
            InitializeComponent();
        }

        private void btnProceed_Click(object sender, EventArgs e)
        {
            try
            {
                iniPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\AT_Permits.ini";
                DictVals = ini_methods.GetIniKeyFieldNvalues("Dimension Placement", iniPath);
                DateTime firstDate = DateTime.Today;
                DateTime secondDate = new DateTime(2025, 12, 30);
                int dResult = DateTime.Compare(firstDate, secondDate);
                //if (dResult != 1)
                {
                    var socketName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName.ToUpper();
                    if (socketName != null && (socketName.Contains("TECHWAVE") || socketName.Contains("IN.TECHWAVE.NET")))
                    {
                        if (rbSingle.Checked == true)
                        {
                            SingleSelection();
                        }
                        else if (rbPLUE.Checked == true)
                        {
                            PlacePLUEDim();
                        }
                        else
                        {
                            FenceZiply();
                        }
                    }
                    else
                    {
                        MessageBox.Show("invalid attempt, access denied", "Dimension Placement", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Close();
                    }
                }
                //else
                //{
                //    MessageBox.Show("tool Expired, :(", "Dimension Placement", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //    this.Close();
                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PlacePLUEDim()
        {
            int cnt = 0;
            Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            try
            {
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                    {
                        
                        SelectionSet sSet = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*LINE", "PL NEW");
                        if (sSet!=null && sSet.Count>0)
                        {
                            foreach (ObjectId PL in sSet.GetObjectIds())
                            {
                                cnt++;
                                Polyline PlNew = acTr.GetObject(PL, OpenMode.ForRead) as Polyline;
                                Point3d MidPnt = General_methods.GetMidPointsForEntity(PlNew);
                                Point3dCollection BuffColl = General_methods.funGetBuffPts(MidPnt, 30);
                                SelectionSet sSetCrs = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, BuffColl, "*LINE", "Easement");
                                if (sSetCrs!=null && sSetCrs.Count>0)
                                {
                                    Polyline pline1 = acTr.GetObject(sSetCrs.GetObjectIds()[0], OpenMode.ForRead) as Polyline;
                                    Point3d Pnt1 = pline1.GetClosestPointTo(MidPnt,true);
                                    double dist = General_methods.GetDistanceBetweenPoints(MidPnt, Pnt1);
                                    layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);
                                    ObjectId TxtId = PlaceDimensionPLUE(acDoc, acTr, dist, MidPnt, Pnt1, DictVals["DimensionLayer"]);
                                    AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                    XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);

                                    double Ang = General_methods.GetAnglePntBetween3dPoints(MidPnt, Pnt1);
                                    Point3d TxtPnt = General_methods.PolarPoint(Pnt1, General_methods.DTR(Ang), 5.0);
                                    dist = General_methods.GetDistanceBetweenPoints(MidPnt, Pnt1);
                                    string Txtinfo = Math.Round(dist,1) + "'" + " EASEMENT";
                                    Ang = General_methods.textReadbleAng(Ang + 90);
                                    PlaceFeature_methods.CreateColouredMText2(acDoc.Editor, TxtPnt, acDoc, "HouseHold No", Txtinfo, 2.1320, General_methods.DTR(Ang), AttachmentPoint.TopLeft, 252);

                                    if (sSetCrs.Count>1)
                                    {
                                        double dblPnt = PlNew.GetDistAtPoint(MidPnt);//SRIDEVI
                                        dblPnt =dblPnt<5.5?dblPnt+5.0: dblPnt - 5.0;
                                        if (PlNew.Length<dblPnt)
                                        {
                                            dblPnt = 1.0;
                                        }
                                        MidPnt = PlNew.GetPointAtDist(dblPnt);
                                        Polyline pline2 = acTr.GetObject(sSetCrs.GetObjectIds()[1], OpenMode.ForRead) as Polyline;
                                        Point3d Pnt2 = pline2.GetClosestPointTo(MidPnt, true);
                                        dist = General_methods.GetDistanceBetweenPoints(MidPnt, Pnt2);
                                        TxtId = PlaceDimensionPLUE(acDoc, acTr, dist, MidPnt, Pnt2, DictVals["DimensionLayer"]);
                                        TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                        XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);

                                        Ang = General_methods.GetAnglePntBetween3dPoints(MidPnt, Pnt2);
                                        TxtPnt = General_methods.PolarPoint(Pnt2, General_methods.DTR(Ang), 5.0);
                                        dist = General_methods.GetDistanceBetweenPoints(MidPnt, Pnt2);
                                        Txtinfo = Math.Round(dist,1) + "'" + " EASEMENT";
                                        Ang = General_methods.textReadbleAng(Ang + 90);
                                        PlaceFeature_methods.CreateColouredMText2(acDoc.Editor, TxtPnt, acDoc, "HouseHold No", Txtinfo, 2.1320, General_methods.DTR(Ang), AttachmentPoint.TopLeft, 252);
                                    }
                                    

                                }
                            }
                        }
                        acTr.Commit();
                        MessageBox.Show("Process Completed");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message+cnt);
            }
        }

        private void SingleSelection()
        {
            Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            try
            {
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    //SelectionSet sSetPline = selectionset_methods.getLinearFeatFrmAcad(acDoc.Editor, "*LINE*");
                    // Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Point3d pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                    while (pos1 != null)
                    {
                        using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                        {
                            Point3d pos2 = acDoc.Editor.GetPoint("\n Pick Next Point").Value;
                            Point3dCollection pntColl = General_methods.funGetBuffPts(pos1, 1.0);
                            SelectionSet sSet = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", "*");
                            Entity Entline = acTr.GetObject(sSet.GetObjectIds()[0], OpenMode.ForRead) as Entity;

                            Point3d pntLine = new Point3d(); Point3d pntLine2 = new Point3d();
                            if (Entline is Teigha.DatabaseServices.Line)
                            {
                                Teigha.DatabaseServices.Line pline = Entline as Teigha.DatabaseServices.Line;
                                pntLine = pline.GetClosestPointTo(pos1, true);
                            }
                            else if (Entline is Polyline)
                            {
                                Polyline pline = Entline as Polyline;
                                pntLine = pline.GetClosestPointTo(pos1, true);
                            }
                            pntColl = General_methods.funGetBuffPts(pos2, 1);
                            sSet = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE*", "*");
                            Entity Entline2 = acTr.GetObject(sSet.GetObjectIds()[0], OpenMode.ForRead) as Entity;
                            if (Entline2 is Teigha.DatabaseServices.Line)
                            {
                                Teigha.DatabaseServices.Line pline2 = acTr.GetObject(sSet.GetObjectIds()[0], OpenMode.ForRead) as Teigha.DatabaseServices.Line;
                                pntLine2 = pline2.GetClosestPointTo(pntLine, true);
                            }
                            else if (Entline2 is Polyline)
                            {
                                Polyline pline2 = acTr.GetObject(sSet.GetObjectIds()[0], OpenMode.ForRead) as Polyline;
                                pntLine2 = pline2.GetClosestPointTo(pntLine, true);
                            }

                            double distROW = General_methods.GetDistanceBetweenPoints(pntLine, pntLine2);
                            //PlaceFeature_methods.placeDimension(pntLine, pntLine2, distROW.ToString());//(acDoc, acTr, distROW, pntLine, pntLine2);
                            layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"] /*DictVals["DimensionLayer"]*/);
                            ObjectId TxtId = PlaceDimension(acDoc, acTr, distROW, pntLine, pntLine2, DictVals["DimensionLayer"]);
                            //pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                            AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                            XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", DictVals["DimensionLayer"]);
                            PromptPointResult ppr = acDoc.Editor.GetPoint("\nPick First Position: ");
                            if (ppr.Status != PromptStatus.OK) { return; }
                            else { pos1 = ppr.Value; }
                            acTr.Commit();
                            acDoc.Editor.Regen();
                        }

                    }
                }

                //Point3d pntLine = pline.GetClosestPointTo(pos1, true);

                //pntColl = General_methods.funGetBuffPts(pos2, 5);
                //sSet = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, pntColl, "*", "*LINE*");
                //Polyline pline2 = acTr.GetObject(sSet.GetObjectIds()[0], OpenMode.ForRead) as Polyline;
                //Point3d pntLine2 = pline2.GetClosestPointTo(pntLine, true);

                //double distROW = General_methods.GetDistanceBetweenPoints(pntLine, pntLine2);
                //ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distROW, pntLine, pntLine2);


                #region ROW BOC
                //SelectionSet sSet = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*LINE*", "VPORT 1 60");
                //if (sSet != null && sSet.Count > 0)
                //{
                //    foreach (ObjectId objId in sSet.GetObjectIds())
                //    {
                //        Dictlines = new Dictionary<Point3d, Entity>();
                //        Entity ent = acTr.GetObject(objId, OpenMode.ForRead) as Entity;
                //        if (ent is Polyline)
                //        {
                //            Polyline pline = ent as Polyline;
                //            Point3dCollection pntColl1 = General_methods.GetCoordinates3d(pline);
                //            Point3dCollection gridPts3D = new Point3dCollection();
                //            SelectionSet sSetpoly = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, pntColl1, "PROP-MT", "*LINE*");
                //            //BOC - GPS,
                //            if (sSetpoly == null) { continue; }
                //            else
                //            {
                //                foreach (ObjectId item in sSetpoly.GetObjectIds())
                //                {
                //                    gridPts3D = new Point3dCollection();
                //                    Polyline plineTemp = acTr.GetObject(item, OpenMode.ForRead) as Polyline;
                //                    pline.IntersectWith(plineTemp, Intersect.OnBothOperands, gridPts3D, IntPtr.Zero, IntPtr.Zero);
                //                    if (plineTemp.Layer == "PROP-MT")
                //                    {
                //                        if (gridPts3D.Count == 1)
                //                        {
                //                            double dist = plineTemp.GetDistAtPoint(gridPts3D[0]);
                //                            Point3d PropPnt = plineTemp.GetPointAtDist(dist / 2);
                //                            Point3d PropROWPnt = plineTemp.GetPointAtDist((dist / 2) + 2);
                //                            Point3dCollection buffcoll = General_methods.funGetBuffPts(PropPnt, 10.0);
                //                            Zoom_methods.ZoomToEntity(acDoc.Editor, plineTemp);
                //                            SelectionSet ssCbles = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, buffcoll, "BOC-GPS,ROW", "*LINE*");
                //                            foreach (ObjectId cblEntId in ssCbles.GetObjectIds())
                //                            {
                //                                Polyline cblEnt = acTr.GetObject(cblEntId, OpenMode.ForRead) as Polyline;
                //                                if (cblEnt.Layer == "BOC-GPS")
                //                                {
                //                                    Point3d BOCPnt = cblEnt.GetClosestPointTo(PropPnt, true);
                //                                    double distboc = General_methods.GetDistanceBetweenPoints(PropPnt, BOCPnt);
                //                                    ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distboc, PropPnt, BOCPnt);
                //                                }
                //                                else if (cblEnt.Layer == "ROW")
                //                                {
                //                                    Point3d ROWPnt = cblEnt.GetClosestPointTo(PropROWPnt, true);
                //                                    double distROW = General_methods.GetDistanceBetweenPoints(PropROWPnt, ROWPnt);
                //                                    ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distROW, PropROWPnt, ROWPnt);
                //                                }
                //                            }
                //                            //double dbl = plineTemp.GetParameterAtPoint(gridPts3D[0]);
                //                            //Point3d pnt1 = plineTemp.GetPointAtParameter(Math.Ceiling(dbl));
                //                            ////ispointinside poly if yes continue elsse
                //                            ////Point3d pnt2 = plineTemp.GetPointAtParameter(Math.Floor(dbl));
                //                            //Dictlines[pnt1] = plineTemp;

                //                            //double angle = General_methods.GetAnglePntBetween3dPoints(pnt1, pnt2);
                //                            //angle = angle + 90;
                //                            //angle = General_methods.textReadbleAng(angle);
                //                            //angle = General_methods.DTR(angle);
                //                            //Dictlines[gridPts3D[0]] = plineTemp;

                //                        }
                //                        else if (gridPts3D.Count > 1)
                //                        {
                //                            double dist1 = plineTemp.GetDistAtPoint(gridPts3D[0]);
                //                            double dist2 = plineTemp.GetDistAtPoint(gridPts3D[1]);
                //                            double FinalDist = Math.Abs(dist1 - dist2);
                //                            Point3d Temppnt = General_methods.GetMidPntBetweenPoints3d(gridPts3D[0], gridPts3D[1]);

                //                            Point3d Pnt1 = General_methods.GetMidPntBetweenPoints3d(gridPts3D[0], Temppnt);
                //                            Pnt1 = plineTemp.GetClosestPointTo(Pnt1, true);
                //                            double dist = plineTemp.GetDistAtPoint(Pnt1);

                //                            Point3d PropROWPnt = plineTemp.GetPointAtDist(dist + 2);
                //                            Point3dCollection buffcoll = General_methods.funGetBuffPts(Pnt1, 10.0);
                //                            Zoom_methods.ZoomToEntity(acDoc.Editor, plineTemp);
                //                            SelectionSet ssCbles = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, buffcoll, "BOC-GPS,ROW", "*LINE*");
                //                            foreach (ObjectId cblEntId in ssCbles.GetObjectIds())
                //                            {
                //                                Polyline cblEnt = acTr.GetObject(cblEntId, OpenMode.ForRead) as Polyline;
                //                                if (cblEnt.Layer == "BOC-GPS")
                //                                {
                //                                    Point3d BOCPnt = cblEnt.GetClosestPointTo(Pnt1, true);
                //                                    double distboc = General_methods.GetDistanceBetweenPoints(Pnt1, BOCPnt);
                //                                    ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distboc, Pnt1, BOCPnt);
                //                                    //place dimension
                //                                }
                //                                else if (cblEnt.Layer == "ROW")
                //                                {
                //                                    Point3d ROWPnt = cblEnt.GetClosestPointTo(PropROWPnt, true);
                //                                    double distROW = General_methods.GetDistanceBetweenPoints(PropROWPnt, ROWPnt);
                //                                    ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distROW, PropROWPnt, ROWPnt);
                //                                    //Place Dimension
                //                                }
                //                            }

                //                            Point3d Pnt2 = General_methods.GetMidPntBetweenPoints3d(gridPts3D[1], Temppnt);
                //                            Pnt2 = plineTemp.GetClosestPointTo(Pnt2, true);
                //                            dist = plineTemp.GetDistAtPoint(Pnt2);
                //                            buffcoll = General_methods.funGetBuffPts(Pnt2, 10.0);
                //                            Zoom_methods.ZoomToEntity(acDoc.Editor, plineTemp);
                //                            ssCbles = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, buffcoll, "BOC-GPS,ROW", "*LINE*");
                //                            if (ssCbles.Count > 0)
                //                            {
                //                                foreach (ObjectId cblEntId in ssCbles.GetObjectIds())
                //                                {
                //                                    Polyline cblEnt = acTr.GetObject(cblEntId, OpenMode.ForRead) as Polyline;
                //                                    if (cblEnt.Layer == "BOC-GPS")
                //                                    {
                //                                        Point3d BOCPnt = cblEnt.GetClosestPointTo(Pnt2, true);
                //                                        double distboc = General_methods.GetDistanceBetweenPoints(Pnt2, BOCPnt);
                //                                        ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distboc, Pnt2, BOCPnt);
                //                                        //place dimension
                //                                    }
                //                                    else if (cblEnt.Layer == "ROW")
                //                                    {
                //                                        Point3d ROWPnt = cblEnt.GetClosestPointTo(PropROWPnt, true);
                //                                        double distROW = General_methods.GetDistanceBetweenPoints(PropROWPnt, ROWPnt);
                //                                        ObjectId DimId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, distROW, PropROWPnt, ROWPnt);
                //                                        //Place Dimension
                //                                    }
                //                                }
                //                            }

                //                            //Point3d PropPnt1 = plineTemp.GetPointAtDist(dist1 / 2);
                //                            //Point3d PropROWPnt = plineTemp.GetPointAtDist((dist1 / 2) + 2);
                //                            //Point3dCollection buffcoll = General_methods.funGetBuffPts(PropPnt1, 10.0);
                //                            //Zoom_methods.ZoomToEntity(acDoc.Editor, plineTemp);
                //                            //SelectionSet ssCbles = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, buffcoll, "BOC-GPS,ROW", "*LINE*");

                //                            //AlignedDimension dim = new AlignedDimension();

                //                            //double dbl = plineTemp.GetParameterAtPoint(gridPts3D[0]);

                //                            //Point3d pnt1 = plineTemp.GetPointAtParameter(Math.Ceiling(dbl));
                //                            //ispointinside poly if yes continue elsse
                //                            //Point3d pnt2 = plineTemp.GetPointAtParameter(Math.Floor(dbl));
                //                            //Dictlines[pnt1] = plineTemp;

                //                            //double angle = General_methods.GetAnglePntBetween3dPoints(pnt1, pnt2);
                //                            //angle = angle + 90;
                //                            //angle = General_methods.textReadbleAng(angle);
                //                            //angle = General_methods.DTR(angle);
                //                            //Dictlines[gridPts3D[0]] = plineTemp;

                //                        }
                //                    }
                //                }

                //            }
                //        }
                //    }
                //}

                #endregion




                // }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void Fence()
        {
            try
            {
                Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Entity ent = null;
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    // Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Point3d pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                    while (pos1 != new Point3d())
                    {
                        pos1 = new Point3d();
                        using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                        {
                            Point3dCollection pntColl = General_methods.PlacePolylineFence(out ent);
                            if (pntColl != null && pntColl.Count > 0)
                            {
                                //ObjectId objId = Gen_Polyline_Entity(pntColl, "0", 7);
                                SelectionSet ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLayer(acDoc.Editor, pntColl, " E-Right of Way,E-Edge of Pavement,E-Water Lines,P-Conduit,E-Storm Drains,E-Features,E-Sewer", "*LINE*");
                                Point3dCollection pntCol = new Point3dCollection();
                                if (ssSrcLay != null)
                                {
                                    IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                    //Entity entConduit = from Entity entMain in lstEnts where entMain.Layer.ToLower().Contains("conduit") select entMain;
                                    Entity entConduit = lstEnts.Where(x => x.Layer.ToLower().ToString().Contains("conduit")).Select(x => x).FirstOrDefault();
                                    if (entConduit != null)
                                    {
                                        ent.IntersectWith(entConduit, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                        Point3d pntConduit = pntCol[0];
                                        Polyline plineCond = entConduit as Polyline;
                                        double dblPnt = plineCond.GetDistAtPoint(pntConduit);
                                        //Point3d nearpnt = pline.GetClosestPointTo(linePnt, true);
                                        double dbl = plineCond.GetParameterAtPoint(pntConduit);
                                        Point3d pnt1 = plineCond.GetPointAtParameter(Math.Ceiling(dbl));
                                        Point3d pnt2 = plineCond.GetPointAtParameter(Math.Floor(dbl));
                                        foreach (Entity entPly in lstEnts)
                                        {
                                            if (entPly is Polyline && entPly != entConduit && entPly != ent)
                                            {
                                                Polyline pline = entPly as Polyline;
                                                string plinelay = pline.Layer;
                                                Point3d Pnt = pline.GetClosestPointTo(pntConduit, true);
                                                //double dist = General_methods.GetDistanceBetweenPoints(Pnt, pntConduit);
                                                double dist = 2.0;
                                                ObjectId TxtId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, dist, Pnt, pntConduit);
                                                //double dblPnt = pline.GetDistAtPoint(pntConduit);
                                                AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                dblPnt = dblPnt - 5.0;
                                                pntConduit = plineCond.GetPointAtDist(dblPnt - 5.0);
                                            }
                                            else if (entPly != entConduit && entPly is Teigha.DatabaseServices.Line)
                                            {
                                                Teigha.DatabaseServices.Line pline = entPly as Teigha.DatabaseServices.Line;
                                                string plinelay = pline.Layer;
                                                Point3d Pnt = pline.GetClosestPointTo(pntConduit, true);
                                                //double dist = General_methods.GetDistanceBetweenPoints(Pnt, pntConduit);
                                                double dist = 2;
                                                ObjectId TxtId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, dist, Pnt, pntConduit);
                                                AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                dblPnt = dblPnt - 5.0;
                                                pntConduit = plineCond.GetPointAtDist(dblPnt - 5.0);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Entity entEOP1 = lstEnts.Where(x => x.Layer.ToString().Contains("E-Edge of Pavement")).Select(x => x).FirstOrDefault();
                                        ent.IntersectWith(entEOP1, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                        Point3d EOPPnt = pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                        Polyline plineEOP = entEOP1 as Polyline;
                                        Entity entEOP2 = lstEnts.Where(x => x.Layer.ToString().Contains("E-Edge of Pavement")).Select(x => x).LastOrDefault();
                                        Polyline plineEOP2 = entEOP2 as Polyline;
                                        Point3d EOPPnt2 = plineEOP2.GetClosestPointTo(EOPPnt, true);
                                        ObjectId TxtId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, 2.0, EOPPnt, EOPPnt2);
                                        AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                        XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                        Point3d pnt = plineEOP.GetClosestPointTo(EOPPnt, true);
                                        double dblPnt = plineEOP.GetDistAtPoint(EOPPnt);
                                        //Point3d nearpnt = pline.GetClosestPointTo(linePnt, true);
                                        dblPnt = dblPnt - 5.0;
                                        EOPPnt = plineEOP.GetPointAtDist(dblPnt - 10.0);
                                        Point3d ROWPnt1 = new Point3d();
                                        Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains("E-Right of Way")).Select(x => x).FirstOrDefault();
                                        if (entROW1 is Polyline)
                                        {
                                            Polyline PlineROW = entROW1 as Polyline; ROWPnt1 = PlineROW.GetClosestPointTo(EOPPnt, true);
                                        }
                                        else
                                        {
                                            Teigha.DatabaseServices.Line PlineROW = entROW1 as Teigha.DatabaseServices.Line; ROWPnt1 = PlineROW.GetClosestPointTo(EOPPnt, true);
                                        }

                                        Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains("E-Right of Way")).Select(x => x).LastOrDefault();
                                        Point3d ROWPnt2 = new Point3d();
                                        if (entROW2 is Polyline)
                                        {
                                            Polyline PlineROW2 = entROW2 as Polyline; ROWPnt2 = PlineROW2.GetClosestPointTo(ROWPnt1, true);
                                        }
                                        else
                                        {
                                            Teigha.DatabaseServices.Line PlineROW2 = entROW2 as Teigha.DatabaseServices.Line; ROWPnt2 = PlineROW2.GetClosestPointTo(ROWPnt1, true);
                                        }
                                        TxtId = PlaceFeature_methods.PlaceDimension(acDoc, acTr, 2.0, ROWPnt1, ROWPnt2);
                                        AlignedDimension TxtEnt1 = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                        XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt1, "LayerName", TxtEnt1.Layer);
                                    }
                                    //PlaceFeature_methods.EraseEnts(acDoc, ent);
                                }
                                else
                                {
                                    MessageBox.Show("Given layer Names are not found in Drawing file.", "ConduitBlockPlace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                            //else
                            //{
                            //    MessageBox.Show("Given layer Names are not found in Drawing file.", "ConduitBlockPlace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //}
                            acTr.Commit();
                            pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FenceZiply()
        {
            try
            {

                ObjectId entId = new ObjectId(); List<Entity> entlstRec = new List<Entity>();
                Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Entity ent = null;
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    //Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    Point3d pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                    while (pos1 != new Point3d())
                    {
                        pos1 = new Point3d(); Point3d ROWPnt1 = new Point3d();
                        using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                        {
                            double dblPnt = 0.0; Polyline plineEOP2 = new Polyline(); Polyline plineEOP = new Polyline();
                            Point3d EOPPnt2 = new Point3d(); Point3d EOPPnt = new Point3d();
                            //Point3dCollection pntColl = General_methods.PlacePolylineFence(out ent);
                            Point3dCollection pntColl = PlaceRectangle(out entId);
                            pntColl.Add(pntColl[0]); bool isleft = false;
                            ent = acTr.GetObject(entId, OpenMode.ForRead) as Entity;
                            entlstRec.Add(ent);
                            if (pntColl != null && pntColl.Count > 0)
                            {
                                if (rbFenceRWEOP.Checked == true)
                                {
                                    SelectionSet ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["EOPLayer"] + "," + DictVals["CentreLineLayer"]);// "FN_FOC,Fn-Center Line");
                                    Point3dCollection pntCol = new Point3dCollection(); Point3d Pnt = new Point3d(); Polyline plineCond = new Polyline(); Point3d pntConduit = new Point3d();
                                    if (ssSrcLay != null && ssSrcLay.Count > 0)
                                    {
                                        IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                        Entity entPropMT = lstEnts.Where(x => x.Layer.ToLower().ToString().Contains(DictVals["CentreLineLayer"]/*"fn-center line"*/)).Select(x => x).FirstOrDefault();
                                        if (entPropMT != null)
                                        {
                                            //ent.IntersectWith(entPropMT, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                            //pntConduit = pntCol[0];
                                            IList<Entity> lstEnts1 = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                            Entity entEOP1 = lstEnts1.Where(x => x.Layer.ToString().Contains(DictVals["EOPLayer"]/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                            ent.IntersectWith(entEOP1, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                            Point3d CrsPnt =pntCol.Count==2? General_methods.GetDistanceBetweenPoints(pntCol[0], pntColl[0]) < General_methods.GetDistanceBetweenPoints(pntCol[1], pntColl[0]) ? pntCol[0] : pntCol[1]:pntCol[0];
                                            EOPPnt = CrsPnt;// pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                            double Ang = General_methods.GetAnglePntBetween3dPoints(pntCol[0], pntCol[1]);
                                            if ((Ang>45 && Ang<110) || (Ang > 210 && Ang < 300))
                                            {
                                                isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[1].Y > /*Math.Abs(*/pntCol[0].Y/*)*/ ? false : true : false;
                                                
                                            }
                                            else
                                            {
                                                isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[1].X > /*Math.Abs(*/pntCol[0].X/*)*/ ? false : true : false;
                                            }
                                            
                                            plineEOP = entEOP1 as Polyline;
                                            plineCond = entPropMT as Polyline;
                                            pntConduit = plineCond.GetClosestPointTo(EOPPnt, true);

                                            //pntConduit = plineCond.GetPointAtDist(dblPnt);
                                            dblPnt = plineCond.GetDistAtPoint(pntConduit);//SRIDEVI
                                            dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                            pntConduit = plineCond.GetPointAtDist(dblPnt);
                                            ////Point3d pnts = plineEOP2.GetClosestPointTo(pntConduit, true);
                                            ////pntConduit = EOPPnt2 == pnts ? plineCond.GetPointAtDist(dblPnt - 16) : pntConduit;
                                            ////dblPnt = EOPPnt2 == pnts ? dblPnt - 16 : dblPnt;

                                            //dblPnt =plineCond.GetDistAtPoint(pntConduit);
                                            //double dbl = plineCond.GetParameterAtPoint(pntConduit);
                                            int i = 1;
                                            foreach (Entity entPly in lstEnts)
                                            {
                                                if (entPly is Polyline && entPly != entPropMT && entPly != ent)
                                                {
                                                    Polyline pline = entPly as Polyline;
                                                    string plinelay = pline.Layer;
                                                    dblPnt = i != 1 ? isleft == false ? dblPnt + 8.0 : dblPnt - 8 : dblPnt; i++;
                                                    pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                    Pnt = pline.GetClosestPointTo(pntConduit, true);
                                                    //double dist = General_methods.GetDistanceBetweenPoints(Pnt, pntConduit);
                                                    double dist = 2.0;
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);

                                                    ObjectId TxtId = PlaceDimension(acDoc, acTr, dist, Pnt, pntConduit, DictVals["DimensionLayer"]);
                                                    AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                    XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                    //double dblPnt = pline.GetDistAtPoint(pntConduit);
                                                    //dblPnt = i!=1?dblPnt - 8.0:dblPnt;i++;
                                                    //pntConduit = plineCond.GetPointAtDist(dblPnt);// - 5.0);
                                                }
                                                else if (entPly != entPropMT && entPly is Line)
                                                {
                                                    Line pline = entPly as Line;
                                                    string plinelay = pline.Layer;
                                                    dblPnt = i != 1 ? isleft == false ? dblPnt + 8.0 : dblPnt - 8 : dblPnt; i++;
                                                    pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                    Pnt = pline.GetClosestPointTo(pntConduit, true);
                                                    //double dist = General_methods.GetDistanceBetweenPoints(Pnt, pntConduit);
                                                    double dist = 2.0;
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);
                                                    ObjectId TxtId = PlaceDimension(acDoc, acTr, dist, Pnt, pntConduit, DictVals["DimensionLayer"]);
                                                    AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                    XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);

                                                    //dblPnt = i != 1 ? dblPnt - 8.0 : dblPnt; i++;
                                                    pntConduit = plineCond.GetPointAtDist(dblPnt);// - 5.0);
                                                }
                                                else if (entPly is BlockReference)
                                                {
                                                    BlockReference blkEnt = entPly as BlockReference;
                                                    Point3d BlkPnt = blkEnt.Position;
                                                    Pnt = plineCond.GetClosestPointTo(BlkPnt, true);
                                                    double dist = 2;
                                                    layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);
                                                    ObjectId TxtId = PlaceDimension(acDoc, acTr, dist, BlkPnt, Pnt, DictVals["DimensionLayer"]);
                                                    AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                    XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                    dblPnt = dblPnt - 5.0;
                                                    pntConduit = plineCond.GetPointAtDist(dblPnt - 5.0);
                                                }
                                            }
                                        }

                                    }
                                    ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["EOPLayer"]);// "Q - ROAD_R-W,FN_FOC");
                                    pntCol = new Point3dCollection();
                                    if (ssSrcLay != null && ssSrcLay.Count > 0)
                                    {
                                        IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                        Entity entEOP1 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["EOPLayer"]/*"FN_FOC"*/)).Select(x => x).FirstOrDefault();
                                        //ent.IntersectWith(entEOP1, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                        //EOPPnt = pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                        dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                        pntConduit = plineCond.GetPointAtDist(dblPnt);
                                        plineEOP = entEOP1 as Polyline;
                                        EOPPnt = plineEOP.GetClosestPointTo(pntConduit, true);

                                        Entity entEOP2 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["EOPLayer"]/*"FN_FOC"*/)).Select(x => x).LastOrDefault();
                                        plineEOP2 = entEOP2 as Polyline;
                                        EOPPnt2 = plineEOP2.GetClosestPointTo(EOPPnt, true);
                                        layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"] /*DictVals["DimensionLayer"]*/);
                                        ObjectId TxtId = PlaceDimension(acDoc, acTr, 2.0, EOPPnt, EOPPnt2, DictVals["DimensionLayer"]/*DictVals["DimensionLayer"]*/);
                                        AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                        XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                        Point3d pnt = plineEOP.GetClosestPointTo(EOPPnt, true);
                                        //dblPnt = plineEOP.GetDistAtPoint(pnt);
                                        //dblPnt = dblPnt - 8.0;
                                    }


                                    ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["ROWLayer"]);// "Q - ROAD_R-W");

                                    if (ssSrcLay != null && ssSrcLay.Count > 1)
                                    {
                                        IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                        //EOPPnt = plineEOP.GetPointAtDist(dblPnt-8);// - 8.0);
                                        ROWPnt1 = new Point3d();
                                        Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["ROWLayer"])).Select(x => x).FirstOrDefault();
                                        if (entROW1 is Polyline)
                                        {
                                            Polyline PlineROW = entROW1 as Polyline;

                                            //EOPPnt = plineEOP.GetPointAtDist(dblPnt);
                                            pntConduit = plineCond.GetClosestPointTo(EOPPnt, true);
                                            dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                            pntConduit = plineCond.GetPointAtDist(dblPnt);
                                            ROWPnt1 = PlineROW.GetClosestPointTo(pntConduit, true);
                                            dblPnt = PlineROW.GetDistAtPoint(ROWPnt1);//SRIDEVI
                                                                                      //dblPnt = dblPnt - 8.0;
                                                                                      //ROWPnt1 = PlineROW.GetPointAtDist(dblPnt);
                                            Point3d pnts = PlineROW.GetClosestPointTo(EOPPnt, true);
                                            ROWPnt1 = ROWPnt1 == pnts ? PlineROW.GetPointAtDist(dblPnt + 8.0) : ROWPnt1;
                                            dblPnt = ROWPnt1 == pnts ? dblPnt = isleft == false ? dblPnt - 8.0 : dblPnt + 8 : dblPnt;

                                        }
                                        else
                                        {
                                            Line PlineROW = entROW1 as Line; ROWPnt1 = PlineROW.GetClosestPointTo(Pnt, true);
                                            dblPnt = PlineROW.GetDistAtPoint(ROWPnt1);//SRIDEVI
                                            dblPnt = dblPnt - 16.0;
                                            ROWPnt1 = PlineROW.GetPointAtDist(dblPnt);
                                            ROWPnt1 = PlineROW.GetPointAtDist(dblPnt);
                                            Point3d pnts = PlineROW.GetClosestPointTo(EOPPnt, true);
                                            ROWPnt1 = ROWPnt1 == pnts ? PlineROW.GetPointAtDist(dblPnt - 8) : ROWPnt1;
                                            dblPnt = ROWPnt1 == pnts ? dblPnt - 8 : dblPnt;
                                        }
                                        Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["ROWLayer"])).Select(x => x).LastOrDefault();
                                        Point3d ROWPnt2 = new Point3d();
                                        if (entROW2 is Polyline)
                                        {
                                            Polyline PlineROW2 = entROW2 as Polyline; ROWPnt2 = PlineROW2.GetClosestPointTo(ROWPnt1, true);
                                        }
                                        else
                                        {
                                            Line PlineROW2 = entROW2 as Line; ROWPnt2 = PlineROW2.GetClosestPointTo(ROWPnt1, true);
                                        }
                                        ObjectId TxtId = PlaceDimension(acDoc, acTr, 2.0, ROWPnt1, ROWPnt2, DictVals["DimensionLayer"]);
                                        AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                        XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);


                                    }
                                    #region R/W TO U/E
                                    ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["UELay"]);
                                    if (ssSrcLay != null && ssSrcLay.Count > 0)
                                    {
                                        ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["ROWLayer"]);
                                        if (ssSrcLay != null && ssSrcLay.Count > 0)
                                        {
                                            ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["UELay"] + "," + DictVals["ROWLayer"]);
                                            IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                            Entity legend = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["UELay"])).Select(x => x).FirstOrDefault();
                                            SelectionSet ssSrcLayRW = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["ROWLayer"]);
                                            lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                            Entity entROW1 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["ROWLayer"])).Select(x => x).FirstOrDefault();
                                            ent.IntersectWith(entROW1, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                            Point3d ROWPnt = pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                            Polyline plineROW = entROW1 as Polyline;
                                            Polyline plineLegend = legend as Polyline;
                                            //Point3d pntLegend = plineLegend.GetClosestPointTo(ROWPnt, true);
                                            //ObjectId TxtId = PlaceDimension(acDoc, acTr, 2.0, ROWPnt, pntLegend, DictVals["DimensionLayer"]);

                                            //pntConduit = plineCond.GetClosestPointTo(EOPPnt, true);
                                            dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                            //pntConduit = plineCond.GetPointAtDist(dblPnt);
                                            ROWPnt = plineROW.GetPointAtDist(dblPnt);
                                            //ROWPnt = plineROW.GetClosestPointTo(pntConduit, true);
                                            Point3d pntLegend = plineLegend.GetClosestPointTo(ROWPnt, true);
                                            double dist1 = General_methods.GetDistanceBetweenPoints(pntLegend, ROWPnt);
                                            if (dist1>20)
                                            {
                                                Entity entROW22 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["ROWLayer"])).Select(x => x).LastOrDefault();
                                                Polyline plineROW22 = entROW22 as Polyline;
                                                ROWPnt = plineROW22.GetClosestPointTo(pntLegend, true);
                                            }
                                            ObjectId TxtId = PlaceDimension(acDoc, acTr, 2.0, ROWPnt, pntLegend, DictVals["DimensionLayer"]);
                                            AlignedDimension alignEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                            XDATA_methods.ADDXdataNew(acDoc.Editor, alignEnt, "LayerName", alignEnt.Layer);
                                            double Ang = General_methods.GetAnglePntBetween3dPoints(ROWPnt, pntLegend);
                                            Point3d TxtPnt = General_methods.PolarPoint(pntLegend, General_methods.DTR(Ang), 5.0);
                                            double dist = General_methods.GetDistanceBetweenPoints(ROWPnt, pntLegend);
                                            string Txtinfo = Math.Round(dist) + "'" + " EASEMENT";
                                            Ang = General_methods.textReadbleAng(Ang + 90);
                                            PlaceFeature_methods.CreateColouredMText2(acDoc.Editor, TxtPnt, acDoc, "HouseHold No", Txtinfo, 2.1320, General_methods.DTR(Ang), AttachmentPoint.TopLeft,252);

                                            Entity entROW2 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["ROWLayer"])).Select(x => x).LastOrDefault();
                                            pntCol = new Point3dCollection();
                                            ent.IntersectWith(entROW2, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                            ROWPnt = pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                            Polyline plineROW2 = entROW2 as Polyline;
                                            Entity legend2 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["UELay"])).Select(x => x).LastOrDefault();
                                            Polyline plineLegend2 = legend2 as Polyline;
                                            //Point3d pntLegend = plineLegend.GetClosestPointTo(ROWPnt, true);
                                            //ObjectId TxtId = PlaceDimension(acDoc, acTr, 2.0, ROWPnt, pntLegend, DictVals["DimensionLayer"]);

                                            //pntConduit = plineCond.GetClosestPointTo(EOPPnt, true);
                                            //dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                            //pntConduit = plineCond.GetPointAtDist(dblPnt);
                                            //ROWPnt = plineROW.GetPointAtDist(dblPnt);
                                            ROWPnt = plineROW2.GetClosestPointTo(ROWPnt, true);
                                            pntLegend = plineLegend2.GetClosestPointTo(ROWPnt, true);
                                            double dist2 = General_methods.GetDistanceBetweenPoints(pntLegend, ROWPnt);
                                            if (dist1 > 20)
                                            {
                                                Entity entROW22 = lstEnts.Where(x => x.Layer.ToString().Contains(DictVals["ROWLayer"])).Select(x => x).FirstOrDefault();
                                                Polyline plineROW22 = entROW22 as Polyline;
                                                ROWPnt = plineROW22.GetClosestPointTo(pntLegend, true);
                                            }
                                            TxtId = PlaceDimension(acDoc, acTr, 2.0, ROWPnt, pntLegend, DictVals["DimensionLayer"]);
                                            alignEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                            XDATA_methods.ADDXdataNew(acDoc.Editor, alignEnt, "LayerName", alignEnt.Layer);

                                            Ang = General_methods.GetAnglePntBetween3dPoints(ROWPnt, pntLegend);
                                            TxtPnt = General_methods.PolarPoint(pntLegend, General_methods.DTR(Ang), 5.0);
                                            dist = General_methods.GetDistanceBetweenPoints(ROWPnt, pntLegend);
                                            Txtinfo = Math.Round(dist) + "'" + " EASEMENT";
                                            Ang = General_methods.textReadbleAng(Ang + 90);
                                            PlaceFeature_methods.CreateColouredMText2(acDoc.Editor, TxtPnt, acDoc, "HouseHold No", Txtinfo, 2.1320, General_methods.DTR(Ang), AttachmentPoint.TopLeft,252);

                                        }
                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region Bore-PL new
                                    Point3dCollection pntCol = new Point3dCollection(); Point3d Pnt = new Point3d(); Polyline plineCond = new Polyline(); Point3d pntConduit = new Point3d();
                                    SelectionSet ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["BoreLayer"]);
                                    if (ssSrcLay != null && ssSrcLay.Count > 0)
                                    {
                                        ssSrcLay = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntColl, "*LINE", DictVals["BoreLayer"] + "," + DictVals["UtilityLayers"]);// "BORE - 1 DUCT,PL NEW,ELECTRIC LINE,Fn-Sanitary_Sewer,LEGEND-1");//PL NEW,ELECTRIC LINE,LEGEND-1,
                                        if (ssSrcLay != null && ssSrcLay.Count > 0)
                                        {
                                            IList<Entity> lstEnts = General_methods.GetlstFromsSet(acDoc.Editor, ssSrcLay);
                                            Entity entPropMT = lstEnts.Where(x => x.Layer.ToUpper().ToString().Contains(DictVals["BoreLayer"]/*"BORE - 1 DUCT"*/)).Select(x => x).FirstOrDefault();
                                            if (entPropMT != null)
                                            {
                                                //ent.IntersectWith(entPropMT, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                                //pntConduit = pntCol[0];

                                                //Entity entBorePnt = lstEnts.Where(x => x.Layer.ToString().Contains("FN_FOC")).Select(x => x).FirstOrDefault();
                                                ent.IntersectWith(entPropMT, Intersect.OnBothOperands, pntCol, IntPtr.Zero, IntPtr.Zero);
                                                Point3d BorePnt = pntCol.Count == 2 ? General_methods.GetDistanceBetweenPoints(pntCol[0], pntColl[0]) < General_methods.GetDistanceBetweenPoints(pntCol[1], pntColl[0]) ? pntCol[0] : pntCol[1] : pntCol[0];
                                                //Point3d BorePnt = pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                                //plineEOP = entEOP1 as Polyline;
                                                //Point3d CrsPnt = pntCol.Count == 2 ? General_methods.GetDistanceBetweenPoints(pntCol[0], pntColl[0]) < General_methods.GetDistanceBetweenPoints(pntCol[1], pntColl[0]) ? pntCol[0] : pntCol[1] : pntCol[0];
                                                //EOPPnt = CrsPnt;// pntCol.Count > 0 ? pntCol.Count == 1 ? pntCol[0] : pntCol.Count == 2 ? General_methods.GetMidPntBetweenPoints3d(pntCol[0], pntCol[1]) : new Point3d() : new Point3d();
                                                //isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[0].X > /*Math.Abs(*/pntCol[1].X/*)*/ ? false : true : false;
                                                double Ang = General_methods.GetAnglePntBetween3dPoints(pntCol[0], pntCol[1]);
                                                if ((Ang > 45 && Ang < 110) || (Ang > 210 && Ang < 300))
                                                {
                                                    isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[1].Y > /*Math.Abs(*/pntCol[0].Y/*)*/ ? false : true : false;

                                                }
                                                else
                                                {
                                                    isleft = pntCol.Count > 1 ? /*Math.Abs(*/pntCol[1].X > /*Math.Abs(*/pntCol[0].X/*)*/ ? false : true : false;
                                                }

                                                plineCond = entPropMT as Polyline;
                                                pntConduit = plineCond.GetClosestPointTo(BorePnt, true);
                                                
                                                //pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                try
                                                {
                                                    dblPnt = plineCond.GetDistAtPoint(pntConduit);//SRIDEVI
                                                                                                  //dblPnt = dblPnt - 4.0;

                                                    dblPnt = isleft == false ? dblPnt + 8.0 : dblPnt - 8;
                                                    pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                    //pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                    ////Point3d pnts = plineEOP2.GetClosestPointTo(pntConduit, true);
                                                    ////pntConduit = EOPPnt2 == pnts ? plineCond.GetPointAtDist(dblPnt - 16) : pntConduit;
                                                    ////dblPnt = EOPPnt2 == pnts ? dblPnt - 16 : dblPnt;

                                                    //dblPnt =plineCond.GetDistAtPoint(pntConduit);
                                                    //double dbl = plineCond.GetParameterAtPoint(pntConduit);
                                                    int i = 1;
                                                    foreach (Entity entPly in lstEnts)
                                                    {
                                                        if (entPly is Polyline && entPly != entPropMT && entPly != ent)
                                                        {
                                                            Polyline pline = entPly as Polyline;
                                                            string plinelay = pline.Layer;
                                                            dblPnt = i != 1 ? isleft == false ? dblPnt + 6.0 : dblPnt - 6.0 : dblPnt; i++;
                                                            //dblPnt = i != 1 ? dblPnt - 6.0 : dblPnt; i++;
                                                            try
                                                            {
                                                                pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                                Pnt = pline.GetClosestPointTo(pntConduit, true);
                                                                double dist = General_methods.GetDistanceBetweenPoints(Pnt, pntConduit);
                                                                dist = Math.Round(dist);// 2.0;
                                                                layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);

                                                                ObjectId TxtId = PlaceDimension(acDoc, acTr, dist, Pnt, pntConduit, DictVals["DimensionLayer"]);
                                                                AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                                XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                                //double dblPnt = pline.GetDistAtPoint(pntConduit);
                                                                //dblPnt = i!=1?dblPnt - 8.0:dblPnt;i++;
                                                                pntConduit = plineCond.GetPointAtDist(dblPnt);// - 5.0);
                                                            }
                                                            catch (Exception)
                                                            {
                                                            }

                                                        }
                                                        else if (entPly != entPropMT && entPly is Line)
                                                        {
                                                            Line pline = entPly as Line;
                                                            string plinelay = pline.Layer;
                                                            dblPnt = i != 1 ? isleft == false ? dblPnt + 6.0 : dblPnt - 6.0 : dblPnt; i++;
                                                            //dblPnt = i != 1 ? dblPnt - 8.0 : dblPnt; i++;
                                                            try
                                                            {
                                                                pntConduit = plineCond.GetPointAtDist(dblPnt);
                                                                Pnt = pline.GetClosestPointTo(pntConduit, true);
                                                                double dist = General_methods.GetDistanceBetweenPoints(Pnt, pntConduit);
                                                                dist = Math.Round(dist);// 2.0;
                                                                layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);
                                                                ObjectId TxtId = PlaceDimension(acDoc, acTr, dist, Pnt, pntConduit, DictVals["DimensionLayer"]);
                                                                AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                                XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                                //dblPnt = i != 1 ? dblPnt - 8.0 : dblPnt; i++;
                                                                pntConduit = plineCond.GetPointAtDist(dblPnt);// - 5.0);
                                                            }
                                                            catch (Exception)
                                                            {
                                                            }

                                                        }
                                                        else if (entPly is BlockReference)
                                                        {
                                                            BlockReference blkEnt = entPly as BlockReference;
                                                            Point3d BlkPnt = blkEnt.Position;
                                                            Pnt = plineCond.GetClosestPointTo(BlkPnt, true);
                                                            double dist = 2;
                                                            layer_methods.CrtTandChgLayer(acDoc.Editor, DictVals["DimensionLayer"]);
                                                            ObjectId TxtId = PlaceDimension(acDoc, acTr, dist, BlkPnt, Pnt, DictVals["DimensionLayer"]);
                                                            AlignedDimension TxtEnt = acTr.GetObject(TxtId, OpenMode.ForWrite) as AlignedDimension;
                                                            XDATA_methods.ADDXdataNew(acDoc.Editor, TxtEnt, "LayerName", TxtEnt.Layer);
                                                            dblPnt = dblPnt - 5.0;
                                                            pntConduit = plineCond.GetPointAtDist(dblPnt - 5.0);
                                                        }
                                                    }
                                                }
                                                catch (Exception)
                                                {
                                                }
                                                
                                            }
                                        }
                                    }

                                    #endregion
                                }
                            }
                            else
                            {
                                MessageBox.Show("Given layer Names are not found in Drawing file.", "ConduitBlockPlace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                            //else
                            //{
                            //    MessageBox.Show("Given layer Names are not found in Drawing file.", "ConduitBlockPlace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            //}
                            acTr.Commit();
                            acDoc.Editor.Regen();
                            pos1 = acDoc.Editor.GetPoint("\nPick First Position:").Value;
                        }
                        acDoc.Editor.Regen();
                    }
                    if (entlstRec.Count > 0)
                    {
                        foreach (Entity item in entlstRec)
                        {
                            General_methods.delEntity(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //public addXdataValue()
        //{
        //    List<TypedValue> xdataList = new List<TypedValue>();
        //    xdataList.Add(new TypedValue((int)DxfCode.ExtendedDataRegAppName, odRecord.TableName));
        //    List<string> lst = ODUtils.GetODFieldNames(acDoc, odTable.Name);
        //    FieldDefinitions fldDEfs = odTable.FieldDefinitions;
        //    for (int i = 0; i < odTable.FieldDefinitions.Count; i++)
        //    {
        //        int fldIndx = fldDEfs.GetColumnIndex(odTable.FieldDefinitions[i].Name);
        //        xdataList.Add(new TypedValue((int)DxfCode.ExtendedDataAsciiString, odTable.FieldDefinitions[i].Name + "=" + odRecord[fldIndx].StrValue));

        //        //xdataList.Add(new TypedValue((int)DxfCode.ExtendedDataAsciiString, odRecord[fldIndx].StrValue));
        //    }
        //    XDATA_methods.AddRegAppTableRecord(acDoc.Editor, odRecord.TableName);
        //    ResultBuffer rb = new ResultBuffer(xdataList.ToArray());
        //    ent.XData = rb;
        //}
        public Point3dCollection PlaceRectangle(out ObjectId entId) // This method can have any name
        {
            Point3dCollection pts = new Point3dCollection(); Polyline3d poly = new Polyline3d();
            Document doc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor; ObjectId objId = new ObjectId();
            Matrix3d mat = ed.CurrentUserCoordinateSystem;
            PromptPointResult res = ed.GetPoint("\n Pick a Position: ");
            if (res.Status == PromptStatus.OK)
            {
                Point3d basePt = res.Value.TransformBy(mat);
                DrawJigWithDynDim jig = new DrawJigWithDynDim(basePt, 25, 25);
                if (jig.DragMe() == PromptStatus.OK)
                {
                    pts = new Point3dCollection();
                    pts.Add(basePt);
                    pts.Add(basePt + new Vector3d(0, jig.height, 0));
                    pts.Add(basePt + new Vector3d(jig.width, jig.height, 0));
                    pts.Add(basePt + new Vector3d(jig.width, 0, 0));
                    using (poly = new Polyline3d(Poly3dType.SimplePoly, pts, true))
                    {
                        using (BlockTableRecord btr = db.CurrentSpaceId.Open(OpenMode.ForWrite) as BlockTableRecord)
                        {
                            btr.AppendEntity(poly);
                            objId = poly.ObjectId;
                            //ent = poly;
                        }
                    }
                }
            }
            entId = objId;
            return pts;

        }
        internal static ObjectId PlaceDimensionPLUE(Document acDoc, Transaction acTr, double Dist, Point3d Pnt1, Point3d Pnt2, string Layr)
        {
            ObjectId id = new ObjectId();
            try
            {
                using (Transaction acTrans = acDoc.Database.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    // Create the aligned dimension
                    AlignedDimension acAliDim = new AlignedDimension();
                    acAliDim.SetDatabaseDefaults();
                    acAliDim.XLine1Point = Pnt1;
                    acAliDim.XLine2Point = Pnt2;
                    Point3d Midpnt = General_methods.GetMidPntBetweenPoints3d(Pnt1, Pnt2);
                    double dblAngle = General_methods.GetAnglePntBetween3dPoints(Pnt1, Pnt2);
                    double dist = General_methods.GetDistanceBetweenPoints(Pnt1, Pnt2);
                    acAliDim.DimLinePoint = Midpnt;

                    //acAliDim.TextRotation =General_methods.textReadbleAng(General_methods.DTR(dblAngle));
                    acAliDim.TextRotation = General_methods.DTR(General_methods.textReadbleAng(dblAngle));
                    acAliDim.DimensionStyle = acDoc.Database.Dimstyle;
                    acAliDim.Dimasz = 0.15;
                    acAliDim.Linetype = "ByLayer";
                    // Add the new object to Model space and the transaction
                    //acBlkTblRec.AppendEntity(acAliDim);
                    //acTrans.AddNewlyCreatedDBObject(acAliDim, true);
                    // Append a suffix to the dimension text
                    //acAliDim.DimensionText = "";
                    acAliDim.DimensionText = Math.Round(dist,1).ToString() + "'";
                    acAliDim.Dimaunit = 0;
                    acAliDim.Suffix = "'";
                    acAliDim.Layer = Layr;
                    acAliDim.ColorIndex = 7;
                    //acAliDim.Prefix = "";
                    //acAliDim.Dimtxt = Dist;
                    acAliDim.TextRotation = 0.0;
                    acAliDim.Dimtxt = 0.14;
                    acAliDim.Dimclrt = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 7);
                    acBlkTblRec.AppendEntity(acAliDim);
                    acTrans.AddNewlyCreatedDBObject(acAliDim, true);
                    //PromptStringOptions pStrOpts = new PromptStringOptions("");
                    //pStrOpts.Message = "\nEnter a new text suffix for the dimension: ";
                    //pStrOpts.AllowSpaces = true;
                    //PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
                    //if (pStrRes.Status == PromptStatus.OK)
                    //{
                    //    acAliDim.Suffix = pStrRes.StringResult;
                    //}
                    // Commit the changes and dispose of the transaction
                    id = acAliDim.ObjectId;
                    acTrans.Commit(); acTrans.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return id;
        }

        internal static ObjectId PlaceDimension(Document acDoc, Transaction acTr, double Dist, Point3d Pnt1, Point3d Pnt2, string Layr)
        {
            ObjectId id = new ObjectId();
            try
            {
                using (Transaction acTrans = acDoc.Database.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead) as BlockTable;
                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    // Create the aligned dimension
                    AlignedDimension acAliDim = new AlignedDimension();
                    acAliDim.SetDatabaseDefaults();
                    acAliDim.XLine1Point = Pnt1;
                    acAliDim.XLine2Point = Pnt2;
                    Point3d Midpnt = General_methods.GetMidPntBetweenPoints3d(Pnt1, Pnt2);
                    double dblAngle = General_methods.GetAnglePntBetween3dPoints(Pnt1, Pnt2);
                    double dist = General_methods.GetDistanceBetweenPoints(Pnt1, Pnt2);
                    acAliDim.DimLinePoint = Midpnt;

                    //acAliDim.TextRotation =General_methods.textReadbleAng(General_methods.DTR(dblAngle));
                    acAliDim.TextRotation = General_methods.DTR(General_methods.textReadbleAng(dblAngle));
                    acAliDim.DimensionStyle = acDoc.Database.Dimstyle;
                    acAliDim.Dimasz = 0.15;
                    acAliDim.Linetype = "ByLayer";
                    // Add the new object to Model space and the transaction
                    //acBlkTblRec.AppendEntity(acAliDim);
                    //acTrans.AddNewlyCreatedDBObject(acAliDim, true);
                    // Append a suffix to the dimension text
                    //acAliDim.DimensionText = "";
                    acAliDim.DimensionText = Math.Round(dist).ToString() + "'";
                    acAliDim.Dimaunit = 0;
                    acAliDim.Suffix = "'";
                    acAliDim.Layer = Layr;
                    acAliDim.ColorIndex = 7;
                    //acAliDim.Prefix = "";
                    //acAliDim.Dimtxt = Dist;
                    acAliDim.TextRotation = 0.0;
                    acAliDim.Dimtxt = 0.14;
                    acAliDim.Dimclrt = Teigha.Colors.Color.FromColorIndex(ColorMethod.ByLayer, 7);
                    acBlkTblRec.AppendEntity(acAliDim);
                    acTrans.AddNewlyCreatedDBObject(acAliDim, true);
                    //PromptStringOptions pStrOpts = new PromptStringOptions("");
                    //pStrOpts.Message = "\nEnter a new text suffix for the dimension: ";
                    //pStrOpts.AllowSpaces = true;
                    //PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);
                    //if (pStrRes.Status == PromptStatus.OK)
                    //{
                    //    acAliDim.Suffix = pStrRes.StringResult;
                    //}
                    // Commit the changes and dispose of the transaction
                    id = acAliDim.ObjectId;
                    acTrans.Commit(); acTrans.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return id;
        }


    }
}
