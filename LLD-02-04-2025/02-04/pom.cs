using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using IntelliCAD.ApplicationServices;
using Teigha.DatabaseServices;
using IntelliCAD.EditorInput;
using System.Security.Cryptography;
using Teigha.Geometry;
using System.Collections;
using Nini.Config;
using Application = IntelliCAD.ApplicationServices.Application;
using AttributeCollection = Teigha.DatabaseServices.AttributeCollection;
//using DocumentFormat.OpenXml.Bibliography;
using Editor = IntelliCAD.EditorInput.Editor;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Drawing;
//using DocumentFormat.OpenXml.Presentation;
using Teigha.Colors;
using Path = System.IO.Path;
using Color = Teigha.Colors.Color;




namespace ZiplyPermits
{
    public partial class Layout_Generation_Ziply_Permits : Form
    {
        string iniPath = string.Empty;
        List<string> lstLayNames = new List<string>();
        SortedDictionary<int, Entity> DictGrids = new SortedDictionary<int, Entity>();
        Dictionary<string, string> DictCVR = new Dictionary<string, string>(); Dictionary<string, Dictionary<string, string>> DictSheets = new Dictionary<string, Dictionary<string, string>>();
        Dictionary<string, string> DictVals = new Dictionary<string, string>(); SortedDictionary<string, Entity> SLDGrids = new SortedDictionary<string, Entity>(new NumericPrefixComparer());
        SortedDictionary<string, bool> val = new SortedDictionary<string, bool>();
        public string filename = string.Empty;
        public static string cityPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),"Config","Ziply_Permits","Layout Generation_Ziply Permits", "Coversheet", "city.dwg");
        public static string countpath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Config", "Ziply_Permits", "Layout Generation_Ziply Permits", "Coversheet", "county.dwg");
        public static string wsDotPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Config", "Ziply_Permits", "Layout Generation_Ziply Permits", "Coversheet", "ws-dot.dwg");
        public static string railroadPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Config", "Ziply_Permits", "Layout Generation_Ziply Permits", "Coversheet", "railroad.dwg");
        public static string jpnPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Config", "Ziply_Permits", "Layout Generation_Ziply Permits", "Coversheet", "jpn.dwg");
        public static string tcp = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Config", "Ziply_Permits", "Layout Generation_Ziply Permits", "Coversheet", "tcp.dwg");
        public static Dictionary<string, string> BlockPath = new Dictionary<string, string> 
                    { 
                        {"CITY", cityPath} ,
                        {"COUNTY", countpath},
                        {"WS DOT", wsDotPath},
                        {"RAILROAD",railroadPath},
                        {"JPN", jpnPath },
                        {"TCP",tcp }
                    };

        private string[] permitsData = new[] { "", "CITY", "TCP", "COUNTY", "WS DOT", "RAILROAD", "JPN" };
        private int sum = 0;
        private Point3d tickpoint = new Point3d();
        UpdateExistingData data = new UpdateExistingData();

        public Layout_Generation_Ziply_Permits()
        {
            InitializeComponent();
            tickmarkData.Items.Clear();
            tickmarkData.Items.AddRange(permitsData);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            iniPath = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Ziply_Permits.ini";
            Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            DateTime firstDate = DateTime.Today;
            DateTime secondDate = new DateTime(2026, 12, 30);
            int dResult = DateTime.Compare(firstDate, secondDate);
            if (dResult != 1)
            {
                var socketName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName.ToUpper();
                if (socketName != null && (socketName.Contains("INTECHWAVE") || socketName.Contains("IN.TECHWAVE.NET") || socketName.Contains("TECHWAVEANZ") || socketName.Contains("TWINH-DES") || socketName.Contains("TWINH-LAP")))
                {
                    try
                    {
                        DictVals = ini_methods.GetIniKeyFieldNvalues("Layout Creation", iniPath);
                        //DictCVR = ini_methods.GetIniKeyFieldNvalues("CVR", iniPath);
                        using (DocumentLock docLk = acDoc.LockDocument())
                        {
                            //deleteLayouts(acDoc);
                            using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                            {
                                if (acDoc == null)
                                {
                                    MessageBox.Show("No Active Document Found");
                                    return;
                                }
                                Database db = acDoc.Database;
                                filename = Path.GetFileNameWithoutExtension(db.Filename);
                                DictSheets = new Dictionary<string, Dictionary<string, string>>();
                                DictGrids = new SortedDictionary<int, Entity>();
                                SelectionSet Ssetgrid = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*Line", DictVals["GridLayer"]);

                                foreach (ObjectId objId in Ssetgrid.GetObjectIds())
                                {
                                    bool presence = false;
                                    string grdnum = string.Empty;
                                    Polyline pline = acTr.GetObject(objId, OpenMode.ForRead) as Polyline;
                                    if (pline != null)
                                    {
                                        Point3dCollection grdpts = General_methods.GetCoordinates1(pline);
                                        if (grdpts[0].DistanceTo(grdpts[grdpts.Count - 1]) > 0.001)
                                        {
                                            grdpts.Add(grdpts[0]);
                                        }
                                        SelectionSet ssetgridnumber = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, grdpts, "*Text", DictVals["GridsNumberLayer"]);
                                        if (ssetgridnumber != null && ssetgridnumber.Count > 0)
                                        {

                                            MText mt = acTr.GetObject(ssetgridnumber[0].ObjectId, OpenMode.ForRead) as MText;
                                            if (mt != null)
                                            {
                                                grdnum = mt.Text;
                                            }
                                        }
                                        SelectionSet sMicroDucts = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, grdpts, "*", "CALLOUTS");

                                        if (sMicroDucts != null && sMicroDucts.Count > 0)
                                        {
                                            foreach (ObjectId id in sMicroDucts.GetObjectIds())
                                            {
                                                Entity ent = acTr.GetObject(id, OpenMode.ForRead) as Entity;
                                                if (ent != null & ent is MLeader)
                                                {
                                                    MLeader lead = ent as MLeader;
                                                    string text = lead.MText.Text;
                                                    if (text.Contains("MICRODUCT"))
                                                    {
                                                        presence = true;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        val[grdnum] = presence;
                                    }
                                }
                                //SLDGrids = new Dictionary<string, Entity>();/*Need to check with the error*/
                                
                                SelectionSet sGrids = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*LINE", DictVals["GridLayer"]);
                                if (sGrids != null && sGrids.Count > 0)
                                {
                                    List<Entity> lstGrids = General_methods.GetlstFromsSet(acDoc.Editor, sGrids);
                                    for (int i = 0; i < lstGrids.Count; i++)
                                    {
                                        Entity entGrid = lstGrids[i];
                                        if (entGrid is Polyline)
                                        {
                                            Polyline blkRef = (Polyline)entGrid;
                                            Zoom_methods.ZoomToPolyline(blkRef);
                                            if (blkRef != null)
                                            {
                                                Point3dCollection pntcoll = General_methods.GetCoordinates3d(blkRef);
                                                SelectionSet sSetGridnum = selectionset_methods.GetAcSelectionSetCrossingPolgonGeom(acDoc.Editor, pntcoll, "*TEXT*", DictVals["GridsNumberLayer"]);
                                                if (sSetGridnum.Count > 0)
                                                {
                                                    int j = 1;
                                                    foreach (ObjectId objId in sSetGridnum.GetObjectIds())
                                                    {
                                                        Entity entTxt = acTr.GetObject(objId, OpenMode.ForRead) as Entity;
                                                        if ((entTxt is MText || entTxt is DBText) && entTxt.Layer == DictVals["GridsNumberLayer"])
                                                        {
                                                            if (entTxt is MText)
                                                            {
                                                                MText txtEnt = entTxt as MText;
                                                                string num = txtEnt.Text;
                                                                if (num.All(char.IsDigit) == true)
                                                                {
                                                                    if (!DictGrids.ContainsKey(int.Parse(num)))
                                                                    {
                                                                        DictGrids[int.Parse(num)] = entGrid;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (!SLDGrids.ContainsKey(txtEnt.Contents.Split(';').LastOrDefault().Replace("}", "")))
                                                                    {
                                                                        SLDGrids[txtEnt.Contents.Split(';').LastOrDefault().Replace("}", "")] = entGrid;
                                                                    }
                                                                    else
                                                                    {
                                                                        SLDGrids[txtEnt.Contents.Split(';').LastOrDefault().Replace("}", "") + "|" + j] = entGrid;
                                                                    }
                                                                }

                                                                getsheetInfo(acDoc, acTr, blkRef, num);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if (DictGrids != null && DictGrids.Count > 0)
                                {
                                    if (vancouver_check.Checked)
                                    {
                                        sum += 4;
                                    }
                                    else
                                    {
                                        sum += 3;
                                    }
                                    foreach (int eachGrdno in DictGrids.Keys)
                                    {
                                        sum += 1;
                                    }
                                    if (constr_typ.Checked)
                                    {
                                        sum += 1;
                                    }
                                    if (ped_details.Checked)
                                    {
                                        sum += 1;
                                    }
                                    if (cabi_net.Checked)
                                    {
                                        sum += 1;
                                    }
                                    if (flower_pot.Checked)
                                    {
                                        sum += 1;
                                    }
                                    if (fdh.Checked)
                                    {
                                        sum += 1;
                                    }
                                    if(vancouver_check.Checked)
                                    {
                                        if (sidewalk.Checked)
                                        {
                                            sum += 1;
                                        }
                                        if (pav.Checked)
                                        {
                                            sum += 1;
                                        }
                                        if (trench.Checked)
                                        {
                                            sum += 1;
                                        }
                                        if (transverse.Checked)
                                        {
                                            sum += 1;
                                        }
                                    }
                                    
                                }
                                //City", "TCP", "COUNTY", "RAILROAD", "JPN"

                                int countoftotal = 0;
                                int count = 0;
                                if (SLDGrids != null)
                                {
                                    int nuym = SLDGrids.Count;
                                    sum += nuym;
                                }
                                CreateSheet1(acDoc, acTr);
                                CreateSheet2(acDoc, acTr);
                                countoftotal = 2;
                                count = 2;
                                
                                if (vancouver_check.Checked)
                                {
                                    CreateSheet3(acDoc, acTr);
                                    countoftotal = 3;
                                    count = 3;
                                }

                                if (DictGrids != null && DictGrids.Count > 0)
                                {
                                    int cnt = 1; int Grdno = DictGrids.Keys.ElementAt(0);
                                    foreach (int eachGrdno in DictGrids.Keys)
                                    {
                                        CreateLayoutNew(eachGrdno.ToString(), DictGrids[eachGrdno], acDoc, acTr, Grdno);
                                        cnt += 1; Grdno++;
                                        countoftotal++; count++;
                                        foreach (string item in SLDGrids.Keys)
                                        {
                                            string GrdNom = new string(item.Where(x => char.IsDigit(x)).ToArray());
                                            //string GrdNom = item.Where(Char.IsDigit).ToString();
                                            if (eachGrdno == int.Parse(GrdNom))
                                            {
                                                CreateLayoutNew(item.ToString(), SLDGrids[item], acDoc, acTr, Grdno);
                                                cnt += 1; /*Grdno++;*/
                                                
                                            }
                                        }
                                    }
                                }
                                
                                //Default sheets
                                if (constr_typ.Checked)
                                {
                                    ConstructionTypical(acDoc, acTr, countoftotal);
                                    countoftotal += 1;
                                }
                                if (ped_details.Checked)
                                {
                                    PED(acDoc, acTr, countoftotal);
                                    countoftotal += 1;
                                }
                                if (cabi_net.Checked)
                                {
                                    Cabinet(acDoc, acTr, countoftotal);
                                    countoftotal += 1;
                                }
                                if (flower_pot.Checked)
                                {
                                    FlowerPot(acDoc, acTr, countoftotal);
                                    countoftotal += 1;
                                }
                                if (fdh.Checked)
                                {
                                    VaultMountFDH(acDoc, acTr, countoftotal);
                                    countoftotal += 1;
                                }
                                if (vancouver_check.Checked)
                                {
                                    if (sidewalk.Checked)
                                    {
                                        Sidewalkdesc(acDoc, acTr, countoftotal);
                                        countoftotal += 1;
                                    }
                                    if (pav.Checked)
                                    {
                                        Pavement(acDoc, acTr, countoftotal);
                                        countoftotal += 1;
                                    }
                                    if (trench.Checked)
                                    {
                                        TrenchRestoration(acDoc, acTr, countoftotal);
                                        countoftotal += 1;
                                    }
                                    if (transverse.Checked)
                                    {
                                        TrenchTransverse(acDoc, acTr, countoftotal);
                                        countoftotal += 1;
                                    }
                                }
                                CreateSheet_Additional_9(acDoc, acTr, countoftotal);
                                AddDatainLayouts(countoftotal, count);
                                acTr.Commit();
                                SLDGrids.Clear();
                                acDoc.Editor.Regen();
                                MessageBox.Show("Layout Generated Sucessfully!");
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        MessageBox.Show(ex.StackTrace);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("invalid attempt, access denied", "Layout Creation_Ziply Permits", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }
            }

        }

        public void AddDatainLayouts(int lastsheetnum, int count)
        {
            try
            {
                Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database db = acDoc.Database;
                using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                {

                    DBDictionary layoutDict = (DBDictionary)acTr.GetObject(db.LayoutDictionaryId, OpenMode.ForWrite);
                    foreach (DBDictionaryEntry entry in layoutDict)
                    {
                        Layout layout = (Layout)acTr.GetObject(entry.Value, OpenMode.ForWrite);
                        if (layout.LayoutName == "CD1")
                        {
                            BlockTableRecord layoutBlock = (BlockTableRecord)acTr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);
                            MText mtext = new MText
                            {
                                Location = new Point3d(0.5462, 2.9438, 0),
                                TextHeight = 0.085,
                                Width = 2.5418,
                                Height = 1.8930,
                                Contents = $"CD1 OF CD{lastsheetnum}. COVER SHEET\\P" +
                                $"CD2 OF CD{lastsheetnum}. LEGEND\\P" +
                                $"CD3 OF CD{lastsheetnum}. KEYMAP\\P" +
                                $"CD4 OF CD{lastsheetnum}. TO CD{count} OF CD{lastsheetnum} PLAN VIEWS\\P" +
                                $"CD{count + 2} OF CD{lastsheetnum}. CONSTRUCTION DETAILS\\P" +
                                $"CD{count + 3} OF CD{lastsheetnum}. PEDESTAL DETAILS\\P" +
                                $"CD{count + 4} OF CD{lastsheetnum}. HH &CABINET PLACING DETAILS\\P" +
                                $"CD{count + 5} OF CD{lastsheetnum}. HARD SURFACE CUTS-1\\P" +
                                $"CD{count + 6} OF CD{lastsheetnum}. HARD SURFACE CUTS-2 &\\P" +
                                $"ARTERIAL SURFACE CUTS-1\\P" +
                                $"CD{count + 7} OF CD{lastsheetnum}. ARTERIAL SURFACE CUTS-2\\P" +
                                $"CD{count + 8} OF CD{lastsheetnum}. ARTERIAL SURFACE CUTS-3\\P" +
                                $"CD{lastsheetnum} OF CD{lastsheetnum}. ADDITIONAL NOTES & DETAILS"
                            };

                            // Add the MText to the layout's BlockTableRecord
                            layoutBlock.AppendEntity(mtext);
                            acTr.AddNewlyCreatedDBObject(mtext, true);
                            break;
                        }
                    }
                    acTr.Commit();
                }
                acDoc.Editor.Regen();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void CreateSheet1(Document acDoc, Transaction acTr)
        {
            try
            {
                Database db = acDoc.Database;
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                //string fullPath = db.Filename;
                //string filename = Path.GetFileNameWithoutExtension(db.Filename);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("CRST", "CD1");
                layBlkData.Add("TOT", "CD" + sum.ToString());
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("MUNCIPALITY", muncipal.Text);

                //Above Hashtable Updates the attributes in the layouts 
                //Add tickmark points
                
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;

                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD1", acDoc);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);
                string BlkPath = string.Empty;
                layBlkPt = new Point3d(0.4652, 3.3965, 0);
                BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\IndexofSheet.dwg";
                InsertBlockGrid_N2P(BlkPath, "IndexofSheet", layBlkPt, layBlkData, DictVals, 0, 1.0);


                
                


                LayoutManager layoutMgr = LayoutManager.Current;
                Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                BlockTableRecord btr1 = acTr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                MText acMText = new MText();
                acMText.Location = new Point3d(0.7306, 9.6310, 0); // Set the location of the MText
                acMText.TextHeight = 0.1670; // Set the text height
                string val = string.Empty;
                val = muncipal.Text;
                string finalval = string.Empty;
                string addVal = tickmarkData.Text;
                if (tickmarkData.Text == "CITY")
                {
                    //addVal = "OF";
                    //if city
                    //PERMITTING AGENCY:
                    // CITY OF "VAL"
                    finalval = "CITY OF " + val;
                }
                else if (tickmarkData.Text == "WS DOT" || tickmarkData.Text == "RAILROAD")
                {
                    addVal = "";
                }
                //if(tickmarkData.Text == "CITY")
                //{

                //}
                acMText.Contents = (tickmarkData.Text == "CITY") ? "PERMITTING AGENCY:\\P" + finalval : "PERMITTING AGENCY:\\P" + val + " " + addVal; // Set the content of the MText
                acMText.ShowBorders = true;
                acMText.Width = 3.2978;
                // Add the new object to the block table record
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);

                acMText = new MText();
                acMText.Location = new Point3d(7.0366, 9.6861, 0); // Set the location of the MText
                acMText.TextHeight = 0.2170; // Set the text height
                acMText.Contents = "CONSTRUCTION PACKAGE# " + PrjNum.Text; // Set the content of the MText
                acMText.Attachment = AttachmentPoint.TopCenter;
                //acMText.ShowBorders = true;
                // Add the new object to the block table record
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);

                acMText = new MText();
                acMText.Location = new Point3d(6.8376, 8.9382, 0); // Set the location of the MText
                acMText.TextHeight = 0.1670; // Set the text height
                acMText.Contents = PrjAdd.Text; // Set the content of the MText
                acMText.Attachment = AttachmentPoint.TopCenter;
                //acMText.ShowBorders = true;
                // Add the new object to the block table record
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);


                acMText = new MText();
                acMText.Location = new Point3d(5.5788, 7.4316, 0); // Set the location of the MText
                acMText.TextHeight = 0.1800; // Set the text height
                acMText.Contents = "AERIAL & UG PERMIT"; // Set the content of the MText
                acMText.Attachment = AttachmentPoint.TopLeft;
                acMText.ShowBorders = true;
                // Add the new object to the block table record
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);

                acMText = new MText();
                acMText.Location = new Point3d(6.8333, 6.8516, 0); // Set the location of the MText
                acMText.TextHeight = 0.1800; // Set the text height
                acMText.Contents = @"\pxqc;\L{VICINITY MAP}\l"; ; // Set the content of the MText
                acMText.Attachment = AttachmentPoint.TopLeft;
                //acMText.
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);

                acMText = new MText();
                acMText.Location = new Point3d(11.6052, 6.5530, 0);
                acMText.TextHeight = 0.09;
                acMText.Contents = "MATERIALS:";
                acMText.Attachment = AttachmentPoint.TopLeft;
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);

                acMText = new MText();
                acMText.Location = new Point3d(11.5834, 4.8311, 0);
                acMText.TextHeight = 0.09;
                acMText.Contents = "PATH FOOTAGE:";
                acMText.Attachment = AttachmentPoint.TopLeft;
                btr1.AppendEntity(acMText);
                acTr.AddNewlyCreatedDBObject(acMText, true);

                //DrawingProperties drawingProperties = acDoc.DwgProperties;




                PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                acPlSet.CopyFrom(lay);
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "acad.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      //if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      //{
                //      //    vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //      //    layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      //}
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 1.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 1.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                lay.Dispose();


                acDoc.Editor.Regen();

                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\N2PBlocks\\CVR1.dwg";
                ////objId = PlaceFeature_methods.InsertBlockGrid_N2P(BlkPath, "NBNco-TITLE", layBlkPt, layBlkData, DictCVR, 0, 0.0395);

                //objId = InsertBlockGrid_N2P(BlkPath, "CVR1", layBlkPt, layBlkData, DictCVR, 0, 0.0395);
                //LayoutManager LM = LayoutManager.Current;
                //string currentLo = LM.CurrentLayout;
                //DBDictionary LayoutDict = acTr.GetObject(acDoc.Database.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;
                //Layout CurrentLo = acTr.GetObject((ObjectId)LayoutDict[currentLo], OpenMode.ForRead) as Layout;
                //BlockTableRecord BlkTblRec = acTr.GetObject(CurrentLo.BlockTableRecordId, OpenMode.ForRead) as BlockTableRecord;
                //foreach (ObjectId ID in BlkTblRec)
                //{
                //    Viewport VP = acTr.GetObject(ID, OpenMode.ForRead) as Viewport;
                //    if (VP != null)
                //    {
                //        VP.UpgradeOpen();
                //        VP.Erase();
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private bool ValidDbExtents(Point3d min, Point3d max)
        {
            return !(min.X > 0 && min.Y > 0 && min.Z > 0 && max.X < 0 && max.Y < 0 && max.Z < 0);
        }
        private void CreateSheet2(Document acDoc, Transaction acTr)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();

                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);

                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD2");
                layBlkData.Add("TOT", "CD" + sum.ToString());
                layBlkData.Add("MUNCIPALITY", muncipal.Text);

                //Attribute 
                //Add tickmark points
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;

                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD2", acDoc);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.4258, -0.2264, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Legend Drawing.dwg";
                InsertBlockGrid(BlkPath1, "Legend Drawing", layBlkPt1, 0, 0.99);

               

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void PED(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);

                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                // Open the created layout
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                string BlkPath = string.Empty;
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);
                Point3d layBlkPt1 = new Point3d(13.2000, 9.7929, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\PED.dwg";
                InsertBlockGrid(BlkPath1, "PED-DETAILS", layBlkPt1, 0, 0.99);
                
                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Cabinet(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                layBlkData.Add("MUNCIPALITY", muncipal.Text);

                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                // Open the created layout
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                string BlkPath = string.Empty;
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.2122, -0.2219, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Cabinet.dwg";
                InsertBlockGrid(BlkPath1, "Cabinet", layBlkPt1, 0, 0.99);

                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void ConstructionTypical(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();

                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                string BlkPath = string.Empty;
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(0.4048, 0.1467, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\ConstructionTypical.dwg";
                InsertBlockGrid(BlkPath1, "ConstructionTypical", layBlkPt1, 0, 0.10);

                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FlowerPot(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();

                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.2256, -0.2375, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\FlowerPot_Specs.dwg";
                InsertBlockGrid(BlkPath1, "FlowerPot_Specs", layBlkPt1, 0, 0.99);

               

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void VaultMountFDH(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();

                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                string BlkPath = string.Empty;
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.1054, -0.2665, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\VaultMountFDH.dwg";
                InsertBlockGrid(BlkPath1, "VaultMountFDH", layBlkPt1, 0, 0.99);

                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Sidewalkdesc(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();

                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.2580, -0.2812, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Sidewalkdesc.dwg";
                InsertBlockGrid(BlkPath1, "Sidewalk_desc", layBlkPt1, 0, 0.99);

                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Pavement(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
               
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);
                Point3d layBlkPt1 = new Point3d(5.4170, 5.4832, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Pavement.dwg";
                InsertBlockGrid(BlkPath1, "PavementRestorationUnit_details", layBlkPt1, 0, 0.99);

                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TrenchRestoration(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();

                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.2400, -0.2610, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\TrenchRestoration.dwg";
                InsertBlockGrid(BlkPath1, "TrenchRestoration", layBlkPt1, 0, 0.99);

               

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TrenchTransverse(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
               
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.2249, -0.1746, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\TrenchTransverse.dwg";
                InsertBlockGrid(BlkPath1, "TrenchTransverse", layBlkPt1, 0, 0.99);


                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CreateSheet_Additional_9(Document acDoc, Transaction acTr, int val)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD" + $"{val + 1}".ToString());
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Hashtable layBlkData1 = new Hashtable();
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + $"{val + 1}", acDoc);
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                Point3d layBlkPt1 = new Point3d(-0.2300, -0.2102, 0);
                string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\AdditionalNotes.dwg";
                InsertBlockGrid(BlkPath1, "AdditionalNotes", layBlkPt1, 0, 0.99);


                

                //PlaceFeature_methods.InsertBlockGrid(BlkPath, "Legend", layBlkPt, layBlkData1, DictVals, 0, 1.0);

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                //acPlSet.CopyFrom(lay);
                //PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                //PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                //plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                //ext = lay.GetMaximumExtents();
                //lay.ApplyToViewport(
                //  acTr, 2,
                //  vp =>
                //  {
                //      //vp.ResizeViewport(ext, 0.1);
                //      vp.ResizeViewportNew_CP(ext, 0.84);
                //      if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                //      {
                //          vp.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                //      }
                //      // Finally we lock the view to prevent meddling

                //      double scaleFactor = 1.0 / 40.0; // 1:40 scale
                //      vp.CustomScale = scaleFactor;

                //      double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                //      double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                //      // Open layout for write to apply the offset
                //      layout.UpgradeOpen();

                //      // Set the plot offset
                //      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                //      layout.CopyFrom(acPlSet);
                //      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //      vp.Locked = true;
                //  }
                //);
                //lay.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void CreateSheet3(Document acDoc, Transaction acTr)
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                Point3d layBlkPt = new Point3d();
                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                string strSht = (DictGrids.Count + 4).ToString();
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITTLE", title_text.Text);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT #", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("CRST", "CD3");
                layBlkData.Add("TOT", "CD" + sum.ToString());
                layBlkData.Add("MUNCIPALITY", muncipal.Text);

               

                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD3", acDoc);
                // Open the created layout
                // Open the created layout
                layBlkPt = new Point3d(0.1512, 0.1008, 0);
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                string outlineblockpath = BlockPath[tickmarkData.Text];
                string name = tickmarkData.Text;
                string BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Outline.dwg";
                ObjectId objId = InsertBlockGrid_N2P(outlineblockpath, name, layBlkPt, layBlkData, DictVals, 0, 0.99);

                
                

                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);
                //lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                ext = lay.GetMaximumExtents();
                Extents3d drawingExtents = GetDrawingExtents(db);
                PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                acPlSet.CopyFrom(lay);
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                LayoutManager layoutMgr = LayoutManager.Current;
                Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                plotSettings.CopyFrom(layout);
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWF6 ePlot.pc3");
                ext = lay.GetMaximumExtents();
                lay.ApplyToViewport(
                  acTr, 2,
                  vp =>
                  {
                      //vp.ResizeViewport(ext,0.1);
                      vp.ResizeViewportNew_CP_sheet3(ext);
                      vp.FitContentToViewport(drawingExtents, 1.0);
                      //if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                      //{
                      //vp.FitContentToViewport2(new Extents3d(ext[0], ext[1]), 1.0);
                      //    layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                      //}
                      // Finally we lock the view to prevent meddling

                      double scaleFactor = 1.0; // 1:40 scale
                      vp.CustomScale = scaleFactor;

                      double plotOffsetX = 1.0;//10.0 // Example offset in mm or inches depending on the units
                      double plotOffsetY = 1.0; //20.0// Example offset in mm or inches depending on the units

                      // Open layout for write to apply the offset
                      layout.UpgradeOpen();

                      // Set the plot offset
                      acPlSetVdr.SetPlotOrigin(acPlSet, new Point2d(plotOffsetX, plotOffsetY));
                      layout.CopyFrom(acPlSet);
                      lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "Monochrome.ctb", "DWF6 ePlot.pc3");
                      vp.Locked = true;
                  }
                );
                lay.Dispose();
                acDoc.Editor.Regen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CreateLayoutNew(string eachGrdnm, Entity entity, Document acDoc, Transaction acTr, int ShtNo)
        {
            try
            {
                System.IO.FileInfo fi = new System.IO.FileInfo(acDoc.Name);
                string docName = fi.Name.Substring(0, fi.Name.Length - 4);
                var ext = new Extents2d();
                //Point3d layBlkPt = new Point3d(0.15,0.2,0.0);
                Point3d layBlkPt = new Point3d();

                Hashtable layBlkData = new Hashtable();
                layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("SCALE", "1:40");
                string strSht = (DictGrids.Count + 4).ToString();// < 10 ? "00" + (DictGrids.Count + 3).ToString() : (DictGrids.Count + 3) >= 10 && (DictGrids.Count + 3) < 100 ? "0" + (DictGrids.Count + 3).ToString() : (DictGrids.Count + 3).ToString();
                string eachGrdnum = ShtNo < 10 ? ShtNo.ToString() : ShtNo.ToString();
                layBlkData.Add("PROJECTNO", PrjNum.Text);
                layBlkData.Add("COUNTY", County_1.Text);
                layBlkData.Add("CITY", City_1.Text);
                layBlkData.Add("TITLE", filename);
                layBlkData.Add("SCALE", DictVals["Scaling"]);
                layBlkData.Add("PROJECTADDRESS", PrjAdd.Text);
                layBlkData.Add("PROJECT#", PrjAdd.Text);
                layBlkData.Add("TOTAL PAGES", strSht);
                layBlkData.Add("OF", strSht);
                layBlkData.Add("ENGR", "TECHWAVE");
                layBlkData.Add("PHONE", "310-922-9638");
                layBlkData.Add("COAREA", textBox2.Text);
                layBlkData.Add("EXCHCODE", textBox3.Text);
                layBlkData.Add("FILE", file_text.Text);
                layBlkData.Add("TAX", textBox4.Text);
                layBlkData.Add("TWN", twnshp.Text);
                layBlkData.Add("RNG", rng_text.Text);
                layBlkData.Add("SEC", sec_text.Text);
                layBlkData.Add("DATE", DateTime.Now.ToString("MM/dd/yyyy"));
                layBlkData.Add("MUNCIPALITY", muncipal.Text);
                string finalval = eachGrdnm.All(Char.IsDigit) ? ShtNo.ToString() : eachGrdnm;
                layBlkData.Add("CRST", "CD" + finalval);//changed from StNo
                layBlkData.Add("TOT", "CD" + sum.ToString());
                Point3dCollection gridExtnts = General_methods.GetCoordinates3d(entity);
                gridExtnts.Add(gridExtnts[0]);
                var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD" + finalval, acDoc);
                // Open the created layout
                var lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                IntelliCAD.ApplicationServices.Application.SetSystemVariable("PSLTSCALE", 0);
                string shtname = eachGrdnm.All(Char.IsDigit) == true ? "Sheet" : "Sheet1";
                layBlkPt = new Point3d(0.0247, -0.0172, 0);
                string BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Outline_Block.dwg";
                string outlineBlkPath = BlockPath[tickmarkData.Text];
                ObjectId objId = InsertBlockGrid_N2P(outlineBlkPath, shtname/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.99);

                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Layout Generation_Ziply Permits\\Tickmark.dwg";
                //Point3d pointssss = new Point3d();
                //{ "", "CITY", "TCP", "COUNTY","WS DOT", "RAILROAD", "JPN" }
                //if (tickmarkData.Text == "RAILROAD")
                //{
                //    pointssss = new Point3d(13.6168, 6.3705, 0);
                //}
                //else if (tickmarkData.Text == "CITY")
                //{
                //    pointssss = new Point3d(13.6170, 7.0017, 0);
                //}
                //else if (tickmarkData.Text == "TCP")
                //{
                //    pointssss = new Point3d(13.6170, 6.8382, 0);
                //}

                //else if (tickmarkData.Text == "COUNTY")
                //{
                //    pointssss = new Point3d(13.6170, 6.6897, 0);
                //}

                //else if (tickmarkData.Text == "WS DOT")
                //{
                //    pointssss = new Point3d(13.6170, 6.5293, 0);
                //}
                //else if (tickmarkData.Text == "JPN")
                //{
                //    pointssss = new Point3d(13.6092, 6.2276, 0);
                //}
                //else
                //{
                //    pointssss = new Point3d(13.6170, 7.0017, 0);
                //}
                //InsertBlockGrid_N2P(BlkPath, "Tick", pointssss, layBlkData, DictVals, 0, 1.0);
                LayoutManager layoutMgr = LayoutManager.Current;
                Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                BlockTableRecord btr12 = acTr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                if (IsNumeric(eachGrdnm))
                {
                    if (val[eachGrdnm])
                    {
                        MText acMText = new MText();
                        acMText.Location = new Point3d(9.8595, 10.0072, 0); // Set the location of the MText
                        acMText.TextHeight = 0.1000; // Set the text height
                        acMText.Contents = "NOTE: ALL PROPOSED 1-MICRODUCTS ARE \\P1\" ROLLPIPES EXCEPT ROAD CROSSINGS"; // Set the content of the MText
                        acMText.ShowBorders = true;
                        acMText.Color = Color.FromRgb(255, 0, 0);
                        // Add the new object to the block table record
                        btr12.AppendEntity(acMText);
                        acTr.AddNewlyCreatedDBObject(acMText, true);
                    }
                }
                
                //lay.SetPlotSettings("ANSI_full_bleed_B_(11.00_x_17.00_Inches)", "acad.ctb", "DWG To PDF.pc3");

                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");
                //layBlkPt = new Point3d(14.0794, 3.6730, 0.0);
                //layBlkPt = new Point3d(13.7911, 3.5898, 0.0);
                //layBlkData = new Hashtable();
                //layBlkData.Add("DATE1", DateTime.Today.ToString("dd-MM-yyyy"));
                //layBlkData.Add("DEC1", "INITIAL REVIEW");
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\REV BLOCK 1.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "REV BLOCK 1"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 0.51);// 2);

                //layBlkPt = new Point3d(0.1, 0.05, 0.05);
                //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Construction.dwg";// Sheet.dwg";
                //objId = InsertBlockGrid_N2P(BlkPath, "Construction"/*"Sheet"*/, layBlkPt, layBlkData, DictVals, 0, 1.0);
                //Application.SetSystemVariable("PSLTSCALE", 1);
                if (DictSheets.ContainsKey(eachGrdnm.ToString()))
                {
                    if (DictSheets[eachGrdnm.ToString()].ContainsKey("X"))
                    {
                        layBlkData = new Hashtable();
                        string Num = DictSheets[eachGrdnm.ToString()]["X"].Contains("|") ? DictSheets[eachGrdnm.ToString()]["X"].Split('|')[0] : DictSheets[eachGrdnm.ToString()]["X"];
                        string NUm = Num.Replace("SHEET", "").Trim();
                        //string NUm = DictSheets[eachGrdnm.ToString()]["X"].Replace("SHEET", "").Trim();
                        //int index = DictGrids.Keys.ToList().IndexOf(NUm)+1;
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + index.ToString()+" OF "+ strSht);
                        double dblAngle = General_methods.DTR(90);
                        double dist = DictSheets[eachGrdnm.ToString()]["X"].Contains("|") ? Convert.ToDouble(DictSheets[eachGrdnm.ToString()]["X"].Split('|').LastOrDefault()) : 6.5;
                        Point3d AdjPnt = new Point3d(0.3777, 5.8890, 0.0);  /*new Point3d(layBlkPt.X + 0.5, layBlkPt.Y + 0.5, 0.0)*/
                        string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Adjacent_Arrow_4.dwg";
                        layBlkData.Add("NEXTPAGE", "SEE SHEET:" + "CD" + NUm);
                        ObjectId objId1 = PlaceFeature_methods.InsertBlockGrid(BlkPath1, "Adjacent_Arrow_4", AdjPnt, layBlkData, DictVals, dblAngle, 1.01);
                    }
                    if (DictSheets[eachGrdnm.ToString()].ContainsKey("Y"))
                    {
                        layBlkData = new Hashtable();
                        string Num = DictSheets[eachGrdnm.ToString()]["Y"].Contains("|") ? DictSheets[eachGrdnm.ToString()]["Y"].Split('|')[0] : DictSheets[eachGrdnm.ToString()]["Y"];
                        string NUm = Num.Replace("SHEET", "").Trim();
                        //int index = DictGrids.Keys.ToList().IndexOf(NUm)+1;
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + index.ToString() + " OF " + strSht);
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + DictSheets[eachGrdnum.ToString()]["Y"].Replace("SHEET", "").Trim());
                        double dblAngle = General_methods.DTR(0);
                        double dist = DictSheets[eachGrdnm.ToString()]["Y"].Contains("|") ? Convert.ToDouble(DictSheets[eachGrdnm.ToString()]["Y"].Split('|').LastOrDefault()) : 7.8;
                        Point3d AdjPnt = new Point3d(7.3952, 9.3526, 0.0);   /*new Point3d(layBlkPt.X +0.5, layBlkPt.Y + 0.5, 0.0);*/
                        layBlkData.Add("NEXTPAGE", "SEE SHEET:" + "CD" + NUm);
                        string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Adjacent_Arrow_4.dwg";
                        ObjectId objId1 = PlaceFeature_methods.InsertBlockGrid(BlkPath1, "Adjacent_Arrow_4", AdjPnt, layBlkData, DictVals, dblAngle, 1.01);
                    }
                    if (DictSheets[eachGrdnm.ToString()].ContainsKey("Z"))
                    {
                        layBlkData = new Hashtable();
                        string Num = DictSheets[eachGrdnm.ToString()]["Z"].Contains("|") ? DictSheets[eachGrdnm.ToString()]["Z"].Split('|')[0] : DictSheets[eachGrdnm.ToString()]["Z"];
                        string NUm = Num.Replace("SHEET", "").Trim();
                        //string NUm = DictSheets[eachGrdnm.ToString()]["Z"].Replace("SHEET", "").Trim();
                        //int index = DictGrids.Keys.ToList().IndexOf(NUm)+1;
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + index.ToString() + " OF " + strSht);
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + DictSheets[eachGrdnum.ToString()]["Z"].Replace("SHEET", "").Trim());
                        double dblAngle = General_methods.DTR(90);
                        double dist = DictSheets[eachGrdnm.ToString()]["Z"].Contains("|") ? Convert.ToDouble(DictSheets[eachGrdnm.ToString()]["Z"].Split('|').LastOrDefault()) : 4.8;
                        Point3d AdjPnt = new Point3d(12.7399, 5.7437, 0.0); /*new Point3d(layBlkPt.X + 0.5, layBlkPt.Y + 0.5, 0.0);*/
                        layBlkData.Add("NEXTPAGE", "SEE SHEET:" + "CD" + NUm);
                        string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Adjacent_Arrow_4.dwg";
                        ObjectId objId1 = PlaceFeature_methods.InsertBlockGrid(BlkPath1, "Adjacent_Arrow_4", AdjPnt, layBlkData, DictVals, dblAngle, 1.01);
                    }
                    if (DictSheets[eachGrdnm.ToString()].ContainsKey("A"))
                    {
                        layBlkData = new Hashtable();
                        string Num = DictSheets[eachGrdnm.ToString()]["A"].Contains("|") ? DictSheets[eachGrdnm.ToString()]["A"].Split('|')[0] : DictSheets[eachGrdnm.ToString()]["A"];
                        string NUm = Num.Replace("SHEET", "").Trim();
                        //string NUm = DictSheets[eachGrdnm.ToString()]["A"].Replace("SHEET", "").Trim();
                        //int index = DictGrids.Keys.ToList().IndexOf(NUm)+1;
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + index.ToString() + " OF " + strSht);
                        //layBlkData.Add("NEXTPAGE", "SEE SHEET " + DictSheets[eachGrdnum.ToString()]["A"].Replace("SHEET", "").Trim());
                        double dblAngle = General_methods.DTR(0);
                        double dist = DictSheets[eachGrdnm.ToString()]["A"].Contains("|") ? Convert.ToDouble(DictSheets[eachGrdnm.ToString()]["A"].Split('|').LastOrDefault()) : 7.8;
                        Point3d AdjPnt = new Point3d(4.8479, 0.3901, 0.0); /*new Point3d(layBlkPt.X + 0.5, layBlkPt.Y + 0.5, 0.0);*/
                        layBlkData.Add("NEXTPAGE", "SEE SHEET:" + "CD" + NUm);
                        string BlkPath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\Ziply_Permits\\Adjacent_Arrow_4.dwg";
                        ObjectId objId1 = PlaceFeature_methods.InsertBlockGrid(BlkPath1, "Adjacent_Arrow_4", AdjPnt, layBlkData, DictVals, dblAngle, 1.01);
                    }
                }




                PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                acPlSet.CopyFrom(lay);
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                //LayoutManager layoutMgr = LayoutManager.Current;
                //Layout layout = acTr.GetObject(layoutMgr.GetLayoutId(layoutMgr.CurrentLayout), OpenMode.ForWrite) as Layout;
                PlotSettings plotSettings = new PlotSettings(layout.ModelType);
                plotSettings.CopyFrom(layout);

                // Use the plot settings validator to set the plot scale
                PlotSettingsValidator plotSettingsValidator = PlotSettingsValidator.Current;

                // Set the plot scale type
                plotSettingsValidator.SetPlotType(plotSettings, Teigha.DatabaseServices.PlotType.Layout);

                // Commit the changes
                layout.CopyFrom(plotSettings);

                //double plotOffsetX = 2.0;//10.0 // Example offset in mm or inches depending on the units
                //double plotOffsetY = 2.0; //20.0// Example offset in mm or inches depending on the units

                //// Open layout for write to apply the offset
                //layout.UpgradeOpen();

                //// Set the plot offset
                //plotSettingsValidator.SetPlotOrigin(plotSettings, new Point2d(plotOffsetX, plotOffsetY));

                //// Copy the updated plot settings back to the layout
                //layout.CopyFrom(plotSettings);



                //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "acad.ctb", "DWG To PDF.pc3");//Previous one
                lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "acad.ctb", "DWF6 ePlot.pc3");//sridevi changed Recently

                //layout.CopyFrom(plotSettings);

                if (eachGrdnm.All(Char.IsDigit) == true)
                {
                    //layBlkPt = new Point3d(12.3, 0.9752, 0.0);
                    //BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\SCALE.dwg";
                    //objId = InsertBlockGrid_N2P(BlkPath, "SCALE", layBlkPt, layBlkData, DictVals, 0, 1.0);
                    ext = lay.GetMaximumExtents();
                    lay.ApplyToViewport(
                      acTr, 2,
                      vp =>
                      {
                          vp.ResizeViewportNew_ZiplyPermits(ext, 0.84);
                          if (ValidDbExtents(gridExtnts[0], gridExtnts[2]))
                          {
                              vp.FitContentToViewport_ATT(new Extents3d(gridExtnts[0], gridExtnts[2]), 1.05);
                              layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                          }
                          // Finally we lock the view to prevent meddling
                          double scaling = Convert.ToDouble(DictVals["Scaling"]);
                          double scaleFactor = 1.0 / scaling; // As per Config file
                          //double scaleFactor = 1.0 / 40.0; // 1:40 scale
                          vp.CustomScale = scaleFactor;
                          vp.Color = Color.FromRgb(255, 0, 0);


                          double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                          double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                          // Open layout for write to apply the offset
                          layout.UpgradeOpen();

                          // Set the plot offset
                          plotSettingsValidator.SetPlotOrigin(plotSettings, new Point2d(plotOffsetX, plotOffsetY));

                          #region Annotatie Scale

                          AnnotationScale annotScale = new AnnotationScale
                          {
                              Name = "1:40",
                              PaperUnits = 1.0,
                              DrawingUnits = 40.0
                          };

                          // Set the viewport's annotation scale to 1:40
                          vp.AnnotationScale = annotScale;

                          try
                          {
                              vp.UpdateDisplay();
                          }
                          catch (Exception)
                          {
                          }
                          //vp.ResizeViewportNew_ATT(ext, 0.84);
                          //if (ValidDbExtents(gridExtnts[0], gridExtnts[2]))
                          //{
                          //    vp.FitContentToViewport_ATT(new Extents3d(gridExtnts[0], gridExtnts[2]), 1.05);
                          //    layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                          //}
                          #endregion
                          //vp.StandardScale = StandardScaleType.Scale1To40;
                          layout.CopyFrom(plotSettings);


                          ObjectIdCollection layerIds = new ObjectIdCollection();

                          // Iterate layers and add ObjectIds to ObjectIdCollection.
                          LayerTable lt = acTr.GetObject(acDoc.Database.LayerTableId, OpenMode.ForRead) as LayerTable;
                          foreach (ObjectId idd in lt)
                          {
                              LayerTableRecord ltr = acTr.GetObject(idd, OpenMode.ForRead) as LayerTableRecord;
                              if (DictVals["GridsNumberLayer"].Contains(ltr.Name))
                              {
                                  layerIds.Add(idd);
                              }
                          }

                          //Selected Viewport for write.


                          //Freeze viewport layers.
                          vp.FreezeLayersInViewport(layerIds.GetEnumerator());
                          vp.Locked = true;
                      }
                    );
                    lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "acad.ctb", "DWF6 ePlot.pc3");
                    lay.Dispose();

                    //id = LayoutManager.Current.CreateAndMakeLayoutCurrent(eachGrdnum, acDoc);
                    //// Open the created layout
                    //lay = (Layout)acTr.GetObject(id, OpenMode.ForWrite);
                    //SelectionSet sSet = selectionset_methods.GetAcSelectionSetAllBlkGeomLayerLayout(acDoc.Editor, "INSERT", "Sheet", "0", eachGrdnum);
                    //if (sSet.Count > 0)
                    //{
                    //    BlockReference blkEnt1 = acTr.GetObject(sSet.GetObjectIds()[0], OpenMode.ForRead) as BlockReference;
                    //    Point3d BlkPos = blkEnt1.Position;
                    //    gridExtnts = General_methods.GetBlkRefExtents(entity);
                    //    //lay.SetPlotSettings("ISO_full_bleed_A1_(841.00_x_594.00_MM)", "monochrome.ctb", "DWG To PDF.pc3");
                    //    ext = lay.GetMaximumExtents();
                    //    Point3d Min = new Point3d(BlkPos.X + 26.08, BlkPos.Y + 1.76, 0.0);
                    //    Point3d Max = new Point3d(BlkPos.X + 29.58, BlkPos.Y + 4.06, 0.0);

                    //    Viewport vp1 = new Viewport();
                    //    var btr = (BlockTableRecord)acTr.GetObject(lay.BlockTableRecordId, OpenMode.ForWrite);
                    //    // Add it to the database
                    //    btr.AppendEntity(vp1);
                    //    acTr.AddNewlyCreatedDBObject(vp1, true);
                    //    vp1.On = true; vp1.GridOn = false;
                    //    vp1.ResizeViewportKeyPlan_ATT(Min, Max, 0.84);
                    //    //if (ValidDbExtents(gridExtnts[0], gridExtnts[1]))
                    //    //{
                    //    //    vp1.FitContentToViewport2(new Extents3d(gridExtnts[0], gridExtnts[1]), 1.0);
                    //    //    layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                    //    //}
                    //    vp1.Locked = true;
                    //    acPlSetVdr.SetPlotRotation(acPlSet, PlotRotation.Degrees180);
                    //    lay.Dispose();
                    //}
                    BlockTableRecord btr1 = acTr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite) as BlockTableRecord;
                    MText acMText = new MText();
                    acMText.Location = new Point3d(9.4275, 0.2105, 0); // Set the location of the MText
                    acMText.TextHeight = 0.15; // Set the text height
                    int eachGrdno = int.Parse(eachGrdnm);
                    int profilecnt = 0;
                    foreach (string item in SLDGrids.Keys)
                    {
                        string GrdNom = new string(item.Where(x => char.IsDigit(x)).ToArray());
                        if (GrdNom == eachGrdnm)
                        {
                            profilecnt++;
                        }
                    }
                    string val = string.Empty;

                    if (profilecnt > 0)
                    {
                        if (profilecnt == 1)
                        {
                            val = (ShtNo + 1).ToString();
                        }
                        else if (profilecnt == 1)
                        {
                            val = (ShtNo + 1).ToString() + " & " + (ShtNo + 2).ToString();
                        }
                        else
                        {
                            for (int i = 1; i <= profilecnt; i++)
                            {
                                val = string.IsNullOrEmpty(val) ? (ShtNo + i).ToString() + " , " : i == profilecnt - 1 ? val + (ShtNo + i).ToString() + " & " : i == profilecnt ? val + (ShtNo + i).ToString() : val + (ShtNo + i).ToString() + " & ";
                            }
                        }
                        acMText.Contents = "SEE PROFILE DRAWING ON SHEET " + val; // Set the content of the MText
                        acMText.ShowBorders = true;
                        // Add the new object to the block table record
                        btr1.AppendEntity(acMText);
                        acTr.AddNewlyCreatedDBObject(acMText, true);
                    }
                }
                else
                {
                    //CreateTextInLayout(lay);
                    string val = string.Empty;
                    if (SLDGrids != null)
                    {
                        foreach (string key in SLDGrids.Keys)
                        {
                            if (key.Contains(eachGrdnm))
                            {

                                val = key;
                                break;
                            }
                        }
                    }
                    
                    id = LayoutManager.Current.CreateAndMakeLayoutCurrent("CD"+val.ToString(), acDoc);
                    layBlkPt = new Point3d(10.33, 2.022, 0.0);
                    //layBlkPt = new Point3d(12.8004, 1.4496, 0.0);
                    BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\SCALE.dwg";
                    objId = InsertBlockGrid_N2P(BlkPath, "SCALE", layBlkPt, layBlkData, DictVals, 0, 1.0);

                    layBlkPt = new Point3d(5.37, 1.8, 0.0);
                    BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\ScaleVal.dwg";
                    objId = InsertBlockGrid_N2P(BlkPath, "ScaleVal", layBlkPt, layBlkData, DictVals, 0, 1.0);


                    //layBlkPt = new Point3d(12.02, 2.6441, 0.0);
                    layBlkPt = new Point3d(12.8004, 1.4496, 0.0);
                    BlkPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\ATT_Permits\\Scale_Ver.dwg";
                    objId = InsertBlockGrid_N2P(BlkPath, "Scale_Ver", layBlkPt, layBlkData, DictVals, 0, 1.0);
                    ext = lay.GetMaximumExtents();
                    //lay.ApplyToViewport(
                    //  acTr, 2,
                    //  vp =>
                    //  {
                    //      vp.ResizeViewportNew_ATT1(ext, 0.84);
                    //      vp.LinetypeScale = 1.0; 
                    //      if (ValidDbExtents(gridExtnts[0], gridExtnts[2]))
                    //      {
                    //          vp.FitContentToViewport_ATT(new Extents3d(gridExtnts[0], gridExtnts[2]), 1.05);
                    //          layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                    //      }
                    //      // Finally we lock the view to prevent meddling

                    //      vp.Locked = true;
                    //  }
                    //);

                    //lay.Dispose();

                    lay.ApplyToViewport(
                      acTr, 2,
                      vp =>
                      {
                          vp.ResizeViewportNew_ZiplyPermits(ext, 0.84);
                          if (ValidDbExtents(gridExtnts[0], gridExtnts[2]))
                          {
                              vp.FitContentToViewport_ATT(new Extents3d(gridExtnts[0], gridExtnts[2]), 1.05);
                              layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                          }
                          // Finally we lock the view to prevent meddling
                          double scaleFactor = 1.0 / 40.0; // 1:40 scale
                          vp.CustomScale = scaleFactor;


                          double plotOffsetX = 3.0;//10.0 // Example offset in mm or inches depending on the units
                          double plotOffsetY = 3.0; //20.0// Example offset in mm or inches depending on the units

                          // Open layout for write to apply the offset
                          layout.UpgradeOpen();

                          // Set the plot offset
                          plotSettingsValidator.SetPlotOrigin(plotSettings, new Point2d(plotOffsetX, plotOffsetY));

                          #region Annotatie Scale

                          AnnotationScale annotScale = new AnnotationScale
                          {
                              Name = "1:40",
                              PaperUnits = 1.0,
                              DrawingUnits = 40.0
                          };

                          // Set the viewport's annotation scale to 1:40
                          vp.AnnotationScale = annotScale;

                          try
                          {
                              vp.UpdateDisplay();
                          }
                          catch (Exception)
                          {
                          }
                          //vp.ResizeViewportNew_ATT(ext, 0.84);
                          //if (ValidDbExtents(gridExtnts[0], gridExtnts[2]))
                          //{
                          //    vp.FitContentToViewport_ATT(new Extents3d(gridExtnts[0], gridExtnts[2]), 1.05);
                          //    layer_methods.CrtTandChgLayer(acDoc.Editor, "VIEWPORT");
                          //}
                          #endregion
                          //vp.StandardScale = StandardScaleType.Scale1To40;
                          layout.CopyFrom(plotSettings);
                          vp.Locked = true;
                      }
                    );
                    lay.SetPlotSettings("ANSI_expand_B_(11.00_x_17.00_Inches)", "acad.ctb", "DWF6 ePlot.pc3");
                    lay.Dispose();

                }
                acDoc.Editor.Regen();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show(ex.StackTrace);
            }
        }

        public static bool CheckIfCallout(Entity ent)
        {
            Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            bool presence = false;
            try
            {
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                    {
                        if (ent != null)
                        {
                            if (ent is Polyline)
                            {
                                Point3dCollection pnts = General_methods.GetCoordinates3d(ent);
                                pnts.Add(pnts[0]);
                                SelectionSet sset = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pnts, "*", "CALLOUTS");
                                if (sset != null && sset.Count > 0)
                                {
                                    foreach (ObjectId id in sset.GetObjectIds())
                                    {
                                        Entity ents = acTr.GetObject(id, OpenMode.ForWrite) as Entity;
                                        if (ent is MLeader)
                                        {
                                            if (ent is MLeader)
                                            {
                                                MLeader lead = ent as MLeader;
                                                string val = lead.MText.Text;
                                                if (val.Contains("MICRODUCT"))
                                                {
                                                    presence = true;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            return presence;
        }
        public static Extents3d GetDrawingExtents(Database db)
        {
            Point3d extMin = db.Extmin;
            Point3d extMax = db.Extmax;
            return new Extents3d(extMax, extMin);
        }
        public static ObjectId InsertBlockGrid_N2P(string strBlkPath, string blkName, Point3d insPt, Hashtable hshList, Dictionary<string, string> DictVals, double dbAng, double blkScale)// string strFATID, string strCust, string strFATType)
        {

            BlockReference BlkRef = null;
            ObjectId BlkTblRecId = new ObjectId();
            try
            {
                Editor ed = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database db = ed.Document.Database;
                using (DocumentLock docLock = ed.Document.LockDocument())
                {
                    using (Transaction Trans = db.TransactionManager.StartTransaction())
                    {
                        BlockTable BlkTbl = Trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord LoRec = Trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlkTblRecId = PlaceFeature_methods.GetNonErasedTableRecordId(BlkTbl.Id, blkName);
                        if (File.Exists(strBlkPath))
                        {
                            BlkTbl.UpgradeOpen();
                            using (Database tempDb = new Database(false, true))
                            {
                                tempDb.ReadDwgFile(strBlkPath, FileShare.Read, true, null);
                                db.Insert(blkName, tempDb, false);
                            }
                            BlkTblRecId = PlaceFeature_methods.GetNonErasedTableRecordId(BlkTbl.Id, blkName);
                            LoRec.UpgradeOpen();
                            BlkRef = new BlockReference(insPt, BlkTblRecId);
                            BlkRef.TransformBy(Matrix3d.Scaling(blkScale, insPt));
                            BlkRef.ColorIndex = 7; BlkRef.Layer = "0";
                            BlkRef.ScaleFactors = blkScale == 0.98 ? new Scale3d(0.99, 0.99, 0.99) : new Scale3d(blkScale, blkScale, blkScale);
                            LoRec.AppendEntity(BlkRef);
                            Trans.AddNewlyCreatedDBObject(BlkRef, true);
                            BlockTableRecord BlkTblRec = Trans.GetObject(BlkTblRecId, OpenMode.ForWrite) as BlockTableRecord;
                            Teigha.DatabaseServices.AttributeCollection attColl = BlkRef.AttributeCollection;
                            if (BlkTblRec.HasAttributeDefinitions)
                            {
                                foreach (ObjectId objId in BlkTblRec)
                                {
                                    AttributeDefinition atrDef = Trans.GetObject(objId, OpenMode.ForWrite) as AttributeDefinition;
                                    if (atrDef != null)
                                    {
                                        if (hshList != null && hshList.Count > 0)
                                        {
                                            if (hshList.ContainsKey(atrDef.Tag) == true && !string.IsNullOrEmpty(hshList[atrDef.Tag].ToString()))
                                            {
                                                //if (atrDef.Tag=="CADREF")
                                                //{
                                                //    atrDef.Height = 0.1;
                                                //}
                                                atrDef.TextString = hshList[atrDef.Tag].ToString();
                                            }
                                            //else if (DictVals.ContainsKey(atrDef.Tag) == true)
                                            //{
                                            //    atrDef.TextString = DictVals[atrDef.Tag].ToString();
                                            //}
                                            AttributeReference AttRef = new AttributeReference();
                                            AttRef.SetAttributeFromBlock(atrDef, BlkRef.BlockTransform);
                                            BlkRef.AttributeCollection.AppendAttribute(AttRef);
                                            Trans.AddNewlyCreatedDBObject(AttRef, true);
                                        }
                                    }
                                }
                                if (dbAng != -1.0)//updating block rotation
                                {
                                    BlkRef.Rotation = dbAng;
                                }
                            }
                        }
                        Trans.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nError in function InsertBlock");
            }
            return BlkTblRecId;
        }

        public static ObjectId InsertBlockGrid(string strBlkPath, string blkName, Point3d insPt, double dbAng, double blkScale)// string strFATID, string strCust, string strFATType)
        {

            BlockReference BlkRef = null;
            ObjectId BlkTblRecId = new ObjectId();
            try
            {
                Editor ed = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                Database db = ed.Document.Database;
                using (DocumentLock docLock = ed.Document.LockDocument())
                {
                    using (Transaction Trans = db.TransactionManager.StartTransaction())
                    {
                        BlockTable BlkTbl = Trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;
                        BlockTableRecord LoRec = Trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        BlkTblRecId = PlaceFeature_methods.GetNonErasedTableRecordId(BlkTbl.Id, blkName);
                        if (File.Exists(strBlkPath))
                        {
                            BlkTbl.UpgradeOpen();
                            using (Database tempDb = new Database(false, true))
                            {
                                tempDb.ReadDwgFile(strBlkPath, FileShare.Read, true, null);
                                db.Insert(blkName, tempDb, false);
                            }
                            BlkTblRecId = PlaceFeature_methods.GetNonErasedTableRecordId(BlkTbl.Id, blkName);
                            LoRec.UpgradeOpen();
                            BlkRef = new BlockReference(insPt, BlkTblRecId);
                            BlkRef.TransformBy(Matrix3d.Scaling(blkScale, insPt));
                            BlkRef.ColorIndex = 7; BlkRef.Layer = "0";
                            BlkRef.ScaleFactors = blkScale == 0.98 ? new Scale3d(0.99, 0.99, 0.99) : new Scale3d(blkScale, blkScale, blkScale);
                            LoRec.AppendEntity(BlkRef);
                            Trans.AddNewlyCreatedDBObject(BlkRef, true);
                            BlockTableRecord BlkTblRec = Trans.GetObject(BlkTblRecId, OpenMode.ForWrite) as BlockTableRecord;
                            Teigha.DatabaseServices.AttributeCollection attColl = BlkRef.AttributeCollection;
                            if (BlkTblRec.HasAttributeDefinitions)
                            {
                                foreach (ObjectId objId in BlkTblRec)
                                {
                                    AttributeDefinition atrDef = Trans.GetObject(objId, OpenMode.ForWrite) as AttributeDefinition;
                                    if (atrDef != null)
                                    {
                                        //if (hshList != null && hshList.Count > 0)
                                        //{
                                        //if (hshList.ContainsKey(atrDef.Tag) == true && !string.IsNullOrEmpty(hshList[atrDef.Tag].ToString()))
                                        //{
                                        //    //if (atrDef.Tag=="CADREF")
                                        //    //{
                                        //    //    atrDef.Height = 0.1;
                                        //    //}
                                        //    atrDef.TextString = hshList[atrDef.Tag].ToString();
                                        //}
                                        //else if (DictVals.ContainsKey(atrDef.Tag) == true)
                                        //{
                                        //    atrDef.TextString = DictVals[atrDef.Tag].ToString();
                                        //}
                                        AttributeReference AttRef = new AttributeReference();
                                        AttRef.SetAttributeFromBlock(atrDef, BlkRef.BlockTransform);
                                        BlkRef.AttributeCollection.AppendAttribute(AttRef);
                                        Trans.AddNewlyCreatedDBObject(AttRef, true);
                                        //}
                                    }
                                }
                                if (dbAng != -1.0)//updating block rotation
                                {
                                    BlkRef.Rotation = dbAng;
                                }
                            }
                        }
                        Trans.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + "\nError in function InsertBlock");
            }
            return BlkTblRecId;
        }
        private void getsheetInfo(Document acDoc, Transaction acTr, Polyline blkEnt, string grid_Num)
        {
            try
            {
                Dictionary<string, string> sheets = new Dictionary<string, string>();
                Point3dCollection Pnts = General_methods.GetCoordinates3d(blkEnt);
                Point3d BlkPnt = Pnts[0];
                Point3dCollection pnts = General_methods.GetCoordinates3d(blkEnt);
                SelectionSet sSet = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pnts, "INSERT", "Sheet Joints");
                if (sSet != null && sSet.Count > 0)
                {
                    foreach (ObjectId BlkId in sSet.GetObjectIds())
                    {
                        BlockReference blkAdjEnt = acTr.GetObject(BlkId, OpenMode.ForRead) as BlockReference;
                        double Ang = blkAdjEnt.Rotation;
                        Ang = General_methods.RTD(Ang);
                        if ((Ang >= 0 && Ang <= 10) || (Ang >= 350 && Ang <= 360))
                        {
                            double dist = BlkPnt.Y - blkAdjEnt.Position.Y;

                            if (Math.Abs(dist) > 200)
                            {
                                double dist1 = Math.Abs(Pnts[3].X - blkAdjEnt.Position.X);
                                string nxtGrid_No = General_methods.getBlkAttVal(blkAdjEnt, "JOIN_A");
                                nxtGrid_No = new String(nxtGrid_No.Where(Char.IsDigit).ToArray());
                                sheets["A"] = !string.IsNullOrEmpty(nxtGrid_No) ? sheets.ContainsKey("A") ? sheets["A"] + "/" + nxtGrid_No : "SHEET " + nxtGrid_No : "";
                                double finedist = (dist1 / 50) * 1.175;
                                sheets["A"] = sheets["A"] + "|" + finedist;
                            }
                            else
                            {
                                double dist1 = Math.Abs(BlkPnt.X - blkAdjEnt.Position.X);
                                string nxtGrid_No = General_methods.getBlkAttVal(blkAdjEnt, "JOIN_B");
                                nxtGrid_No = new String(nxtGrid_No.Where(Char.IsDigit).ToArray());
                                sheets["Y"] = !string.IsNullOrEmpty(nxtGrid_No) ? sheets.ContainsKey("Y") ? sheets["Y"] + "/" + nxtGrid_No : "SHEET " + nxtGrid_No : "";
                                double finedist = (dist1 / 50) * 1.175;
                                sheets["Y"] = sheets["Y"] + "|" + finedist;
                            }
                        }
                        else if (Ang >= 80 && Ang <= 100)
                        {
                            double dist = BlkPnt.X - blkAdjEnt.Position.X;// General_methods.GetDistanceBetweenPoints(BlkPnt, blkAdjEnt.Position);
                            if (Math.Abs(dist) > 200)
                            {
                                double dist1 = Math.Abs(Pnts[3].Y - blkAdjEnt.Position.Y);
                                string nxtGrid_No = General_methods.getBlkAttVal(blkAdjEnt, "JOIN_A");
                                nxtGrid_No = new String(nxtGrid_No.Where(Char.IsDigit).ToArray());
                                sheets["Z"] = !string.IsNullOrEmpty(nxtGrid_No) ? sheets.ContainsKey("Z") ? sheets["Z"] + "/" + nxtGrid_No : "SHEET " + nxtGrid_No : "";
                                double finedist = (dist1 / 50) * 1.5;
                                sheets["Z"] = sheets["Z"] + "|" + finedist;
                            }
                            else
                            {
                                double dist1 = Math.Abs(Pnts[3].Y - blkAdjEnt.Position.Y);
                                string nxtGrid_No = General_methods.getBlkAttVal(blkAdjEnt, "JOIN_B");
                                nxtGrid_No = new String(nxtGrid_No.Where(Char.IsDigit).ToArray());
                                sheets["X"] = !string.IsNullOrEmpty(nxtGrid_No) ? sheets.ContainsKey("X") ? sheets["X"] + "/" + nxtGrid_No : "SHEET " + nxtGrid_No : "";
                                double finedist = (dist1 / 50) * 1.5;
                                sheets["X"] = sheets["X"] + "|" + finedist;
                            }
                        }

                    }
                    int eachGrdno = int.Parse(grid_Num);
                    string eachGrdnum = eachGrdno.ToString();// eachGrdno < 10 ? "00" + eachGrdno.ToString() : eachGrdno >= 10 && eachGrdno < 100 ? "0" + eachGrdno.ToString() : eachGrdno.ToString();
                    DictSheets[eachGrdnum] = sheets;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public bool IsNumeric(string str)
        {
            return str.All(c => "0123456789".Contains(c));
        }
        public void deleteLayouts()
        {
            try
            {
                Document acDoc = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (DocumentLock doclk = acDoc.LockDocument())
                {
                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                    {
                        DBDictionary db_Layouts = acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForWrite) as DBDictionary;
                        foreach (DBDictionaryEntry item in db_Layouts)
                        {
                            Layout lay = ((Layout)(acTrans.GetObject(item.Value, OpenMode.ForRead)));
                            string layname = lay.LayoutName;
                            if (!lstLayNames.Contains(layname) && layname.ToUpper() != "MODEL")
                            {
                                LayoutManager.Current.DeleteLayout(layname);
                            }
                        }
                        MessageBox.Show("Process Completed...Proceed with Layout Generation");
                        acTrans.Commit();
                        //acTrans.Dispose();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message + " in " + System.Reflection.MethodBase.GetCurrentMethod().Name);
            }
        }

        private void sheetjoin_Click(object sender, EventArgs e)
        {
            SheetJoints_Ziply sheetJoints_Ziply = new SheetJoints_Ziply();
            sheetJoints_Ziply.StartProcess();
        }

        private void dLayout_Click(object sender, EventArgs e)
        {
            deleteLayouts();
        }

        private void updateExistingData_Click(object sender, EventArgs e)
        {
            data.Show();
        }
    }
    public class NumericPrefixComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            int numX = ExtractNumericPart(x);
            int numY = ExtractNumericPart(y);

            int result = numX.CompareTo(numY);
            if (result == 0)
            {
                // If numbers are the same, fall back to full string comparison
                return string.Compare(x, y, StringComparison.OrdinalIgnoreCase);
            }

            return result;
        }
        private int ExtractNumericPart(string input)
        {
            if (string.IsNullOrEmpty(input))
                return int.MaxValue;

            var digits = new string(input.TakeWhile(char.IsDigit).ToArray());
            return int.TryParse(digits, out int number) ? number : int.MaxValue;
        }
    }
}
