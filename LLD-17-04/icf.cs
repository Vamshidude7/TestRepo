using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using IntelliCAD.ApplicationServices;
using IntelliCAD.EditorInput;
using System.Collections;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using Application = IntelliCAD.ApplicationServices.Application;
using Teigha.DatabaseServices;
using Teigha.LayerManager;
using Teigha.Runtime;
using System.Text.RegularExpressions;
using Teigha.Geometry;
using DocumentFormat.OpenXml.Wordprocessing;
using Document = IntelliCAD.ApplicationServices.Document;
using DocumentFormat.OpenXml.Office2016.Presentation.Command;
using ProjNet.CoordinateSystems;
using ProjNet.CoordinateSystems.Transformations;
using System.Globalization;
using System.Security.Cryptography;
using System.Threading;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace Test
{
    public partial class ReplaceBlocks_ICF : Form
    {
        public ReplaceBlocks_ICF()
        {
            InitializeComponent();
        }

        private void addlayers_Click(object sender, EventArgs e)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database db = acDoc.Database;
            string sourcePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\UGL_ICF\\Drawing Template.dwg";

            // Prefix-to-category mapping
            Dictionary<string, string> layerCategories = new Dictionary<string, string>
            {
                { "AB", "As-Built layers" },
                { "C", "Cadastral layers" },
                { "D", "Design (proposed) layers" },
                { "E", "Environmental and Culturally sensitive layers" },
                { "X", "Existing infrastructure and features" },
                { "G", "General plan layers" }
            };

            try
            {
                using (Database sourceDb = new Database(false, true))
                {
                    sourceDb.ReadDwgFile(sourcePath, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                    sourceDb.CloseInput(true);
                    Database destDb = acDoc.Database;

                    using (Transaction sourceTrans = sourceDb.TransactionManager.StartTransaction())
                    using (Transaction destTrans = destDb.TransactionManager.StartTransaction())
                    {
                        LayerTable sourceLayerTable = (LayerTable)sourceTrans.GetObject(sourceDb.LayerTableId, OpenMode.ForRead);
                        LayerTable destLayerTable = (LayerTable)destTrans.GetObject(destDb.LayerTableId, OpenMode.ForWrite);
                        LinetypeTable destLinetypeTable = (LinetypeTable)destTrans.GetObject(destDb.LinetypeTableId, OpenMode.ForRead);

                        Dictionary<string, List<string>> categoryLayers = new Dictionary<string, List<string>>();

                        if (sourceLayerTable != null)
                        {
                            foreach (ObjectId id in sourceLayerTable)
                            {
                                if (!id.IsValid) continue;

                                LayerTableRecord sourceLayer = (LayerTableRecord)sourceTrans.GetObject(id, OpenMode.ForRead);
                                string layerName = sourceLayer.Name;
                                string category = "Uncategorized";

                                foreach (var entry in layerCategories)
                                {
                                    if (layerName.StartsWith(entry.Key, StringComparison.OrdinalIgnoreCase))
                                    {
                                        category = entry.Value;
                                        break;
                                    }
                                }

                                if (!categoryLayers.ContainsKey(category))
                                {
                                    categoryLayers[category] = new List<string>();
                                }

                                LayerTableRecord destLayer;
                                bool layerExists = destLayerTable.Has(layerName);

                                if (layerExists)
                                {
                                    destLayer = (LayerTableRecord)destTrans.GetObject(destLayerTable[layerName], OpenMode.ForWrite);
                                }
                                else
                                {
                                    destLayer = new LayerTableRecord { Name = layerName };
                                }

                                // Update properties
                                destLayer.Color = sourceLayer.Color;
                                destLayer.IsPlottable = sourceLayer.IsPlottable;
                                destLayer.LineWeight = sourceLayer.LineWeight;

                                string linetypeName = "Continuous";
                                if (!sourceLayer.LinetypeObjectId.IsNull)
                                {
                                    LinetypeTableRecord ltr = sourceTrans.GetObject(sourceLayer.LinetypeObjectId, OpenMode.ForRead) as LinetypeTableRecord;
                                    if (ltr != null)
                                        linetypeName = ltr.Name;
                                }

                                if (destLinetypeTable.Has(linetypeName))
                                {
                                    destLayer.LinetypeObjectId = destLinetypeTable[linetypeName];
                                }
                                else
                                {
                                    destLayer.LinetypeObjectId = destDb.ContinuousLinetype;
                                }

                                if (!layerExists)
                                {
                                    destLayerTable.UpgradeOpen();
                                    destLayerTable.Add(destLayer);
                                    destTrans.AddNewlyCreatedDBObject(destLayer, true);
                                }

                                categoryLayers[category].Add(layerName);
                            }
                        }

                        destTrans.Commit();
                        sourceTrans.Commit();
                        MessageBox.Show("Layers and Layer Groups added Successfully!!!");
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show(ex.StackTrace);
            }
        }
        public Dictionary<string, string> DicVals = new Dictionary<string, string>();

        private void replace_blocks_Click(object sender, EventArgs e)
        {
            List<Entity> ducts = new List<Entity>();
            Dictionary<string, List<string>> FFNData = new Dictionary<string, List<string>> {
                { "72F", new List<string> {"SMOF G657A2 LL DUCT","SMOF G657A2 LL BURIED"} },
                { "144F", new List<string>{"SMOF G657A2 LL DUCT","SMOF G657A2 LL BURIED"} },
                { "360F", new List<string> {"SMOF G657A2 LL DUCT","SMOF G657A2 LL Underwater" } }

            };
            Dictionary<string, List<string>> EFNData = new Dictionary<string, List<string>> {
                { "72F", new List<string> { "SMOF G654C ULL AB DUCT", "SMOF G654C ULL AB BURIED" } },
                { "144F", new List<string>{ "SMOF G654C ULL AB DUCT", "SMOF G654C ULL AB BURIED" } },
                { "360F", new List<string> { "SMOF G654C ULL AB DUCT", "SMOF G654C ULL AB Underwater" } }

            };

            IntelliCAD.ApplicationServices.Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database db = acDoc.Database;
            string iniPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\UGL_ICF\\ReplaceBlocks.ini";
            DicVals = ini_methods.GetIniKeyFieldNvalues("ReplaceBlocks", iniPath);
            try
            {
                using (DocumentLock docLk = acDoc.LockDocument())
                {
                    if (acDoc == null)
                    {
                        MessageBox.Show("No active document found.");
                        return;
                    }
                    string blkpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\UGL_ICF";
                    using (Transaction acTr = acDoc.TransactionManager.StartTransaction())
                    {
                        bool delete = false;
                        SelectionSet set = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*", DicVals["Layer"]);
                        if (set != null && set.Count > 0)
                        {
                            foreach (ObjectId id in set.GetObjectIds())
                            {
                                Entity ent = acTr.GetObject(id, OpenMode.ForWrite, true) as Entity;
                                if (ent is DBPoint point)
                                {

                                    Point3d position = point.Position;
                                    string pit_id = XDATA_methods.GetxdataValueFld(point, "ACAD", "pit_id");
                                    string type = XDATA_methods.GetxdataValueFld(point, "ACAD", "sub_type");
                                    string size = XDATA_methods.GetxdataValueFld(point, "ACAD", "pit_size");
                                    string material = XDATA_methods.GetxdataValueFld(point, "ACAD", "pit_material");
                                    string faptype = XDATA_methods.GetxdataValueFld(point, "ACAD", "equip_type");
                                    string excCode = XDATA_methods.GetxdataValueFld(point, "ACAD", "exchange_code");
                                    string equip_id = XDATA_methods.GetxdataValueFld(point, "ACAD", "equip_id");
                                    Hashtable tab = new Hashtable();
                                    tab.Add("ID", pit_id);
                                    if (faptype.Contains("Exchange"))
                                    {
                                        Hashtable tab_1 = new Hashtable();
                                        tab_1.Add("ID", equip_id);
                                        string exchange = blkpath + "\\Proposed Infra\\X-EXCHANGE.dwg";
                                        Gen_Block_Enity_with_Att_ConduitBlock(exchange, "X-EXCHANGE", position, tab_1, 0.0, 0.0, "Exchange_site");
                                        delete = true;
                                    }
                                    else if (faptype.Contains("FAP") && faptype.Contains("EFN"))
                                    {
                                        Hashtable tab_1 = new Hashtable();
                                        string[] stringarray = equip_id.Split(':');
                                        tab_1.Add("(EXCH)", excCode);
                                        tab_1.Add("FAP", stringarray[1]);
                                        string exchange = blkpath + "\\Proposed Infra\\FAP EFN Proposed.dwg";
                                        Gen_Block_Enity_with_Att_ConduitBlock(exchange, "FAP EFN Proposed", position, tab_1, 0.0, 0.0, "D-Proposed EFN FAP");
                                        delete = true;
                                    }
                                    else if (faptype.Contains("FAP") && faptype.Contains("FFN"))
                                    {
                                        Hashtable tab_1 = new Hashtable();
                                        string[] stringarray = equip_id.Split(':');
                                        tab_1.Add("(EXCH)", excCode);
                                        tab_1.Add("FAP", stringarray[1]);
                                        string exchange = blkpath + "\\Proposed Infra\\FAP FFN Proposed.dwg";
                                        Gen_Block_Enity_with_Att_ConduitBlock(exchange, "FAP EFN Proposed", position, tab_1, 0.0, 0.0, "D-Proposed FFN FAP");
                                        delete = true;
                                    }
                                    if (type == "proposed")
                                    {
                                        if (material.Contains("Manhole"))
                                        {
                                            string Manhole = blkpath + "\\Proposed Infra\\D-PIT-Manhole proposed.dwg";

                                            Gen_Block_Enity_with_Att_ConduitBlock(Manhole, "D-PIT-Manhole proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                            delete = true;
                                        }
                                        else if (material.Contains("Thermoplastic"))
                                        {
                                            if (size == "1")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-1 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-1 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "2")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-2 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-1 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "3")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-3 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-3 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "4")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-4 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-4 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "5")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-5 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-5 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "6")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-6 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-6 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "7")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-7 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-7 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "8")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-8 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-8 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }
                                            else if (size == "9")
                                            {
                                                string pit = blkpath + "\\Proposed Infra\\D-PIT-9 proposed.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "D-PIT-9 proposed", position, tab, 0.0, 0.0, "D-Proposed Pit");
                                                delete = true;
                                            }

                                        }
                                        else if (material.Contains("Other"))
                                        {
                                            GenCircle_NEW(1.5, position, "D-Proposed Pit");
                                        }
                                    }
                                    else if (type == "existing")
                                    {
                                        if (material.Contains("Manhole"))
                                        {
                                            string Manhole = blkpath + "\\Existing Infra\\X-PIT-MANHOLE.dwg";
                                            Gen_Block_Enity_with_Att_ConduitBlock(Manhole, "Manhole", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                            delete = true;
                                        }
                                        else if (material.Contains("Thermoplastic"))
                                        {
                                            if (size == "1")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-1.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-1", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");

                                                delete = true;
                                            }
                                            else if (size == "2")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-2.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-2", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "3")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-3.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-3", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "4")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-4.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-4", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "5")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-5.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-5", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "6")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-6.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-6", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "7")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-7.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-7", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "8")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-8.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-8", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "9")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-9.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-9", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                        }
                                        else if (material.Contains("Hand Moulded") || material.Contains("Injection Moulded"))
                                        {
                                            if (size == "A")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-A.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-A", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "B")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-B.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-B", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "C")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-C.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-C", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "D")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-D.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-D", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "1")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-1.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-1", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "2")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-2.dwg";

                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-2", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "3")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-3.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-3", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "4")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-4.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-4", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "5")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-5.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-5", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "6")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-6.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-6", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "7")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-7.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-7", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "8")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-8.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-8", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;
                                            }
                                            else if (size == "9")
                                            {
                                                string pit = blkpath + "\\Existing Infra\\X-PIT-9.dwg";
                                                Gen_Block_Enity_with_Att_ConduitBlock(pit, "X-PIT-9", position, tab, 0.0, 0.0, "X-TELSTRA-Pit");
                                                delete = true;

                                            }
                                        }
                                        else if (material.Contains("Other"))
                                        {
                                            GenCircle_NEW(1.5, position, "X-TELSTRA-Pit");
                                            delete = true;
                                        }

                                    }
                                    if (delete)
                                    {
                                        ent.Erase();
                                    }
                                    tab.Clear();
                                }

                            }
                        }
                        #region 04/04/2025
                        SelectionSet SurveyPnts = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*Point", DicVals["SurveyPointsLayer"]);

                        if (SurveyPnts != null && SurveyPnts.Count > 0)
                        {
                            foreach (ObjectId id in SurveyPnts.GetObjectIds())
                            {
                                Entity ent = acTr.GetObject(id, OpenMode.ForWrite, true) as Entity;
                                if (ent is DBPoint point)
                                {
                                    Point3d PointPos = point.Position;
                                    string Type = XDATA_methods.GetxdataValueFld(point, "ACAD", "point_markup_type");

                                    if (Type == "Survey Peg proposed")
                                    {
                                        string distance = XDATA_methods.GetxdataValueFld(point, "ACAD", "distance_nearest_feature");
                                        string latitude = XDATA_methods.GetxdataValueFld(point, "ACAD", "Y");
                                        string longitude = XDATA_methods.GetxdataValueFld(point, "ACAD", "X");
                                        string peg_no = XDATA_methods.GetxdataValueFld(point, "ACAD", "peg_no");
                                        string secondPart = string.Empty;
                                        if (!string.IsNullOrWhiteSpace(peg_no) && peg_no.Contains("-"))
                                        {
                                            var parts = peg_no.Split('-');
                                            if (parts.Length > 1)
                                            {
                                                secondPart = parts[1];
                                            }
                                        }
                                        Hashtable hashtable = new Hashtable();
                                        List<string> eastings = ConvertToUTM(latitude, longitude);

                                        if (eastings != null && eastings.Count >= 2)
                                        {
                                            string easting = eastings[0];
                                            string northing = eastings[1];
                                            double new1 = Math.Round(Convert.ToDouble(easting), 2);
                                            double new2 = Math.Round(Convert.ToDouble(northing), 2);
                                            string new1_1 = new1.ToString();
                                            string new2_2 = new2.ToString();
                                            hashtable.Add("CABLE_OFFSET", distance);
                                            hashtable.Add("COORDINATES_1", "E " + new1_1);
                                            hashtable.Add("COORDINATES_2", "N " + new2_2);
                                            hashtable.Add("PEG_NO.", secondPart);
                                        }
                                        string blkpath1 = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Config\\UGL_ICF\\DesignStake.dwg";
                                        Gen_Block_Enity_with_Att_ConduitBlock(blkpath1, "Design_Stake", PointPos, hashtable, 0.0, 0.0, "Design Stake");
                                    }
                                    else if (Type == "Flag")
                                    {
                                        string comment = XDATA_methods.GetxdataValueFld(point, "ACAD", "comments");
                                        Point3d newPnt = General_methods.PolarPoint(PointPos, General_methods.DTR(45.0), 15.0);
                                        ObjectId objId = CreateColouredMText(acDoc.Editor, newPnt, acDoc, comment, 5.0);
                                    }
                                }
                            }
                        }
                        SelectionSet sGrids = selectionset_methods.GetAcSelectionSetAllGeomLayer(acDoc.Editor, "*insert", "GRID");
                        if (sGrids != null && sGrids.Count > 0)
                        {
                            foreach (ObjectId ids in sGrids.GetObjectIds())
                            {
                                Entity ent = acTr.GetObject(ids, OpenMode.ForRead) as Entity;
                                Point3dCollection pntsGrid = General_methods.GetBlkRefExtentsAll(ent);
                                SelectionSet sCadastres = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntsGrid, "*line", DicVals["CadastreLayer"]);
                                if (sCadastres != null && sCadastres.Count > 0)
                                {
                                    foreach (ObjectId objId in sCadastres.GetObjectIds())
                                    {
                                        string grdnum = string.Empty;
                                        Polyline pline = acTr.GetObject(objId, OpenMode.ForRead) as Polyline;
                                        if (pline != null)
                                        {
                                            Point3dCollection pnts = General_methods.GetCoordinates1(pline);
                                            if (pnts[0].DistanceTo(pnts[pnts.Count - 1]) > 0.001)
                                            {
                                                pnts.Add(pnts[0]);
                                            }

                                            Point3d centroid = CalculateCentroid(pnts);
                                            string address = XDATA_methods.GetxdataValueFld(pline, "ACAD", "address");
                                            string title = XDATA_methods.GetxdataValueFld(pline, "ACAD", "title_search_id");
                                            #region Added 21/4
                                            string parcel = XDATA_methods.GetxdataValueFld(pline, "ACAD", "planparcel");
                                            string main = string.Empty;
                                            string part1 = string.Empty;
                                            string part2 = string.Empty;
                                            string addVal = string.Empty;
                                            if (!string.IsNullOrEmpty(parcel))
                                            {
                                                int splitIndex = parcel.IndexOf('L');
                                                if (splitIndex != -1 && splitIndex < parcel.Length - 1)
                                                {
                                                    part1 = parcel.Substring(0, splitIndex); //Gets the substring until L indexNumber
                                                    part2 = parcel.Substring(splitIndex);
                                                }
                                            }
                                            if (!string.IsNullOrEmpty(address) && char.IsLetter(address[0]))
                                            {
                                                string[] parts = address.Split(' ');
                                                if (parts.Length > 1)
                                                    addVal = string.Join(" ", parts[0], parts[1]);
                                            }
                                            else if (!string.IsNullOrEmpty(address) && char.IsNumber(address[0]))
                                            {
                                                Match match = Regex.Match(address, @"^\d+");
                                                if (match.Success)
                                                {
                                                    addVal = match.Value;
                                                }
                                            }
                                            if(!string.IsNullOrEmpty(address) && char.IsNumber(address[0]))
                                            {
                                                PlaceFeature_methods.CreateColouredMText(ed, /*General_methods.PolarPoint(centroid, General_methods.DTR(270), 20)*/centroid, acDoc, addVal, 5.0);
                                            }
                                            else if(!string.IsNullOrEmpty(address) && char.IsLetter(address[0]))
                                            {
                                                PlaceFeature_methods.CreateColouredMText(ed, centroid, acDoc, part2, 5.0);
                                                PlaceFeature_methods.CreateColouredMText(ed, General_methods.PolarPoint(centroid, General_methods.DTR(270), 10), acDoc, part1, 5.0);
                                            }
                                            #endregion 04/21
                                        }
                                    }
                                }
                                SelectionSet selectionSet = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntsGrid, "*line", DicVals["DuctLayer"]);
                                if (selectionSet != null && selectionSet.Count > 0)
                                {
                                    foreach (ObjectId id in selectionSet.GetObjectIds())
                                    {
                                        Entity ent1 = acTr.GetObject(id, OpenMode.ForWrite) as Entity;
                                        if (ent1 is Line line1)
                                        {
                                            //Polyline line1 = ent1 as Polyline;
                                            if (!ducts.Contains(line1))
                                            {
                                                ducts.Add(line1);
                                                double rot = line1.Angle;
                                                Point3d midpnt = General_methods.GetMidPointsForEntity(line1);
                                                Point3dCollection buff = General_methods.funGetBuffPts(midpnt, 1.5);
                                                string FFN = string.Empty;
                                                string EFN = string.Empty;
                                                SelectionSet sDuct = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerFFN"]);
                                                if (sDuct != null && sDuct.Count > 0)
                                                {
                                                    Polyline ffnLine = acTr.GetObject(sDuct.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                                                    FFN = XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_id_unique");
                                                    string size = XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_size");
                                                    string spec = XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_spec");
                                                    Point3d ffnpnt = General_methods.PolarPoint(midpnt, General_methods.DTR(90), 20.0);
                                                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0);
                                                    //PlaceFeature_methods.CreateColouredMText1(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0, rot, AttachmentPoint.MiddleMid);
                                                    CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, "D-PROPOSED FFN CABLE", FFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                                                }
                                                SelectionSet sDuct1 = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerEFN"]);
                                                if (sDuct1 != null && sDuct1.Count > 0)
                                                {
                                                    Polyline efnLine = acTr.GetObject(sDuct1.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                                                    EFN = XDATA_methods.GetxdataValueFld(efnLine, "ACAD", "cable_id_unique");
                                                    string size = XDATA_methods.GetxdataValueFld(efnLine, "ACAD", "cable_size");
                                                    string spec = XDATA_methods.GetxdataValueFld(efnLine, "ACAD", "cable_spec");
                                                    Point3d efnpnt = General_methods.PolarPoint(midpnt, General_methods.DTR(90), 10.0);
                                                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, efnpnt, acDoc, EFN, 5.0);
                                                    CreateColouredMText(acDoc.Editor, efnpnt, acDoc, "D-PROPOSED EFN CABLE", EFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                                                }
                                            }
                                        }
                                        else if (ent1 is Polyline pline1)
                                        {
                                            
                                            if (!ducts.Contains(pline1))
                                            {
                                                ducts.Add(pline1);
                                                Point3dCollection Cordinates = General_methods.GetCoordinates3d(pline1);
                                                Point3d midpnt = General_methods.GetMidPointsForEntity(pline1);
                                                Point3dCollection buff = General_methods.funGetBuffPts(midpnt, 1.5);
                                                double rot = General_methods.GetAnglePntBetween3dPoints(Cordinates[0], Cordinates[Cordinates.Count - 1]);
                                                rot = GetPolylineAngleAtMidpoint(pline1, midpnt);
                                                string FFN = string.Empty;
                                                string EFN = string.Empty;
                                                SelectionSet sDuct = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerFFN"]);
                                                if (sDuct != null && sDuct.Count > 0)
                                                {
                                                    Polyline ffnLine = acTr.GetObject(sDuct.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                                                    FFN = XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_id_unique");
                                                    XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_size");
                                                    XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_spec");
                                                    Point3d ffnpnt = General_methods.PolarPoint(midpnt, /*General_methods.DTR(90)*/General_methods.DTR(rot), 30.0);
                                                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0);
                                                    //PlaceFeature_methods.CreateColouredMText1(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0, rot, AttachmentPoint.MiddleMid);
                                                    CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, "D-PROPOSED FFN CABLE", FFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                                                }
                                                SelectionSet sDuct1 = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerEFN"]);
                                                if (sDuct1 != null && sDuct1.Count > 0)
                                                {
                                                    Polyline efnLine = acTr.GetObject(sDuct1.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                                                    EFN = XDATA_methods.GetxdataValueFld(efnLine, "ACAD", "cable_id_unique");
                                                    Point3d efnpnt = General_methods.PolarPoint(midpnt, /*General_methods.DTR(90)*/ General_methods.DTR(rot), 10.0);
                                                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, efnpnt, acDoc, EFN, 5.0);
                                                    CreateColouredMText(acDoc.Editor, efnpnt, acDoc, "D-PROPOSED EFN CABLE", EFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            //From here
                            #region BackupCode#Vamshi
                            //SelectionSet selectionSet = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, pntsGrid, "*line", DicVals["DuctLayer"]);
                            //if (selectionSet != null && selectionSet.Count > 0)
                            //{
                            //    foreach (ObjectId id in selectionSet.GetObjectIds())
                            //    {
                            //        Entity ent1 = acTr.GetObject(id, OpenMode.ForWrite) as Entity;
                            //        if (ent1 is Line line1)
                            //        {
                            //            //Polyline line1 = ent1 as Polyline;
                            //            if (!ducts.Contains(line1))
                            //            {
                            //                ducts.Add(line1);
                            //                double rot = line1.Angle;
                            //                Point3d midpnt = General_methods.GetMidPointsForEntity(line1);
                            //                Point3dCollection buff = General_methods.funGetBuffPts(midpnt, 10.0);
                            //                string FFN = string.Empty;
                            //                string EFN = string.Empty;
                            //                SelectionSet sDuct = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerFFN"]);
                            //                if (sDuct != null && sDuct.Count > 0)
                            //                {
                            //                    Polyline ffnLine = acTr.GetObject(sDuct.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                            //                    FFN = XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_id_unique");
                            //                    XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_size");
                            //                    XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_spec");
                            //                    Point3d ffnpnt = General_methods.PolarPoint(midpnt, General_methods.DTR(90), 20.0);
                            //                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0);
                            //                    //PlaceFeature_methods.CreateColouredMText1(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0, rot, AttachmentPoint.MiddleMid);
                            //                    CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, "D-PROPOSED FFN CABLE", FFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                            //                }
                            //                SelectionSet sDuct1 = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerEFN"]);
                            //                if (sDuct1 != null && sDuct1.Count > 0)
                            //                {
                            //                    Polyline efnLine = acTr.GetObject(sDuct1.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                            //                    EFN = XDATA_methods.GetxdataValueFld(efnLine, "ACAD", "cable_id_unique");
                            //                    Point3d efnpnt = General_methods.PolarPoint(midpnt, General_methods.DTR(90), 10.0);
                            //                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, efnpnt, acDoc, EFN, 5.0);
                            //                    CreateColouredMText(acDoc.Editor, efnpnt, acDoc, "D-PROPOSED EFN CABLE", FFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                            //                }
                            //            }
                            //        }
                            //        else if (ent1 is Polyline pline1)
                            //        {
                            //            //Polyline line1 = ent1 as Polyline;
                            //            if (!ducts.Contains(pline1))
                            //            {
                            //                ducts.Add(pline1);
                            //                //double rot = pline1.Angle;
                            //                //line1.
                            //                Point3dCollection Cordinates = General_methods.GetCoordinates3d(pline1);

                            //                Point3d midpnt = General_methods.GetMidPointsForEntity(pline1);
                            //                Point3dCollection buff = General_methods.funGetBuffPts(midpnt, 10.0);
                            //                //General_methods.
                            //                double rot = General_methods.GetAnglePntBetween3dPoints(Cordinates[0], Cordinates[Cordinates.Count - 1]);
                            //                string FFN = string.Empty;
                            //                string EFN = string.Empty;
                            //                SelectionSet sDuct = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerFFN"]);
                            //                if (sDuct != null && sDuct.Count > 0)
                            //                {
                            //                    Polyline ffnLine = acTr.GetObject(sDuct.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                            //                    FFN = XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_id_unique");
                            //                    XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_size");
                            //                    XDATA_methods.GetxdataValueFld(ffnLine, "ACAD", "cable_spec");
                            //                    Point3d ffnpnt = General_methods.PolarPoint(midpnt, General_methods.DTR(90), 20.0);
                            //                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0);
                            //                    //PlaceFeature_methods.CreateColouredMText1(acDoc.Editor, ffnpnt, acDoc, FFN, 5.0, rot, AttachmentPoint.MiddleMid);
                            //                    CreateColouredMText(acDoc.Editor, ffnpnt, acDoc, "D-PROPOSED FFN CABLE", FFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                            //                }
                            //                SelectionSet sDuct1 = selectionset_methods.GetAcSelectionSetCrossPolygonLay(acDoc.Editor, buff, "*line", DicVals["CableLayerEFN"]);
                            //                if (sDuct1 != null && sDuct1.Count > 0)
                            //                {
                            //                    Polyline efnLine = acTr.GetObject(sDuct1.GetObjectIds()[0], OpenMode.ForWrite) as Polyline;
                            //                    EFN = XDATA_methods.GetxdataValueFld(efnLine, "ACAD", "cable_id_unique");
                            //                    Point3d efnpnt = General_methods.PolarPoint(midpnt, General_methods.DTR(90), 10.0);
                            //                    //PlaceFeature_methods.CreateColouredMText(acDoc.Editor, efnpnt, acDoc, EFN, 5.0);
                            //                    CreateColouredMText(acDoc.Editor, efnpnt, acDoc, "D-PROPOSED EFN CABLE", FFN, 5.0, rot, AttachmentPoint.MiddleCenter);
                            //                }
                            //            }
                            //        }
                            //    }
                            //}
                            #endregion BackupCode

                            #endregion 04/04/2025
                            acTr.Commit();
                            ducts.Clear();
                            MessageBox.Show("Blocks Replaced Successfully!!!");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show(ex.StackTrace);
            }
        }

        public static double GetPolylineAngleAtMidpoint(Polyline pline, Point3d midpoint)
        {
            double minDist = double.MaxValue;
            double angleAtMid = 0.0;

            for (int i = 0; i < pline.NumberOfVertices - 1; i++)
            {
                Point3d pt1 = pline.GetPoint3dAt(i);
                Point3d pt2 = pline.GetPoint3dAt(i + 1);

                Vector3d segVec = pt2 - pt1;
                Point3d midSeg = pt1 + (segVec * 0.5);

                double distToMid = midpoint.DistanceTo(midSeg);
                if (distToMid < minDist)
                {
                    minDist = distToMid;
                    angleAtMid = Math.Atan2(segVec.Y, segVec.X); // <-- Use Math.Atan2
                }
            }

            return angleAtMid;
        }

        public static ObjectId GenCircle_NEW(double bufferdist, Point3d inspt, string LayName)
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = acDoc.Editor;
            Database acCurDb = acDoc.Database;
            DocumentLock dcl = ed.Document.LockDocument();
            ObjectId cirobjid;
            try
            {
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    // Open the Block table for read
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                    // Open the Block table record Model space for write
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    // Create a circle that is at 2,3 with a radius of 4.25
                    using (Circle acCirc = new Circle())
                    {
                        acCirc.Center = new Point3d(inspt.X, inspt.Y, inspt.Z);
                        acCirc.Radius = bufferdist;
                        layer_methods.CrtTandChgLayer(acDoc.Editor, LayName);
                        acCirc.Layer = LayName;
                        // Add the new object to the block table record and the transaction
                        acBlkTblRec.AppendEntity(acCirc);
                        acTrans.AddNewlyCreatedDBObject(acCirc, true);
                        cirobjid = acCirc.ObjectId;
                    }

                    // Save the new object to the database
                    acTrans.Commit();
                    dcl.Dispose();
                }
            }
#pragma warning disable CS0168 // The variable 'ex' is declared but never used
            catch (System.Exception ex)
#pragma warning restore CS0168 // The variable 'ex' is declared but never used
            {

                throw;
            }
            return cirobjid;
        }

        static public ObjectId CreateColouredMText(Editor ed, Point3d mtLoc, Document doc, string TxtLyr, string strMtxt, double txtht, double txtAngle, AttachmentPoint txtJustify)
        {
            Database db = doc.Database;
            // Variables for our MText entity's identity
            // and location

            ObjectId mtId;
            //Point3d mtLoc = Point3d.Origin;
            DocumentLock docLock = doc.LockDocument();
            Transaction tr = db.TransactionManager.StartTransaction();

            using (tr)
            {
                // Create our new MText and set its properties

                //mt.TextStyleId = ;
                //mt.Width = txtDefWdt;
                //mt.TextHeight = txtDefHt;

                // Open the block table, the model space and
                // add our MText
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                TextStyleTable ts = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                ObjectId mtStyleid = db.Textstyle;

                if (ts.Has("isocpeur"))
                {
                    mtStyleid = ts["isocpeur"];
                }
                else
                {
                    if (ts.Has("CALIBRI"))
                    {
                        mtStyleid = ts["CALIBRI"];
                    }
                    else if (ts.Has("Standard"))
                    {
                        mtStyleid = ts["Standard"];
                    }
                }
                MText mt = new MText();
                mt.Location = mtLoc;
                mt.Contents = strMtxt;
                mt.TextStyleId = mtStyleid;
                mt.TextHeight = txtht;
                mt.Rotation = txtAngle;
                layer_methods.CrtTandChgLayer(ed, TxtLyr);
                mt.Layer = TxtLyr;
                mt.ColorIndex = 256; //if want to update as per ByLayer, then it has to be 256
                mt.Attachment = txtJustify;
                mtId = ms.AppendEntity(mt);
                tr.AddNewlyCreatedDBObject(mt, true);
                // Finally we commit our transaction
                tr.Commit();
                return mtId;
            }
        }

        private static Point3d CalculateCentroid(Point3dCollection points)
        {
            double x = 0, y = 0, area = 0;

            int count = points.Count;
            for (int i = 0; i < count; i++)
            {
                Point3d p1 = points[i];
                Point3d p2 = points[(i + 1) % count];

                double cross = (p1.X * p2.Y) - (p2.X * p1.Y);
                area += cross;
                x += (p1.X + p2.X) * cross;
                y += (p1.Y + p2.Y) * cross;
            }

            area *= 0.5;
            x /= (6 * area);
            y /= (6 * area);
            double z = 0;
            foreach (Point3d pt in points)
            {
                z += pt.Z;
            }
            z /= count;

            return new Point3d(x, y, z);
        }


        public static List<string> ConvertToUTM(string latitudeStr, string longitudeStr)
        {
            var results = new List<string>();


            if (!double.TryParse(latitudeStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double latitude))
            {
                results.Add("Invalid latitude input.");
                return results;
            }

            if (!double.TryParse(longitudeStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double longitude))
            {
                results.Add("Invalid longitude input.");
                return results;
            }

            const double a = 6378137.0;
            const double f = 1 / 298.257223563;
            const double k0 = 0.9996;

            double e = Math.Sqrt(f * (2 - f));
            double e1sq = e * e / (1 - e * e);
            double n = a / Math.Sqrt(1 - Math.Pow(e * Math.Sin(DegToRad(latitude)), 2));

            int zone = (int)Math.Floor((longitude + 180) / 6) + 1;
            double lonOrigin = (zone - 1) * 6 - 180 + 3;
            double lonOriginRad = DegToRad(lonOrigin);

            double latRad = DegToRad(latitude);
            double lonRad = DegToRad(longitude);

            double T = Math.Pow(Math.Tan(latRad), 2);
            double C = e1sq * Math.Pow(Math.Cos(latRad), 2);
            double A = Math.Cos(latRad) * (lonRad - lonOriginRad);

            double M = a * ((1
                - e * e / 4
                - 3 * Math.Pow(e, 4) / 64
                - 5 * Math.Pow(e, 6) / 256) * latRad
                - (3 * e * e / 8
                + 3 * Math.Pow(e, 4) / 32
                + 45 * Math.Pow(e, 6) / 1024) * Math.Sin(2 * latRad)
                + (15 * Math.Pow(e, 4) / 256
                + 45 * Math.Pow(e, 6) / 1024) * Math.Sin(4 * latRad)
                - (35 * Math.Pow(e, 6) / 3072) * Math.Sin(6 * latRad));

            double easting = k0 * n * (A + (1 - T + C) * Math.Pow(A, 3) / 6
                + (5 - 18 * T + T * T + 72 * C - 58 * e1sq) * Math.Pow(A, 5) / 120) + 500000;

            double northing = k0 * (M + n * Math.Tan(latRad) * (A * A / 2
                + (5 - T + 9 * C + 4 * C * C) * Math.Pow(A, 4) / 24
                + (61 - 58 * T + T * T + 600 * C - 330 * e1sq) * Math.Pow(A, 6) / 720));

            if (latitude < 0)
                northing += 10000000;


            results.Add($"{easting}");
            results.Add($"{northing}");

            return results;
        }

        private static double DegToRad(double deg) => deg * Math.PI / 180.0;


        internal static ObjectId Gen_Block_Enity_with_Att_ConduitBlock(string strBlkPath, string strBlkName, Point3d pnt3d, Hashtable hshList, double dbReadAng, double plhandle, string BlkLay)
        {
            ObjectId id = new ObjectId();
            if (!string.IsNullOrEmpty(strBlkPath) && !string.IsNullOrEmpty(strBlkName))
            {
                try
                {
                    Editor ed = IntelliCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                    using (DocumentLock dcl = ed.Document.LockDocument())
                    {
                        TransactionManager tm = ed.Document.Database.TransactionManager;
                        using (Transaction myTm = tm.StartTransaction())
                        {
                            BlockTable btt = (BlockTable)tm.GetObject(ed.Document.Database.BlockTableId, OpenMode.ForRead);

                            if (!btt.Has(strBlkName + ".dwg"))
                            {
                                if (File.Exists(strBlkPath))
                                {
                                    using (Database srcdbBlk = new Database(false, true))
                                    {
                                        srcdbBlk.ReadDwgFile(strBlkPath, FileShare.Read, true, "");
                                        id = ed.Document.Database.Insert(strBlkName + ".dwg", srcdbBlk, false);
                                    }

                                    if (!id.IsNull)
                                    {
                                        BlockTableRecord btr = (BlockTableRecord)tm.GetObject(ed.Document.Database.CurrentSpaceId, OpenMode.ForWrite);
                                        BlockReference bref = new BlockReference(pnt3d, id)
                                        {
                                            Layer = string.IsNullOrEmpty(BlkLay) ? "0" : BlkLay // Explicitly set layer
                                        };

                                        Matrix3d curUCSMatrix = ed.CurrentUserCoordinateSystem;
                                        CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;
                                        bref.TransformBy(Matrix3d.Rotation(General_methods.DTR(dbReadAng), curUCS.Zaxis, pnt3d));
                                        btr.AppendEntity(bref);
                                        tm.AddNewlyCreatedDBObject(bref, true);

                                        BlockTableRecord btrTemp = (BlockTableRecord)tm.GetObject(btt[strBlkName + ".dwg"], OpenMode.ForRead);
                                        if (btrTemp.HasAttributeDefinitions)
                                        {
                                            foreach (ObjectId id1 in btrTemp)
                                            {
                                                AttributeDefinition attDef = myTm.GetObject(id1, OpenMode.ForRead) as AttributeDefinition;
                                                if (attDef != null && !attDef.Constant)
                                                {
                                                    using (AttributeReference attRef = new AttributeReference())
                                                    {
                                                        attRef.SetAttributeFromBlock(attDef, bref.BlockTransform);
                                                        if (hshList.ContainsKey(attRef.Tag))
                                                            attRef.TextString = hshList[attRef.Tag].ToString();

                                                        bref.AttributeCollection.AppendAttribute(attRef);
                                                        myTm.AddNewlyCreatedDBObject(attRef, true);
                                                    }
                                                }
                                            }
                                        }
                                        myTm.Commit();
                                        return bref.ObjectId;
                                    }
                                }
                            }
                            else
                            {
                                ObjectId objblkid = btt[strBlkName + ".dwg"];
                                if (!objblkid.IsNull)
                                {
                                    BlockTableRecord ms = (BlockTableRecord)tm.GetObject(btt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                                    BlockReference br = new BlockReference(pnt3d, objblkid)
                                    {
                                        Layer = string.IsNullOrEmpty(BlkLay) ? "0" : BlkLay // Explicitly set layer
                                    };

                                    Matrix3d curUCSMatrix = ed.CurrentUserCoordinateSystem;
                                    CoordinateSystem3d curUCS = curUCSMatrix.CoordinateSystem3d;
                                    br.TransformBy(Matrix3d.Rotation(General_methods.DTR(dbReadAng), curUCS.Zaxis, pnt3d));
                                    ms.AppendEntity(br);
                                    tm.AddNewlyCreatedDBObject(br, true);

                                    BlockTableRecord btr = (BlockTableRecord)tm.GetObject(objblkid, OpenMode.ForRead);
                                    if (btr.HasAttributeDefinitions)
                                    {
                                        foreach (ObjectId id1 in btr)
                                        {
                                            AttributeDefinition attDef = myTm.GetObject(id1, OpenMode.ForRead) as AttributeDefinition;
                                            if (attDef != null && !attDef.Constant)
                                            {
                                                using (AttributeReference attRef = new AttributeReference())
                                                {
                                                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                                                    if (hshList.ContainsKey(attRef.Tag))
                                                        attRef.TextString = hshList[attRef.Tag].ToString();

                                                    br.AttributeCollection.AppendAttribute(attRef);
                                                    myTm.AddNewlyCreatedDBObject(attRef, true);
                                                }
                                            }
                                        }
                                    }
                                    myTm.Commit();
                                    return br.ObjectId;
                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message + "\nError in function Gen_Block_Enity_with_Att_ConduitBlock");
                }
            }
            return id;
        }

        static public ObjectId CreateColouredMText(Editor ed, Point3d mtLoc, Document doc, string strMtxt, double txtht)
        {
            Database db = doc.Database;
            // Variables for our MText entity's identity
            // and location

            ObjectId mtId;
            //Point3d mtLoc = Point3d.Origin;
            DocumentLock docLock = doc.LockDocument();
            Transaction tr = db.TransactionManager.StartTransaction();

            using (tr)
            {
                // Create our new MText and set its properties

                //mt.TextStyleId = ;
                //mt.Width = txtDefWdt;
                //mt.TextHeight = txtDefHt;

                // Open the block table, the model space and
                // add our MText
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                TextStyleTable ts = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);
                ObjectId mtStyleid = db.Textstyle;
                if (ts.Has("Standard"))
                {
                    mtStyleid = ts["Standard"];
                }
                MText mt = new MText();
                mt.Location = mtLoc;
                mt.Contents = strMtxt;
                mt.TextStyleId = mtStyleid;
                mt.TextHeight = txtht;
                mt.ShowBorders = true;

                mtId = ms.AppendEntity(mt);
                tr.AddNewlyCreatedDBObject(mt, true);
                // Finally we commit our transaction
                tr.Commit();
                return mtId;
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
