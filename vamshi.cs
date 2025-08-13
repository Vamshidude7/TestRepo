public class Vamshi
{
    public bool ValidateBlockStationing(Point3d blockPoint, Polyline centerline, Transaction acTr, Editor ed, Dictionary<string,string> DictVals)
    {

        List<(bool isValid, string actualSta, string expectedStaStr)> results = new List<(bool, string, string)>();
        // This method returns true if stationing matches within tolerance, false otherwise.
        //List<(bool isValid,)>
        // 1. Find closest point on centerline to blockPoint
        Point3d closestPoint = centerline.GetClosestPointTo(blockPoint, false);

        // 2. Get a buffer around closestPoint to find nearby StationValues polylines
        double searchBuffer = 60.0;
        Point3dCollection pntsBuffer = General_methods.funGetBuffPts(closestPoint, searchBuffer);

        SelectionSet sBuffStation = selectionset_methods.GetAcSelectionSetCrossPolygonLay(ed, pntsBuffer, "*line", DictVals["StationValues"]);

        if (sBuffStation == null || sBuffStation.Count == 0)
        {
            ed.WriteMessage("\nNo station values found near centerline.");
            return false;
        }

        var values = new List<(string StationValue, Polyline StationLine)>();

        foreach (ObjectId lineId in sBuffStation.GetObjectIds())
        {
            Entity ent = acTr.GetObject(lineId, OpenMode.ForRead) as Entity;
            if (ent is Polyline pline)
            {
                Point3d midPt = General_methods.GetMidPointsForEntity(pline);
                Point3dCollection midBuff = General_methods.funGetBuffPts(midPt, 10.0);

                SelectionSet sTextBuff = selectionset_methods.GetAcSelectionSetCrossPolygonLay(ed, midBuff, "*text", DictVals["StationValues"]);

                if (sTextBuff != null && sTextBuff.Count > 0)
                {
                    Entity entText = acTr.GetObject(sTextBuff.GetObjectIds()[0], OpenMode.ForRead) as Entity;
                    if (entText is MText mtext)
                    {
                        string mtextContent = mtext.Contents;
                        values.Add((mtextContent, pline));
                    }
                }
            }
        }

        if (values.Count == 0)
        {
            ed.WriteMessage("\nNo station text values found near centerline.");
            return false;
        }

        // Helper method to parse station string like "00+61" into int 61
        int ParseStationValue(string sta)
        {
            string cleaned = sta.Replace("+", "");
            return int.TryParse(cleaned, out int result) ? result : int.MaxValue;
        }

        // Sort values by parsed station value ascending
        values = values.OrderBy(v => ParseStationValue(v.StationValue)).ToList();

        // Take the nearest station polyline to blockPoint from values
        Polyline nearestStationLine = GetNearestPolyline(blockPoint, values.Select(v => v.StationLine).ToList());

        // Take the start point of nearestStationLine as base station point
        Point3d startStaPoint = nearestStationLine.GetPoint3dAt(0);

        // Distance between start station point and closestPoint on centerline
        double dist = General_methods.GetDistanceBetweenPoints(startStaPoint, closestPoint);

        // Base station value (lowest station in list)
        double baseSta = ParseStationValue(values.First().StationValue);

        // Expected station at block point = baseSta + dist
        double expectedSta = baseSta + dist;

        string expectedStaStr = FormatStation(expectedSta);

        // Now get the actual station value from the block's leader or attributes
        // Since you have it in notValidatedBlocks from earlier, you can pass it as parameter or
        // retrieve again here. For demo, let's just assume blockPoint is unique to get MLeader text

        // Find MLeader near blockPoint with STA text, similar to your original logic

        // Make a small buffer around blockPoint to find MLeader
        double leaderSearchRadius = 4.0;
        Point3dCollection leaderBuff = General_methods.funGetBuffPts(blockPoint, leaderSearchRadius);
        SelectionSet sMLeaders = selectionset_methods.GetAcSelectionSetCrossPolygonLay(ed, leaderBuff, "*", "*");

        string actualStaText = null;

        if (sMLeaders != null && sMLeaders.Count > 0)
        {
            foreach (ObjectId id in sMLeaders.GetObjectIds())
            {
                Entity ent = acTr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent is MLeader lead && lead.ContentType.ToString() == "MTextContent")
                {
                    string text = lead.MText?.Text;
                    if (!string.IsNullOrEmpty(text) && text.Contains("STA:"))
                    {
                        int idx = text.IndexOf("STA:") + 4;
                        if (text.Length >= idx + 5)
                        {
                            actualStaText = text.Substring(idx, 5);
                            break;
                        }
                    }
                }
            }
        }

        if (actualStaText == null)
        {
            ed.WriteMessage("\nNo STA text found near block.");
            return false;
        }

        double actualSta = ParseStationValue(actualStaText);

        // Check station difference tolerance (e.g. 0.5)
        if (Math.Abs(expectedSta - actualSta) > 0.5)
        {
            ed.WriteMessage($"\nStation mismatch! Expected: {expectedStaStr}, Actual: {actualStaText}");
            return false;
        }
        else
        {
            ed.WriteMessage($"\nStation validated successfully. Expected and Actual: {expectedStaStr}");
            return true;
        }
    }

}