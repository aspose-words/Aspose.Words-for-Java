package com.aspose.words.examples.linq;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.examples.Utils;

public class ChartSeries {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
        String dataDir = Utils.getDataDir(ChartSeries.class);

        SetChartSeriesNameDynamically(dataDir);
        System.out.println("\nChart template document is populated with the data.\nFile saved at " + dataDir);
	}
	
	public static void SetChartSeriesNameDynamically(String dataDir) throws Exception
	{
		// ExStart:SetChartSeriesNameDynamically
		List<PointData> data = new ArrayList<>();
		data.add(new PointData("12:00:00 AM", 10, 2));
		data.add(new PointData("01:00:00 AM", 15, 4));
		data.add(new PointData("02:00:00 AM", 23, 7));

        List<String> seriesNames = Arrays.asList("Flow","Rainfall");

        Document doc = new Document(dataDir + "ChartTemplate.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, new Object[] { data.toArray(new PointData[0]), seriesNames.toArray(new String[0]) }, new String[] { "data", "seriesNames" });

        doc.save(dataDir + "ChartTemplate_Out.docx");
        // ExEnd:SetChartSeriesNameDynamically
	}
}