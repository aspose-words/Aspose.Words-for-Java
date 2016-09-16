package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.ChartDataPoint;
import com.aspose.words.ChartDataPointCollection;
import com.aspose.words.ChartSeries;
import com.aspose.words.ChartType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.MarkerSymbol;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

public class WorkWithSingleChartDataPointOfAChartSeries {
	
	public static final String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
		
		// Get first series.
		ChartSeries series0 = shape.getChart().getSeries().get(0);

		// Get second series.
		ChartSeries series1 = shape.getChart().getSeries().get(1);
		
		ChartDataPointCollection dataPointCollection = series0.getDataPoints();

		// Add data point to the first and second point of the first series.
		ChartDataPoint dataPoint00 = dataPointCollection.add(0);
		ChartDataPoint dataPoint01 = dataPointCollection.add(1);

		// Set explosion.
		dataPoint00.setExplosion(50);

		// Set marker symbol and size.
		dataPoint00.getMarker().setSymbol(MarkerSymbol.CIRCLE);
		dataPoint00.getMarker().setSize(15);

		dataPoint01.getMarker().setSymbol(MarkerSymbol.DIAMOND);
		dataPoint01.getMarker().setSize(20);

		// Add data point to the third point of the second series.
		ChartDataPoint dataPoint12 = series1.getDataPoints().add(2);
		dataPoint12.setInvertIfNegative(true);
		dataPoint12.getMarker().setSymbol(MarkerSymbol.STAR);
		dataPoint12.getMarker().setSize(20);
		
		doc.save(dataDir + "SingleChartDataPointOfAChartSeries_out.docx");
	}

}
