package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.ChartDataLabel;
import com.aspose.words.ChartDataLabelCollection;
import com.aspose.words.ChartSeries;
import com.aspose.words.ChartType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

public class WorkWithChartDataLabelOfASingleChartSeries {
	
	public static final String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		Shape shape = builder.insertChart(ChartType.BAR, 432, 252);
		
		// Get first series.
		ChartSeries series0 = shape.getChart().getSeries().get(0);
		
		ChartDataLabelCollection dataLabelCollection = series0.getDataLabels();

		// Add data label to the first and second point of the first series.
		ChartDataLabel chartDataLabel00 = dataLabelCollection.add(0);
		ChartDataLabel chartDataLabel01 = dataLabelCollection.add(1);

		// Set properties.
		chartDataLabel00.setShowLegendKey(true);

		// By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
		// positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
		// corresponding data point.
		chartDataLabel00.setShowLeaderLines(true);

		chartDataLabel00.setShowCategoryName(false);
		chartDataLabel00.setShowPercentage(false);
		chartDataLabel00.setShowSeriesName(true);
		chartDataLabel00.setShowValue(true);
		chartDataLabel00.setSeparator("/");

		chartDataLabel01.setShowValue(true);
		
		doc.save(dataDir + "ChartDataLabelOfASingleChartSeries_out.docx");
	}

}