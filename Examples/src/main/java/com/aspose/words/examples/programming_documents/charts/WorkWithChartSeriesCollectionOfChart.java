package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.Chart;
import com.aspose.words.ChartSeriesCollection;
import com.aspose.words.ChartType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;

public class WorkWithChartSeriesCollectionOfChart {

	public static void main(String[] args) throws Exception {
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
		// Chart property of Shape contains all chart related options.
		Chart chart = shape.getChart();
				
		// Get chart series collection.
		ChartSeriesCollection seriesCollection = chart.getSeries();

		// Check series count.
		System.out.println(seriesCollection.getCount());
	}

}
