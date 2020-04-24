package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.Chart;
import com.aspose.words.ChartDataLabelCollection;
import com.aspose.words.ChartSeries;
import com.aspose.words.ChartType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.examples.Utils;

public class DefaultOptionsForDataLabels {

	public static final String dataDir = Utils.getSharedDataDir(DefaultOptionsForDataLabels.class) + "Charts/";

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		// ExStart:DefaultOptionsForDataLabels
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		Shape shape = builder.insertChart(ChartType.PIE, 432, 252);
		Chart chart = shape.getChart();
		chart.getSeries().clear();

		ChartSeries series = chart.getSeries().add("Series 1", new String[] { "Category1", "Category2", "Category3" },
				new double[] { 2.7, 3.2, 0.8 });

		ChartDataLabelCollection labels = series.getDataLabels();
		labels.setShowPercentage(true);
		labels.setShowValue(true);
		labels.setShowLeaderLines(false);
		labels.setSeparator(" - ");

		doc.save(dataDir + "Demo.docx");
		// ExEnd:DefaultOptionsForDataLabels
		System.out.println(
				"\nDefault options for data labels of chart series created successfully.\nFile saved at " + dataDir);
	}

}
