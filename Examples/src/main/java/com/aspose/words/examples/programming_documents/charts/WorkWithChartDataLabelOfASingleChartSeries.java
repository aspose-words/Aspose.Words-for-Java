package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class WorkWithChartDataLabelOfASingleChartSeries {

    public static final String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";

    public static void main(String[] args) throws Exception {
        //ExStart:WorkWithChartDataLabelOfASingleChartSeries
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertChart(ChartType.BAR, 432, 252);

        // Get first series.
        ChartSeries series0 = shape.getChart().getSeries().get(0);

        series0.hasDataLabels(true);

        // Set properties.
        series0.getDataLabels().setShowLegendKey(true);

        // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
        // positioned far outside the end of data points. Leader lines create a visual connection between a data label and its
        // corresponding data point.
        series0.getDataLabels().setShowLeaderLines(true);

        series0.getDataLabels().setShowCategoryName(false);
        series0.getDataLabels().setShowPercentage(false);
        series0.getDataLabels().setShowSeriesName(true);
        series0.getDataLabels().setShowValue(true);
        series0.getDataLabels().setSeparator("/");

        series0.getDataLabels().setShowValue(true);

		doc.save(dataDir + "ChartDataLabelOfASingleChartSeries_out.docx");
        //ExEnd:WorkWithChartDataLabelOfASingleChartSeries
    }

}