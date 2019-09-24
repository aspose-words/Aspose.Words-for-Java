package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.*;

public class WorkWithChartSeriesCollectionOfChart {

    public static void main(String[] args) throws Exception {

        //ExStart:WorkWithChartSeriesCollectionOfChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
        // Chart property of Shape contains all chart related options.
        Chart chart = shape.getChart();

        // Get chart series collection.
        ChartSeriesCollection seriesCollection = chart.getSeries();

        // Check series count.
        System.out.println(seriesCollection.getCount());
        //ExEnd:WorkWithChartSeriesCollectionOfChart
    }

}
