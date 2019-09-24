package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class ChartAppearance {

    public static final String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";

    public static void main(String[] args) throws Exception {
        //ExStart:ChartAppearance
        // Working with Charts through Shape.Chart Object
        changeChartAppearanceUsingShapeChartObject();

        // Working with Single ChartSeries Class
        workingWithSingleChartSeries();

        //All single ChartSeries have default ChartDataPoint options, lets change them
        changeDefaultChartDataPointOptions();
        //ExEnd:ChartAppearance
    }

    private static void changeChartAppearanceUsingShapeChartObject() throws Exception {
        //ExStart:changeChartAppearanceUsingShapeChartObject
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);
        Chart chart = shape.getChart();

        // Determines whether the title shall be shown for this chart. Default is true.
        chart.getTitle().setShow(true);

        // Setting chart Title.
        chart.getTitle().setText("Sample Line Chart Title");

        // Determines whether other chart elements shall be allowed to overlap title.
        chart.getTitle().setOverlay(false);

        // Please note if null or empty value is specified as title text, auto generated title will be shown.

        // Determines how legend shall be shown for this chart.
        chart.getLegend().setPosition(LegendPosition.LEFT);
        chart.getLegend().setOverlay(true);

        doc.save(dataDir + "ChartAppearance_out.docx");
        //ExEnd:changeChartAppearanceUsingShapeChartObject
    }

    private static void workingWithSingleChartSeries() throws Exception {
        //ExStart:workingWithSingleChartSeries
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);

        // Get first series.
        ChartSeries series0 = shape.getChart().getSeries().get(0);

        // Get second series.
        ChartSeries series1 = shape.getChart().getSeries().get(1);

        // Change first series name.
        series0.setName("My Name1");

        // Change second series name.
        series1.setName("My Name2");

        // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
        series0.setSmooth(true);
        series1.setSmooth(true);

        doc.save(dataDir + "SingleChartSeries_out.docx");
        //ExEnd:workingWithSingleChartSeries
    }

    private static void changeDefaultChartDataPointOptions() throws Exception {
        //ExStart:changeDefaultChartDataPointOptions
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.LINE, 432, 252);

        // Get first series.
        ChartSeries series0 = shape.getChart().getSeries().get(0);

        // Get second series.
        ChartSeries series1 = shape.getChart().getSeries().get(1);

        // Specifies whether by default the parent element shall inverts its colors if the value is negative.
        series0.setInvertIfNegative(true);

        // Set default marker symbol and size.
        series0.getMarker().setSymbol(MarkerSymbol.CIRCLE);
        series0.getMarker().setSize(15);

        series1.getMarker().setSymbol(MarkerSymbol.STAR);
        series1.getMarker().setSize(10);

        doc.save(dataDir + "ChartDataPoints_out.docx");//
        //ExEnd:changeDefaultChartDataPointOptions
    }

}
