package com.aspose.words.examples.programming_documents.charts;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.text.SimpleDateFormat;
import java.util.Date;

public class OOXMLCharts {

    public static final String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";

    public static void main(String[] args) throws Exception {

        //ExStart:OOXMLCharts
        // Insert Column chart
        insertColumnChart1();
        insertColumnChart2();

        // Insert Scatter chart
        insertScatterChart();

        // Insert Area chart
        insertAreaChart();

        // Insert Bubble chart
        insertBubbleChart();
        //ExEnd:OOXMLCharts
    }

    public static void insertColumnChart1() throws Exception {
        //ExStart:insertColumnChart1
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data. You can specify different chart types and sizes.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);

        // Chart property of Shape contains all chart related options.
        Chart chart = shape.getChart();

        // Get chart series collection.
        ChartSeriesCollection seriesColl = chart.getSeries();

        // Delete default generated series.
        seriesColl.clear();

        // Create category names array, in this example we have two categories.
        String[] categories = new String[]{"AW Category 1", "AW Category 2"};

        // Adding new series. Please note, data arrays must not be empty and arrays must be the same size.
        seriesColl.add("AW Series 1", categories, new double[]{1, 2});
        seriesColl.add("AW Series 2", categories, new double[]{3, 4});
        seriesColl.add("AW Series 3", categories, new double[]{5, 6});
        seriesColl.add("AW Series 4", categories, new double[]{7, 8});
        seriesColl.add("AW Series 5", categories, new double[]{9, 10});

        doc.save(dataDir + "TestInsertChartColumn1_out.docx");
        //ExEnd:insertColumnChart1
    }

    public static void insertColumnChart2() throws Exception {
        //ExStart:insertColumnChart2
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert Column chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);
        Chart chart = shape.getChart();

        // Use this overload to add series to any type of Bar, Column, Line and Surface charts.
        chart.getSeries().add("AW Series 1", new String[]{"AW Category 1", "AW Category 2"}, new double[]{1, 2});

        doc.save(dataDir + "TestInsertColumnChart2_out.docx");
        //ExEnd:insertColumnChart2
    }

    public static void insertScatterChart() throws Exception {
        //ExStart:insertScatterChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert Scatter chart.
        Shape shape = builder.insertChart(ChartType.SCATTER, 432, 252);
        Chart chart = shape.getChart();

        // Use this overload to add series to any type of Scatter charts.
        chart.getSeries().add("AW Series 1", new double[]{0.7, 1.8, 2.6}, new double[]{2.7, 3.2, 0.8});

        doc.save(dataDir + "TestInsertScatterChart_out.docx");
        //ExEnd:insertScatterChart
    }

    public static void insertAreaChart() throws Exception {
        //ExStart:insertAreaChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert Area chart.
        Shape shape = builder.insertChart(ChartType.AREA, 432, 252);
        Chart chart = shape.getChart();

        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        Date date1 = sdf.parse("01/01/2016");
        Date date2 = sdf.parse("02/02/2016");
        Date date3 = sdf.parse("03/03/2016");
        Date date4 = sdf.parse("04/04/2016");
        Date date5 = sdf.parse("05/05/2016");

        // Use this overload to add series to any type of Area, Radar and Stock charts.
        chart.getSeries().add("AW Series 1", new Date[]{date1, date2, date3, date4, date5}, new double[]{32, 32, 28, 12, 15});

        doc.save(dataDir + "TestInsertAreaChart_out.docx");
        //ExEnd:insertAreaChart
    }

    public static void insertBubbleChart() throws Exception {
        //ExStart:insertBubbleChart
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert Bubble chart.
        Shape shape = builder.insertChart(ChartType.BUBBLE, 432, 252);
        Chart chart = shape.getChart();

        // Use this overload to add series to any type of Bubble charts.
        chart.getSeries().add("AW Series 1", new double[]{0.7, 1.8, 2.6}, new double[]{2.7, 3.2, 0.8}, new double[]{10, 4, 8});

        doc.save(dataDir + "TestInsertBubbleChart_out.docx");
        //ExEnd:insertBubbleChart
    }
}
