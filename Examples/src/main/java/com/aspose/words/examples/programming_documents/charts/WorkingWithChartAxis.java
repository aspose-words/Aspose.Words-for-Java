package com.aspose.words.examples.programming_documents.charts;

/**
 * Created by awaishafeez on 12/19/2017.
 */

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.Date;

public class WorkingWithChartAxis {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getSharedDataDir(OOXMLCharts.class) + "Charts/";

        DefineXYAxisProperties(dataDir);
        SetDateTimeValuesToAxis(dataDir);
        SetNumberFormatForAxis(dataDir);
        SetboundsOfAxis(dataDir);
        SetIntervalUnitBetweenLabelsOnAxis(dataDir);
        HideChartAxis(dataDir);
        TickMultiLineLabelAlignment(dataDir);
    }

    public static void DefineXYAxisProperties(String dataDir) throws Exception {
        // ExStart:DefineXYAxisProperties

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.AREA, 432, 252);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("AW Series 1",
                new Date[]{new Date(2002, 1, 1), new Date(2002, 6, 1), new Date(2015, 7, 1), new Date(2015, 8, 1), new Date(2015, 9, 1)},
                new double[]{640, 320, 280, 120, 150});

        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis.
        xAxis.setCategoryType(AxisCategoryType.CATEGORY);

        // Define X axis properties.
        xAxis.setCrosses(AxisCrosses.CUSTOM);
        xAxis.setCrossesAt(3); // measured in display units of the Y axis (hundreds)
        xAxis.setReverseOrder(true);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);
        xAxis.setTickLabelOffset(200);

        // Define Y axis properties.
        yAxis.setTickLabelPosition(AxisTickLabelPosition.HIGH);
        yAxis.setMajorUnit(100);
        yAxis.setMinorUnit(50);
        yAxis.getDisplayUnit().setUnit(AxisBuiltInUnit.HUNDREDS);
        yAxis.getScaling().setMinimum(new AxisBound(100));
        yAxis.getScaling().setMaximum(new AxisBound(700));

        dataDir = dataDir + "SetAxisProperties_out.docx";
        doc.save(dataDir);
        // ExEnd:DefineXYAxisProperties
        System.out.println("\nProperties of X and Y axis are set successfully.\nFile saved at " + dataDir);
    }

    public static void SetDateTimeValuesToAxis(String dataDir) throws Exception {
        // ExStart:SetDateTimeValuesToAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("AW Series 1",
                new Date[]{new Date(2017, 11, 06), new Date(2017, 11, 9), new Date(2017, 11, 15),
                        new Date(2017, 11, 21), new Date(2017, 11, 25), new Date(2017, 11, 29)},
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2, 5.3}
        );

        // Set X axis bounds.
        ChartAxis xAxis = chart.getAxisX();
        xAxis.getScaling().setMinimum(new AxisBound(new Date(2017, 11, 5).getTime()));
        xAxis.getScaling().setMaximum(new AxisBound(new Date(2017, 12, 3).getTime()));

        // Set major units to a week and minor units to a day.
        xAxis.setMajorUnit(7);
        xAxis.setMinorUnit(1);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);

        dataDir = dataDir + "SetDateTimeValuesToAxis_out.docx";
        doc.save(dataDir);
        // ExEnd:SetDateTimeValuesToAxis
        System.out.println("\nDateTime values are set for chart axis successfully.\nFile saved at " + dataDir);
    }

    public static void SetNumberFormatForAxis(String dataDir) throws Exception {
        // ExStart:SetNumberFormatForAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("AW Series 1",
                new String[]{"Item 1", "Item 2", "Item 3", "Item 4", "Item 5"},
                new double[]{1900000, 850000, 2100000, 600000, 1500000});

        // Set number format.
        chart.getAxisY().getNumberFormat().setFormatCode("#,##0");

        dataDir = dataDir + "FormatAxisNumber_out.docx";
        doc.save(dataDir);
        // ExEnd:SetNumberFormatForAxis
        System.out.println("\nSet number format for axis successfully.\nFile saved at " + dataDir);
    }

    public static void SetboundsOfAxis(String dataDir) throws Exception {
        // ExStart:SetboundsOfAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("AW Series 1",
                new String[]{"Item 1", "Item 2", "Item 3", "Item 4", "Item 5"},
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2});

        chart.getAxisY().getScaling().setMinimum(new AxisBound(0));
        chart.getAxisY().getScaling().setMaximum(new AxisBound(6));

        dataDir = dataDir + "SetboundsOfAxis_out.docx";
        doc.save(dataDir);
        // ExEnd:SetboundsOfAxis
        System.out.println("\nSet Bounds of chart axis successfully.\nFile saved at " + dataDir);
    }

    public static void SetIntervalUnitBetweenLabelsOnAxis(String dataDir) throws Exception {
        // ExStart:SetIntervalUnitBetweenLabelsOnAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("AW Series 1",
                new String[]{"Item 1", "Item 2", "Item 3", "Item 4", "Item 5"},
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2});

        chart.getAxisX().setTickLabelSpacing(2);

        dataDir = dataDir + "SetIntervalUnitBetweenLabelsOnAxis_out.docx";
        doc.save(dataDir);
        // ExEnd:SetIntervalUnitBetweenLabelsOnAxis
        System.out.println("\nSet interval unit between labels on an axis successfully.\nFile saved at " + dataDir);
    }

    public static void HideChartAxis(String dataDir) throws Exception {
        // ExStart:HideChartAxis
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432, 252);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().add("AW Series 1",
                new String[]{"Item 1", "Item 2", "Item 3", "Item 4", "Item 5"},
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2});

        // Hide the Y axis.
        chart.getAxisY().setHidden(true);

        dataDir = dataDir + "HideChartAxis_out.docx";
        doc.save(dataDir);
        // ExEnd:HideChartAxis
        System.out.println("\nY Axis of chart has been hidden successfully.\nFile saved at " + dataDir);
    }

    public static void TickMultiLineLabelAlignment(String dataDir) throws Exception {
        // ExStart:TickMultiLineLabelAlignment
        Document doc = new Document(dataDir + "Document.docx");
        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        ChartAxis axis = shape.getChart().getAxisX();

        //This property has effect only for multi-line labels.
        axis.setTickLabelAlignment(ParagraphAlignment.RIGHT);

        doc.save(dataDir + "Document_out.docx");
        // ExEnd:TickMultiLineLabelAlignment
        System.out.println("\nMulti-Line label for X Axis of chart has been aligned successfully.\nFile saved at " + dataDir);
    }
}
