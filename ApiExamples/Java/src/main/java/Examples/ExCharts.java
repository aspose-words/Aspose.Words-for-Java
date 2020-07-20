package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

@Test
public class ExCharts extends ApiExampleBase {
    @Test
    public void chartTitle() throws Exception {
        //ExStart
        //ExFor:Chart
        //ExFor:Chart.Title
        //ExFor:ChartTitle
        //ExFor:ChartTitle.Overlay
        //ExFor:ChartTitle.Show
        //ExFor:ChartTitle.Text
        //ExSummary:Shows how to insert a chart and change its title.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a bar chart
        Shape chartShape = builder.insertChart(ChartType.BAR, 400.0, 300.0);

        // Get the chart object from the containing shape
        Chart chart = chartShape.getChart();

        // Set the title text, which appears at the top center of the chart and modify its appearance
        ChartTitle title = chart.getTitle();
        title.setText("MyChart");
        title.setOverlay(true);
        title.setShow(true);

        doc.save(getArtifactsDir() + "Charts.ChartTitle.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.ChartTitle.docx");
        chartShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(ShapeType.NON_PRIMITIVE, chartShape.getShapeType());
        Assert.assertTrue(chartShape.hasChart());

        title = chartShape.getChart().getTitle();

        Assert.assertEquals("MyChart", title.getText());
        Assert.assertTrue(title.getOverlay());
        Assert.assertTrue(title.getShow());
    }

    @Test
    public void defineNumberFormatForDataLabels() throws Exception {
        //ExStart
        //ExFor:ChartDataLabelCollection.NumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExSummary:Shows how to set number format for the data labels of the entire series.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        // Delete default generated series
        shape.getChart().getSeries().clear();

        ChartSeries series =
                shape.getChart().getSeries().add("Aspose Test Series", new String[]{"Word", "PDF", "Excel"}, new double[]{2.5, 1.5, 3.5});

        ChartDataLabelCollection dataLabels = series.getDataLabels();
        // Display chart values in the data labels, by default it is false
        dataLabels.setShowValue(true);
        // Set currency format for the data labels of the entire series
        dataLabels.getNumberFormat().setFormatCode("\"$\"#,##0.00");

        doc.save(getArtifactsDir() + "Charts.DefineNumberFormatForDataLabels.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.DefineNumberFormatForDataLabels.docx");
        series = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0);

        Assert.assertEquals("", series.getDataLabels().getNumberFormat().getFormatCode());
    }

    @Test
    public void dataArraysWrongSize() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        // Create category names array, second category will be null
        String[] categories = {"Cat1", null, "Cat3", "Cat4", "Cat5", null};

        // Adding new series with empty (double.NaN) values
        seriesColl.add("AW Series 1", categories, new double[]{1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0});
        seriesColl.add("AW Series 2", categories, new double[]{2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0});

        Assert.assertThrows(IllegalArgumentException.class, () -> seriesColl.add("AW Series 3", categories, new double[]{Double.NaN, 4.0, 5.0, Double.NaN, Double.NaN}));
        Assert.assertThrows(IllegalArgumentException.class, () -> seriesColl.add("AW Series 4", categories, new double[]{Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN}));
    }

    @Test
    public void emptyValuesInChartData() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        // Create category names array, second category will be null
        String[] categories = {"Cat1", null, "Cat3", "Cat4", "Cat5", null};

        // Adding new series with empty (double.NaN) values
        seriesColl.add("AW Series 1", categories, new double[]{1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0});
        seriesColl.add("AW Series 2", categories, new double[]{2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0});
        seriesColl.add("AW Series 3", categories, new double[]{Double.NaN, 4.0, 5.0, Double.NaN, 7.0, 8.0});
        seriesColl.add("AW Series 4", categories,
                new double[]{Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN, 9.0});

        doc.save(getArtifactsDir() + "Charts.EmptyValuesInChartData.docx");
    }

    @Test
    public void axisProperties() throws Exception {
        //ExStart
        //ExFor:ChartAxis
        //ExFor:ChartAxis.CategoryType
        //ExFor:ChartAxis.Crosses
        //ExFor:ChartAxis.ReverseOrder
        //ExFor:ChartAxis.MajorTickMark
        //ExFor:ChartAxis.MinorTickMark
        //ExFor:ChartAxis.MajorUnit
        //ExFor:ChartAxis.MinorUnit
        //ExFor:ChartAxis.TickLabelOffset
        //ExFor:ChartAxis.TickLabelPosition
        //ExFor:ChartAxis.TickLabelSpacingIsAuto
        //ExFor:ChartAxis.TickMarkSpacing
        //ExFor:AxisCategoryType
        //ExFor:AxisCrosses
        //ExFor:Chart.AxisX
        //ExFor:Chart.AxisY
        //ExFor:Chart.AxisZ
        //ExSummary:Shows how to insert chart using the axis options for detailed configuration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data
        chart.getSeries().clear();
        chart.getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{640.0, 320.0, 280.0, 120.0, 150.0});

        // Get chart axes
        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // For 2D charts like the one we made, the Z axis is null
        Assert.assertNull(chart.getAxisZ());

        // Set X-axis options
        xAxis.setCategoryType(AxisCategoryType.CATEGORY);
        xAxis.setCrosses(AxisCrosses.MINIMUM);
        xAxis.setReverseOrder(false);
        xAxis.setMajorTickMark(AxisTickMark.INSIDE);
        xAxis.setMinorTickMark(AxisTickMark.CROSS);
        xAxis.setMajorUnit(10.0);
        xAxis.setMinorUnit(15.0);
        xAxis.setTickLabelOffset(50);
        xAxis.setTickLabelPosition(AxisTickLabelPosition.LOW);
        xAxis.setTickLabelSpacingIsAuto(false);
        xAxis.setTickMarkSpacing(1);

        // Set Y-axis options
        yAxis.setCategoryType(AxisCategoryType.AUTOMATIC);
        yAxis.setCrosses(AxisCrosses.MAXIMUM);
        yAxis.setReverseOrder(true);
        yAxis.setMajorTickMark(AxisTickMark.INSIDE);
        yAxis.setMinorTickMark(AxisTickMark.CROSS);
        yAxis.setMajorUnit(100.0);
        yAxis.setMinorUnit(20.0);
        yAxis.setTickLabelPosition(AxisTickLabelPosition.NEXT_TO_AXIS);

        doc.save(getArtifactsDir() + "Charts.AxisProperties.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisProperties.docx");
        chart = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(AxisCategoryType.CATEGORY, chart.getAxisX().getCategoryType());
        Assert.assertEquals(AxisCrosses.MINIMUM, chart.getAxisX().getCrosses());
        Assert.assertFalse(chart.getAxisX().getReverseOrder());
        Assert.assertEquals(AxisTickMark.INSIDE, chart.getAxisX().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.CROSS, chart.getAxisX().getMinorTickMark());
        Assert.assertEquals(1.0d, chart.getAxisX().getMajorUnit());
        Assert.assertEquals(0.5d, chart.getAxisX().getMinorUnit());
        Assert.assertEquals(50, chart.getAxisX().getTickLabelOffset());
        Assert.assertEquals(AxisTickLabelPosition.LOW, chart.getAxisX().getTickLabelPosition());
        Assert.assertFalse(chart.getAxisX().getTickLabelSpacingIsAuto());
        Assert.assertEquals(1, chart.getAxisX().getTickMarkSpacing());

        Assert.assertEquals(AxisCategoryType.CATEGORY, chart.getAxisY().getCategoryType());
        Assert.assertEquals(AxisCrosses.MAXIMUM, chart.getAxisY().getCrosses());
        Assert.assertTrue(chart.getAxisY().getReverseOrder());
        Assert.assertEquals(AxisTickMark.INSIDE, chart.getAxisY().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.CROSS, chart.getAxisY().getMinorTickMark());
        Assert.assertEquals(100.0d, chart.getAxisY().getMajorUnit());
        Assert.assertEquals(20.0d, chart.getAxisY().getMinorUnit());
        Assert.assertEquals(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisY().getTickLabelPosition());
    }

    @Test
    public void dateTimeValues() throws Exception {
        //ExStart
        //ExFor:AxisBound
        //ExFor:AxisBound.#ctor(Double)
        //ExFor:AxisBound.#ctor(DateTime)
        //ExFor:AxisScaling.Minimum
        //ExFor:AxisScaling.Maximum
        //ExFor:ChartAxis.Scaling
        //ExFor:AxisTickMark
        //ExFor:AxisTickLabelPosition
        //ExFor:AxisTimeUnit
        //ExFor:ChartAxis.BaseTimeUnit
        //ExSummary:Shows how to insert chart with date/time values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data
        chart.getSeries().clear();

        Calendar cal = Calendar.getInstance();

        // Fill data
        chart.getSeries().add("Aspose Test Series",
                new Date[]
                        {
                                DocumentHelper.createDate(2017, 11, 6), DocumentHelper.createDate(2017, 11, 9), DocumentHelper.createDate(2017, 11, 15),
                                DocumentHelper.createDate(2017, 11, 21), DocumentHelper.createDate(2017, 11, 25), DocumentHelper.createDate(2017, 11, 29)
                        },
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2, 5.3});

        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // Set X axis bounds
        xAxis.getScaling().setMinimum(new AxisBound(DocumentHelper.createDate(2017, 11, 5)));
        xAxis.getScaling().setMaximum(new AxisBound(DocumentHelper.createDate(2017, 12, 3)));

        // Set major units to a week and minor units to a day
        xAxis.setBaseTimeUnit(AxisTimeUnit.DAYS);
        xAxis.setMajorUnit(7.0);
        xAxis.setMinorUnit(1.0);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);

        // Define Y axis properties
        yAxis.setTickLabelPosition(AxisTickLabelPosition.HIGH);
        yAxis.setMajorUnit(100.0);
        yAxis.setMinorUnit(50.0);
        yAxis.getDisplayUnit().setUnit(AxisBuiltInUnit.HUNDREDS);
        yAxis.getScaling().setMinimum(new AxisBound(100.0));
        yAxis.getScaling().setMaximum(new AxisBound(700.0));

        doc.save(getArtifactsDir() + "Charts.DateTimeValues.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.DateTimeValues.docx");
        chart = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(AxisTimeUnit.DAYS, chart.getAxisX().getBaseTimeUnit());
        Assert.assertEquals(7.0d, chart.getAxisX().getMajorUnit());
        Assert.assertEquals(1.0d, chart.getAxisX().getMinorUnit());
        Assert.assertEquals(AxisTickMark.CROSS, chart.getAxisX().getMajorTickMark());
        Assert.assertEquals(AxisTickMark.OUTSIDE, chart.getAxisX().getMinorTickMark());

        Assert.assertEquals(AxisTickLabelPosition.HIGH, chart.getAxisY().getTickLabelPosition());
        Assert.assertEquals(100.0d, chart.getAxisY().getMajorUnit());
        Assert.assertEquals(50.0d, chart.getAxisY().getMinorUnit());
        Assert.assertEquals(AxisBuiltInUnit.HUNDREDS, chart.getAxisY().getDisplayUnit().getUnit());
        Assert.assertEquals(new AxisBound(100.0), chart.getAxisY().getScaling().getMinimum());
        Assert.assertEquals(new AxisBound(700.0), chart.getAxisY().getScaling().getMaximum());
    }

    @Test
    public void hideChartAxis() throws Exception {
        //ExStart
        //ExFor:ChartAxis.Hidden
        //ExSummary:Shows how to hide chart axes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Hide both the X and Y axes
        chart.getAxisX().setHidden(true);
        chart.getAxisY().setHidden(true);

        // Clear demo data
        chart.getSeries().clear();
        chart.getSeries().add("AW Series 1",
                new String[]{"Item 1", "Item 2", "Item 3", "Item 4", "Item 5"},
                new double[]{1.2, 0.3, 2.1, 2.9, 4.2});

        doc.save(getArtifactsDir() + "Charts.HideChartAxis.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.HideChartAxis.docx");
        chart = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(chart.getAxisX().getHidden(), true);
        Assert.assertEquals(chart.getAxisY().getHidden(), true);
    }

    @Test
    public void setNumberFormatToChartAxis() throws Exception {
        //ExStart
        //ExFor:ChartAxis.NumberFormat
        //ExFor:Charts.ChartNumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExFor:Charts.ChartNumberFormat.IsLinkedToSource
        //ExSummary:Shows how to set formatting for chart values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data and replace it with a new custom chart series
        chart.getSeries().clear();
        chart.getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0});

        // Set number format
        chart.getAxisY().getNumberFormat().setFormatCode("#,##0");

        // Set this to override the above value and draw the number format from the source cell
        Assert.assertFalse(chart.getAxisY().getNumberFormat().isLinkedToSource());

        doc.save(getArtifactsDir() + "Charts.SetNumberFormatToChartAxis.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.SetNumberFormatToChartAxis.docx");
        chart = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals("#,##0", chart.getAxisY().getNumberFormat().getFormatCode());
    }

    @Test(dataProvider = "testDisplayChartsWithConversionDataProvider")
    public void testDisplayChartsWithConversion(int chartType) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(chartType, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0});

        doc.save(getArtifactsDir() + "Charts.TestDisplayChartsWithConversion.docx");
        doc.save(getArtifactsDir() + "Charts.TestDisplayChartsWithConversion.pdf");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "testDisplayChartsWithConversionDataProvider")
    public static Object[][] testDisplayChartsWithConversionDataProvider() {
        return new Object[][]
                {
                        {ChartType.COLUMN},
                        {ChartType.LINE},
                        {ChartType.PIE},
                        {ChartType.BAR},
                        {ChartType.AREA},
                };
    }

    @Test
    public void surface3DChart() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart
        Shape shape = builder.insertChart(ChartType.SURFACE_3_D, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series 1",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0});

        chart.getSeries().add("Aspose Test Series 2",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{900000.0, 50000.0, 1100000.0, 400000.0, 2500000.0});

        chart.getSeries().add("Aspose Test Series 3",
                new String[]{"Word", "PDF", "Excel", "GoogleDocs", "Note"},
                new double[]{500000.0, 820000.0, 1500000.0, 400000.0, 100000.0});

        doc.save(getArtifactsDir() + "Charts.SurfaceChart.docx");
        doc.save(getArtifactsDir() + "Charts.SurfaceChart.pdf");
    }

    @Test
    public void chartDataLabelCollection() throws Exception {
        //ExStart
        //ExFor:ChartDataLabelCollection.ShowBubbleSize
        //ExFor:ChartDataLabelCollection.ShowCategoryName
        //ExFor:ChartDataLabelCollection.ShowSeriesName
        //ExFor:ChartDataLabelCollection.Separator
        //ExFor:ChartDataLabelCollection.ShowLeaderLines
        //ExFor:ChartDataLabelCollection.ShowLegendKey
        //ExFor:ChartDataLabelCollection.ShowPercentage
        //ExFor:ChartDataLabelCollection.ShowValue
        //ExSummary:Shows how to set default values for the data labels.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert bubble chart
        Chart chart = builder.insertChart(ChartType.BUBBLE, 432.0, 252.0).getChart();
        // Clear demo data
        chart.getSeries().clear();

        ChartSeries bubbleChartSeries = chart.getSeries().add("Aspose Test Series",
                new double[]{2.9, 3.5, 1.1, 4.0, 4.0},
                new double[]{1.9, 8.5, 2.1, 6.0, 1.5},
                new double[]{9.0, 4.5, 2.5, 8.0, 5.0});

        // Set default values for the bubble chart data labels
        ChartDataLabelCollection bubbleChartDataLabels = bubbleChartSeries.getDataLabels();
        bubbleChartDataLabels.setShowBubbleSize(true);
        bubbleChartDataLabels.setShowCategoryName(true);
        bubbleChartDataLabels.setShowSeriesName(true);
        bubbleChartDataLabels.setSeparator(" - ");

        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert pie chart
        Shape shapeWithPieChart = builder.insertChart(ChartType.PIE, 432.0, 252.0);
        // Clear demo data
        shapeWithPieChart.getChart().getSeries().clear();

        ChartSeries pieChartSeries = shapeWithPieChart.getChart().getSeries().add("Aspose Test Series",
                new String[]{"Word", "PDF", "Excel"},
                new double[]{2.7, 3.2, 0.8});

        // Set default values for the pie chart data labels
        ChartDataLabelCollection pieChartDataLabels = pieChartSeries.getDataLabels();
        pieChartDataLabels.setShowLeaderLines(true);
        pieChartDataLabels.setShowLegendKey(true);
        pieChartDataLabels.setShowPercentage(true);
        pieChartDataLabels.setShowValue(true);

        doc.save(getArtifactsDir() + "Charts.ChartDataLabelCollection.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.ChartDataLabelCollection.docx");
        bubbleChartDataLabels = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0).getDataLabels();

        Assert.assertFalse(bubbleChartDataLabels.getShowBubbleSize());
        Assert.assertFalse(bubbleChartDataLabels.getShowCategoryName());
        Assert.assertFalse(bubbleChartDataLabels.getShowSeriesName());
        Assert.assertEquals(",", bubbleChartDataLabels.getSeparator());

        pieChartDataLabels = ((Shape) doc.getChild(NodeType.SHAPE, 1, true)).getChart().getSeries().get(0).getDataLabels();

        Assert.assertFalse(pieChartDataLabels.getShowLeaderLines());
        Assert.assertFalse(pieChartDataLabels.getShowLegendKey());
        Assert.assertFalse(pieChartDataLabels.getShowPercentage());
        Assert.assertFalse(pieChartDataLabels.getShowValue());
    }

    //ExStart
    //ExFor:ChartSeries
    //ExFor:ChartSeries.DataLabels
    //ExFor:ChartSeries.DataPoints
    //ExFor:ChartSeries.Name
    //ExFor:ChartDataLabel
    //ExFor:ChartDataLabel.Index
    //ExFor:ChartDataLabel.IsVisible
    //ExFor:ChartDataLabel.NumberFormat
    //ExFor:ChartDataLabel.Separator
    //ExFor:ChartDataLabel.ShowCategoryName
    //ExFor:ChartDataLabel.ShowDataLabelsRange
    //ExFor:ChartDataLabel.ShowLeaderLines
    //ExFor:ChartDataLabel.ShowLegendKey
    //ExFor:ChartDataLabel.ShowPercentage
    //ExFor:ChartDataLabel.ShowSeriesName
    //ExFor:ChartDataLabel.ShowValue
    //ExFor:ChartDataLabel.IsHidden
    //ExFor:ChartDataLabelCollection
    //ExFor:ChartDataLabelCollection.Add(System.Int32)
    //ExFor:ChartDataLabelCollection.Clear
    //ExFor:ChartDataLabelCollection.Count
    //ExFor:ChartDataLabelCollection.GetEnumerator
    //ExFor:ChartDataLabelCollection.Item(System.Int32)
    //ExFor:ChartDataLabelCollection.RemoveAt(System.Int32)
    //ExSummary:Shows how to apply labels to data points in a chart.
    @Test //ExSkip
    public void chartDataLabels() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a bar chart
        Shape chartShape = builder.insertChart(ChartType.LINE, 400.0, 300.0);

        // Get the chart object from the containing shape
        Chart chart = chartShape.getChart();

        // The chart already contains demo data comprised of 3 series each with 4 categories
        Assert.assertEquals(chart.getSeries().getCount(), 3);
        Assert.assertEquals(chart.getSeries().get(0).getName(), "Series 1");

        // Apply data labels to every series in the graph
        for (ChartSeries series : chart.getSeries()) {
            applyDataLabels(series, 4, "000.0", ", ");
            Assert.assertEquals(series.getDataLabels().getCount(), 4);
        }

        // Get the enumerator for a data label collection
        Iterator<ChartDataLabel> enumerator = chart.getSeries().get(0).getDataLabels().iterator();

        // And use it to go over all the data labels in one series and change their separator
        while (enumerator.hasNext()) {
            Assert.assertEquals(enumerator.next().getSeparator(), ", ");
            enumerator.next().setSeparator(" & ");
        }

        // If the chart looks too busy, we can remove data labels one by one
        chart.getSeries().get(1).getDataLabels().get(2).clearFormat();

        // We can also clear an entire data label collection for one whole series
        chart.getSeries().get(2).getDataLabels().clearFormat();

        doc.save(getArtifactsDir() + "Charts.ChartDataLabels.docx");
    }

    /// <summary>
    /// Apply uniform data labels with custom number format and separator to a number (determined by labelsCount) of data points in a series.
    /// </summary>
    private static void applyDataLabels(ChartSeries series, int labelsCount, String numberFormat, String separator) {
        for (int i = 0; i < labelsCount; i++) {
            series.hasDataLabels(true);
            ChartDataLabel label = series.getDataLabels().get(i);
            Assert.assertFalse(label.isVisible());

            // Edit the appearance of the new data label
            label.setShowCategoryName(true);
            label.setShowSeriesName(true);
            label.setShowValue(true);
            label.setShowLeaderLines(true);
            label.setShowLegendKey(true);
            label.setShowPercentage(false);
            label.isHidden(false);
            Assert.assertFalse(label.getShowDataLabelsRange());

            // Apply number format and separator
            label.getNumberFormat().setFormatCode(numberFormat);
            label.setSeparator(separator);

            // The label automatically becomes visible
            Assert.assertTrue(label.isVisible());
            Assert.assertFalse(label.isHidden());
        }
    }
    //ExEnd

    //ExStart
    //ExFor:ChartSeries.Smooth
    //ExFor:ChartDataPoint
    //ExFor:ChartDataPoint.Index
    //ExFor:ChartDataPointCollection
    //ExFor:ChartDataPointCollection.Add(System.Int32)
    //ExFor:ChartDataPointCollection.Clear
    //ExFor:ChartDataPointCollection.Count
    //ExFor:ChartDataPointCollection.GetEnumerator
    //ExFor:ChartDataPointCollection.Item(System.Int32)
    //ExFor:ChartDataPointCollection.RemoveAt(System.Int32)
    //ExFor:ChartMarker
    //ExFor:ChartMarker.Size
    //ExFor:ChartMarker.Symbol
    //ExFor:IChartDataPoint
    //ExFor:IChartDataPoint.InvertIfNegative
    //ExFor:IChartDataPoint.Marker
    //ExFor:MarkerSymbol
    //ExSummary:Shows how to customize chart data points.
    @Test
    public void chartDataPoint() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line chart, which will have default data that we will use
        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 350.0);
        Chart chart = shape.getChart();

        // Apply diamond-shaped data points to the line of the first series
        for (ChartSeries series : chart.getSeries()) {
            applyDataPoints(series, 4, MarkerSymbol.DIAMOND, 15);
        }

        // We can further decorate a series line by smoothing it
        chart.getSeries().get(0).setSmooth(true);

        // Get the enumerator for the data point collection from one series
        Iterator<ChartDataPoint> enumerator = chart.getSeries().get(0).getDataPoints().iterator();

        // And use it to go over all the data labels in one series and change their separator
        while (enumerator.hasNext()) {
            Assert.assertFalse(enumerator.next().getInvertIfNegative());
        }

        // If the chart looks too busy, we can remove data points one by one
        chart.getSeries().get(1).getDataPoints().removeAt(2);

        // We can also clear an entire data point collection for one whole series
        chart.getSeries().get(2).getDataPoints().clear();

        doc.save(getArtifactsDir() + "Charts.ChartDataPoint.docx");
    }

    /// <summary>
    /// Applies a number of data points to a series
    /// </summary>
    private static void applyDataPoints(ChartSeries series, int dataPointsCount, int markerSymbol, int dataPointSize) {
        for (int i = 0; i < dataPointsCount; i++) {
            ChartDataPoint point = series.getDataPoints().add(i);
            point.getMarker().setSymbol(markerSymbol);
            point.getMarker().setSize(dataPointSize);

            Assert.assertEquals(point.getIndex(), i);
        }
    }
    //ExEnd

    @Test
    public void pieChartExplosion() throws Exception {
        //ExStart
        //ExFor:IChartDataPoint.Explosion
        //ExSummary:Shows how to manipulate the position of the portions of a pie chart.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape shape = builder.insertChart(ChartType.PIE, 500.0, 350.0);
        Chart chart = shape.getChart();

        // In a pie chart, the portions are the data points, which cannot have markers or sizes applied to them
        // However, we can set this variable to move any individual "slice" away from the center of the chart
        ChartDataPoint cdp = chart.getSeries().get(0).getDataPoints().add(0);
        cdp.setExplosion(10);

        cdp = chart.getSeries().get(0).getDataPoints().add(1);
        cdp.setExplosion(40);

        doc.save(getArtifactsDir() + "Charts.PieChartExplosion.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.PieChartExplosion.docx");
        ChartSeries series = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0);

        Assert.assertEquals(10, series.getDataPoints().get(0).getExplosion());
        Assert.assertEquals(40, series.getDataPoints().get(1).getExplosion());
    }

    @Test
    public void bubble3D() throws Exception {
        //ExStart
        //ExFor:ChartDataLabel.ShowBubbleSize
        //ExFor:IChartDataPoint.Bubble3D
        //ExSummary:Shows how to use 3D effects with bubble charts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bubble chart with 3D effects on each bubble
        Shape shape = builder.insertChart(ChartType.BUBBLE_3_D, 500.0, 350.0);
        Chart chart = shape.getChart();

        Assert.assertTrue(chart.getSeries().get(0).getBubble3D());

        // Apply a data label to each bubble that displays the size of its bubble
        for (int i = 0; i < 3; i++) {
            chart.getSeries().get(0).hasDataLabels(true);
            ChartDataLabel cdl = chart.getSeries().get(0).getDataLabels().get(i);
            cdl.setShowBubbleSize(true);
        }

        doc.save(getArtifactsDir() + "Charts.Bubble3D.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.Bubble3D.docx");
        ChartSeries series = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart().getSeries().get(0);

        for (int i = 0; i < 3; i++) {
            Assert.assertTrue(series.getDataLabels().get(i).getShowBubbleSize());
        }
    }

    //ExStart
    //ExFor:ChartAxis.Type
    //ExFor:ChartAxisType
    //ExFor:ChartType
    //ExFor:Chart.Series
    //ExFor:ChartSeriesCollection.Add(String,DateTime[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,Double[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,Double[],Double[],Double[])
    //ExFor:ChartSeriesCollection.Add(String,String[],Double[])
    //ExSummary:Shows how to pick an appropriate graph type for a chart series.
    @Test //ExSkip
    public void chartSeriesCollection() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // There are 4 ways of populating a chart's series collection
        // 1: Each series has a string array of categories, each with a corresponding data value
        // Some of the other possible applications are bar, column, line and surface charts
        Chart chart = appendChart(builder, ChartType.COLUMN, 300.0, 300.0);

        // Create and name 3 categories with a string array
        String[] categories = {"Category 1", "Category 2", "Category 3"};

        // Create 2 series of data, each with one point for every category
        // This will generate a column graph with 3 clusters of 2 bars
        chart.getSeries().add("Series 1", categories, new double[]{76.6, 82.1, 91.6});
        chart.getSeries().add("Series 2", categories, new double[]{64.2, 79.5, 94.0});

        // Categories are distributed along the X-axis while values are distributed along the Y-axis
        Assert.assertEquals(chart.getAxisX().getType(), ChartAxisType.CATEGORY);
        Assert.assertEquals(chart.getAxisY().getType(), ChartAxisType.VALUE);

        // 2: Each series will have a collection of dates with a corresponding value for each date
        // Area, radar and stock charts are some of the appropriate chart types for this
        chart = appendChart(builder, ChartType.AREA, 300.0, 300.0);

        // Create a collection of dates to serve as categories
        Date[] dates = {DocumentHelper.createDate(2014, 3, 31),
                DocumentHelper.createDate(2017, 1, 23),
                DocumentHelper.createDate(2017, 6, 18),
                DocumentHelper.createDate(2019, 11, 22),
                DocumentHelper.createDate(2020, 9, 7)
        };

        // Add one series with one point for each date
        // Our sporadic dates will be distributed along the X-axis in a linear fashion 
        chart.getSeries().add("Series 1", dates, new double[]{15.8, 21.5, 22.9, 28.7, 33.1});

        // 3: Each series will take two data arrays
        // Appropriate for scatter plots
        chart = appendChart(builder, ChartType.SCATTER, 300.0, 300.0);

        // In each series, the first array contains the X-coordinates and the second contains respective Y-coordinates of points
        chart.getSeries().add("Series 1", new double[]{3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6}, new double[]{3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9});
        chart.getSeries().add("Series 2", new double[]{2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3}, new double[]{7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6});

        // Both axes are value axes in this case
        Assert.assertEquals(chart.getAxisX().getType(), ChartAxisType.VALUE);
        Assert.assertEquals(chart.getAxisY().getType(), ChartAxisType.VALUE);

        // 4: Each series will be built from three data arrays, used for bubble charts
        chart = appendChart(builder, ChartType.BUBBLE, 300.0, 300.0);

        // The first two arrays contain X/Y coordinates like above and the third determines the thickness of each point
        chart.getSeries().add("Series 1", new double[]{1.1, 5.0, 9.8}, new double[]{1.2, 4.9, 9.9}, new double[]{2.0, 4.0, 8.0});

        doc.save(getArtifactsDir() + "Charts.ChartSeriesCollection.docx");
    }

    /// <summary>
    /// Get the DocumentBuilder to insert a chart of a specified ChartType, width and height and clean out its default data
    /// </summary>
    private static Chart appendChart(DocumentBuilder builder, /*ChartType*/int chartType, double width, double height) throws Exception {
        Shape chartShape = builder.insertChart(chartType, width, height);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();

        Assert.assertEquals(chart.getSeries().getCount(), 0);

        return chart;
    }
    //ExEnd

    @Test
    public void chartSeriesCollectionModify() throws Exception {
        //ExStart
        //ExFor:ChartSeriesCollection
        //ExFor:ChartSeriesCollection.Clear
        //ExFor:ChartSeriesCollection.Count
        //ExFor:ChartSeriesCollection.GetEnumerator
        //ExFor:ChartSeriesCollection.Item(Int32)
        //ExFor:ChartSeriesCollection.RemoveAt(Int32)
        //ExSummary:Shows how to work with a chart's data collection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a bar chart
        Shape chartShape = builder.insertChart(ChartType.COLUMN, 400.0, 300.0);
        Chart chart = chartShape.getChart();

        // All charts come with demo data
        // This column chart currently has 3 series with 4 categories, which means 4 clusters, 3 columns in each
        ChartSeriesCollection chartData = chart.getSeries();
        Assert.assertEquals(3, chartData.getCount()); //ExSkip

        // Iterate through the series with an enumerator and print their names
        Iterator<ChartSeries> enumerator = chart.getSeries().iterator();

        // And use it to go over all the data labels in one series and change their separator
        while (enumerator.hasNext()) {
            System.out.println(enumerator.next().getName());
        }


        // We can add new data by adding a new series to the collection, with categories and data
        // We will match the existing category/series names in the demo data and add a 4th column to each column cluster
        String[] categories = {"Category 1", "Category 2", "Category 3", "Category 4"};
        chart.getSeries().add("Series 4", categories, new double[]{4.4, 7.0, 3.5, 2.1});
        Assert.assertEquals(4, chartData.getCount()); //ExSkip
        Assert.assertEquals("Series 4", chartData.get(3).getName()); //ExSkip

        // We can remove series by index
        chartData.removeAt(2);
        Assert.assertEquals(3, chartData.getCount()); //ExSkip
        Assert.assertEquals("Series 4", chartData.get(2).getName()); //ExSkip

        // We can also remove out all the series
        // This leaves us with an empty graph and is a convenient way of wiping out demo data
        chartData.clear();
        Assert.assertEquals(0, chartData.getCount()); //ExSkip
        //ExEnd
    }

    @Test
    public void axisScaling() throws Exception {
        //ExStart
        //ExFor:AxisScaleType
        //ExFor:AxisScaling
        //ExFor:AxisScaling.LogBase
        //ExFor:AxisScaling.Type
        //ExSummary:Shows how to set up logarithmic axis scaling.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart and clear its default data series
        Shape chartShape = builder.insertChart(ChartType.SCATTER, 450.0, 300.0);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();

        // Insert a series with X/Y coordinates for 5 points
        chart.getSeries().add("Series 1", new double[]{1.0, 2.0, 3.0, 4.0, 5.0}, new double[]{1.0, 20.0, 400.0, 8000.0, 160000.0});

        // The scaling of the X axis is linear by default, which means it will display "0, 1, 2, 3..."
        // Linear axis scaling is suitable for our X-values, but our Y-values call for a logarithmic scale to be represented accurately on a graph 
        // We can set the scaling of the Y-axis to Logarithmic with a base of 20
        // The Y-axis will now display "1, 20, 400, 8000...", which is ideal for accurate representation of this set of Y-values
        chart.getAxisY().getScaling().setType(AxisScaleType.LOGARITHMIC);
        chart.getAxisY().getScaling().setLogBase(20.0);

        doc.save(getArtifactsDir() + "Charts.AxisScaling.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisScaling.docx");
        chart = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertEquals(AxisScaleType.LINEAR, chart.getAxisX().getScaling().getType());
        Assert.assertEquals(AxisScaleType.LOGARITHMIC, chart.getAxisY().getScaling().getType());
        Assert.assertEquals(20.0d, chart.getAxisY().getScaling().getLogBase());
    }

    @Test
    public void axisBound() throws Exception {
        //ExStart
        //ExFor:AxisBound.#ctor
        //ExFor:AxisBound.IsAuto
        //ExFor:AxisBound.Value
        //ExFor:AxisBound.ValueAsDate
        //ExSummary:Shows how to set custom axis bounds.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart, remove default data and populate it with data from a ChartSeries
        Shape chartShape = builder.insertChart(ChartType.SCATTER, 450.0, 300.0);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();
        chart.getSeries().add("Series 1", new double[]{1.1, 5.4, 7.9, 3.5, 2.1, 9.7}, new double[]{2.1, 0.3, 0.6, 3.3, 1.4, 1.9});

        // By default, the axis bounds are automatically defined so all the series data within the table is included
        Assert.assertTrue(chart.getAxisX().getScaling().getMinimum().isAuto());

        // If we wish to set our own scale bounds, we need to replace them with new ones
        // Both the axis rulers will go from 0 to 10
        chart.getAxisX().getScaling().setMinimum(new AxisBound(0.0));
        chart.getAxisX().getScaling().setMaximum(new AxisBound(10.0));
        chart.getAxisY().getScaling().setMinimum(new AxisBound(0.0));
        chart.getAxisY().getScaling().setMaximum(new AxisBound(10.0));

        // These are custom and not defined automatically
        Assert.assertFalse(chart.getAxisX().getScaling().getMinimum().isAuto());
        Assert.assertFalse(chart.getAxisY().getScaling().getMinimum().isAuto());

        // Create a line graph
        chartShape = builder.insertChart(ChartType.LINE, 450.0, 300.0);
        chart = chartShape.getChart();
        chart.getSeries().clear();

        // Create a collection of dates, which will make up the X axis
        Date[] dates = {DocumentHelper.createDate(1973, 5, 11),
                DocumentHelper.createDate(1981, 2, 4),
                DocumentHelper.createDate(1985, 9, 23),
                DocumentHelper.createDate(1989, 6, 28),
                DocumentHelper.createDate(1994, 12, 15)
        };

        // Assign a Y-value for each date 
        chart.getSeries().add("Series 1", dates, new double[]{3.0, 4.7, 5.9, 7.1, 8.9});

        // These particular bounds will cut off categories from before 1980 and from 1990 and onwards
        // This narrows the amount of categories and values in the viewport from 5 to 3
        // Note that the graph still contains the out-of-range data because we can see the line tend towards it
        chart.getAxisX().getScaling().setMinimum(new AxisBound(DocumentHelper.createDate(1980, 1, 1)));
        chart.getAxisX().getScaling().setMaximum(new AxisBound(DocumentHelper.createDate(1990, 1, 1)));

        doc.save(getArtifactsDir() + "Charts.AxisBound.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisBound.docx");
        chart = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart();

        Assert.assertFalse(chart.getAxisX().getScaling().getMinimum().isAuto());
        Assert.assertEquals(0.0d, chart.getAxisX().getScaling().getMinimum().getValue());
        Assert.assertEquals(10.0d, chart.getAxisX().getScaling().getMaximum().getValue());

        Assert.assertFalse(chart.getAxisY().getScaling().getMinimum().isAuto());
        Assert.assertEquals(0.0d, chart.getAxisY().getScaling().getMinimum().getValue());
        Assert.assertEquals(10.0d, chart.getAxisY().getScaling().getMaximum().getValue());

        chart = ((Shape) doc.getChild(NodeType.SHAPE, 1, true)).getChart();

        Assert.assertFalse(chart.getAxisX().getScaling().getMinimum().isAuto());
        Assert.assertEquals(new AxisBound(DocumentHelper.createDate(1980, 1, 1)), chart.getAxisX().getScaling().getMinimum());
        Assert.assertEquals(new AxisBound(DocumentHelper.createDate(1990, 1, 1)), chart.getAxisX().getScaling().getMaximum());

        Assert.assertTrue(chart.getAxisY().getScaling().getMinimum().isAuto());
    }

    @Test
    public void chartLegend() throws Exception {
        //ExStart
        //ExFor:Chart.Legend
        //ExFor:ChartLegend
        //ExFor:ChartLegend.Overlay
        //ExFor:ChartLegend.Position
        //ExFor:LegendPosition
        //ExSummary:Shows how to edit the appearance of a chart's legend.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line graph
        Shape chartShape = builder.insertChart(ChartType.LINE, 450.0, 300.0);
        Chart chart = chartShape.getChart();

        // Get its legend
        ChartLegend legend = chart.getLegend();

        // By default, other elements of a chart will not overlap with its legend
        Assert.assertFalse(legend.getOverlay());

        // We can move its position by setting this attribute
        legend.setPosition(LegendPosition.TOP_RIGHT);

        doc.save(getArtifactsDir() + "Charts.ChartLegend.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.ChartLegend.docx");

        legend = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart().getLegend();

        Assert.assertFalse(legend.getOverlay());
        Assert.assertEquals(LegendPosition.TOP_RIGHT, legend.getPosition());
    }

    @Test
    public void axisCross() throws Exception {
        //ExStart
        //ExFor:ChartAxis.AxisBetweenCategories
        //ExFor:ChartAxis.CrossesAt
        //ExSummary:Shows how to get a graph axis to cross at a custom location.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart, which is populated by default values
        Shape shape = builder.insertChart(ChartType.COLUMN, 450.0, 250.0);
        Chart chart = shape.getChart();

        // Get the Y-axis to cross at a value of 3.0, making 3.0 the new Y-zero of our column chart
        // This effectively means that all the columns with Y-values about 3.0 will be above the Y-centre and point up,
        // while ones below 3.0 will point down
        ChartAxis axis = chart.getAxisX();
        axis.setAxisBetweenCategories(true);
        axis.setCrosses(AxisCrosses.CUSTOM);
        axis.setCrossesAt(3.0);

        doc.save(getArtifactsDir() + "Charts.AxisCross.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.AxisCross.docx");
        axis = ((Shape) doc.getChild(NodeType.SHAPE, 0, true)).getChart().getAxisX();

        Assert.assertTrue(axis.getAxisBetweenCategories());
        Assert.assertEquals(AxisCrosses.CUSTOM, axis.getCrosses());
        Assert.assertEquals(3.0, axis.getCrossesAt());
    }

    @Test
    public void chartAxisDisplayUnit() throws Exception {
        //ExStart
        //ExFor:AxisBuiltInUnit
        //ExFor:ChartAxis.DisplayUnit
        //ExFor:ChartAxis.MajorUnitIsAuto
        //ExFor:ChartAxis.MajorUnitScale
        //ExFor:ChartAxis.MinorUnitIsAuto
        //ExFor:ChartAxis.MinorUnitScale
        //ExFor:ChartAxis.TickLabelSpacing
        //ExFor:ChartAxis.TickLabelAlignment
        //ExFor:AxisDisplayUnit
        //ExFor:AxisDisplayUnit.CustomUnit
        //ExFor:AxisDisplayUnit.Unit
        //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart, which is populated by default values
        Shape shape = builder.insertChart(ChartType.SCATTER, 450.0, 250.0);
        Chart chart = shape.getChart();

        // Set they Y axis to show major ticks every at every 10 units and minor ticks at every 1 units
        ChartAxis axis = chart.getAxisY();
        axis.setMajorTickMark(AxisTickMark.OUTSIDE);
        axis.setMinorTickMark(AxisTickMark.OUTSIDE);

        axis.setMajorUnit(10.0);
        axis.setMinorUnit(1.0);

        // Stretch out the bounds of the axis out to show 3 major ticks and 27 minor ticks
        axis.getScaling().setMinimum(new AxisBound(-10));
        axis.getScaling().setMaximum(new AxisBound(20.0));

        // Do the same for the X-axis
        axis = chart.getAxisX();
        axis.setMajorTickMark(AxisTickMark.INSIDE);
        axis.setMinorTickMark(AxisTickMark.INSIDE);
        axis.setMajorUnit(10.0);
        axis.getScaling().setMinimum(new AxisBound(-10));
        axis.getScaling().setMaximum(new AxisBound(30.0));

        // We can also use this attribute to set minor tick spacing
        axis.setTickLabelSpacing(2);
        // We can define text alignment when axis tick labels are multi-line
        // MS Word aligns them to the center by default
        axis.setTickLabelAlignment(ParagraphAlignment.RIGHT);

        // Get the axis to display values, but in millions
        axis.getDisplayUnit().setUnit(AxisBuiltInUnit.MILLIONS);
        Assert.assertEquals(AxisBuiltInUnit.MILLIONS, axis.getDisplayUnit().getUnit()); //ExSkip

        // Besides the built-in axis units we can choose from,
        // we can also set the axis to display values in some custom denomination, using the following attribute
        // The statement below is equivalent to the one above
        axis.getDisplayUnit().setCustomUnit(1000000.0);
        Assert.assertEquals(AxisBuiltInUnit.CUSTOM, axis.getDisplayUnit().getUnit()); //ExSkip

        doc.save(getArtifactsDir() + "Charts.ChartAxisDisplayUnit.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Charts.ChartAxisDisplayUnit.docx");
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertEquals(450.0d, shape.getWidth());
        Assert.assertEquals(250.0d, shape.getHeight());

        axis = shape.getChart().getAxisX();

        Assert.assertEquals(AxisTickMark.INSIDE, axis.getMajorTickMark());
        Assert.assertEquals(AxisTickMark.INSIDE, axis.getMinorTickMark());
        Assert.assertEquals(10.0d, axis.getMajorUnit());
        Assert.assertEquals(-10.0d, axis.getScaling().getMinimum().getValue());
        Assert.assertEquals(30.0d, axis.getScaling().getMaximum().getValue());
        Assert.assertEquals(1, axis.getTickLabelSpacing());
        Assert.assertEquals(ParagraphAlignment.RIGHT, axis.getTickLabelAlignment());
        Assert.assertEquals(AxisBuiltInUnit.CUSTOM, axis.getDisplayUnit().getUnit());
        Assert.assertEquals(1000000.0d, axis.getDisplayUnit().getCustomUnit());

        axis = shape.getChart().getAxisY();

        Assert.assertEquals(AxisTickMark.OUTSIDE, axis.getMajorTickMark());
        Assert.assertEquals(AxisTickMark.OUTSIDE, axis.getMinorTickMark());
        Assert.assertEquals(10.0d, axis.getMajorUnit());
        Assert.assertEquals(1.0d, axis.getMinorUnit());
        Assert.assertEquals(-10.0d, axis.getScaling().getMinimum().getValue());
        Assert.assertEquals(20.0d, axis.getScaling().getMaximum().getValue());
    }
}
