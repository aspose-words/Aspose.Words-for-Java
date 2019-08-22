package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ChartType;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.ShapeType;
import com.aspose.words.Chart;
import com.aspose.words.ChartTitle;
import com.aspose.words.ChartSeries;
import com.aspose.words.ChartDataLabel;
import com.aspose.words.ChartSeriesCollection;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.NodeType;
import com.aspose.words.ChartAxisType;
import com.aspose.words.AxisCategoryType;
import com.aspose.words.AxisCrosses;
import com.aspose.words.AxisTickMark;
import com.aspose.words.AxisTickLabelPosition;
import com.aspose.words.AxisTimeUnit;
import com.aspose.words.AxisBuiltInUnit;
import com.aspose.words.AxisScaleType;
import com.aspose.words.ChartAxis;
import com.aspose.ms.System.DateTime;
import com.aspose.words.AxisBound;
import java.util.Iterator;
import com.aspose.words.MarkerSymbol;
import com.aspose.words.ChartDataPoint;
import com.aspose.ms.System.msConsole;
import com.aspose.words.ChartLegend;
import com.aspose.words.LegendPosition;
import com.aspose.words.ParagraphAlignment;
import org.testng.annotations.DataProvider;


@Test
public class ExCharts extends ApiExampleBase
{
    @Test
    public void chartTitle() throws Exception
    {
        //ExStart
        //ExFor:Charts.Chart
        //ExFor:Charts.Chart.Title
        //ExFor:Charts.ChartTitle
        //ExFor:Charts.ChartTitle.Overlay
        //ExFor:Charts.ChartTitle.Show
        //ExFor:Charts.ChartTitle.Text
        //ExSummary:Shows how to insert a chart and change its title.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a bar chart
        Shape chartShape = builder.insertChart(ChartType.BAR, 400.0, 300.0);

        msAssert.areEqual(ShapeType.NON_PRIMITIVE, chartShape.getShapeType());
        Assert.assertTrue(chartShape.hasChart());

        // Get the chart object from the containing shape
        Chart chart = chartShape.getChart();
        
        // Set the title text, which appears at the top center of the chart and modify its appearance
        ChartTitle title = chart.getTitle();
        title.setText("MyChart");
        title.setOverlay(true);
        title.setShow(true);

        doc.save(getArtifactsDir() + "Charts.ChartTitle.docx");
        //ExEnd
    }

    @Test
    public void numberFormat() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();
        chart.getTitle().setText("Data Labels With Different Number Format");

        // Delete default generated series.
        chart.getSeries().clear();

        // Add new series
        ChartSeries series0 =
            chart.getSeries().add("AW Series 0", new String[] { "AW0", "AW1", "AW2" }, new double[] { 2.5, 1.5, 3.5 });

        // Add DataLabel to the first point of the first series.
        ChartDataLabel chartDataLabel0 = series0.getDataLabels().add(0);
        chartDataLabel0.setShowValue(true);

        // Set currency format code.
        chartDataLabel0.getNumberFormat().setFormatCode("\"$\"#,##0.00");

        ChartDataLabel chartDataLabel1 = series0.getDataLabels().add(1);
        chartDataLabel1.setShowValue(true);

        // Set date format code.
        chartDataLabel1.getNumberFormat().setFormatCode("d/mm/yyyy");

        ChartDataLabel chartDataLabel2 = series0.getDataLabels().add(2);
        chartDataLabel2.setShowValue(true);

        // Set percentage format code.
        chartDataLabel2.getNumberFormat().setFormatCode("0.00%");

        // Or you can set format code to be linked to a source cell,
        // in this case NumberFormat will be reset to general and inherited from a source cell.
        chartDataLabel2.getNumberFormat().isLinkedToSource(true);

        doc.save(getArtifactsDir() + "Charts.NumberFormat.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "Charts.NumberFormat.docx", getGoldsDir() + "DocumentBuilder.NumberFormat Gold.docx"));
    }

    @Test
    public void dataArraysWrongSize() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        // Create category names array, second category will be null.
        String[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

        // Adding new series with empty (double.NaN) values.
        seriesColl.add("AW Series 1", categories, new double[] { 1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0 });
        seriesColl.add("AW Series 2", categories, new double[] { 2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0 });

        Assert.That(
            () => seriesColl.add("AW Series 3", categories, new double[] { Double.NaN, 4.0, 5.0, Double.NaN, Double.NaN }),
            Throws.<IllegalArgumentException>TypeOf());
        Assert.That(
            () => seriesColl.add("AW Series 4", categories,
                new double[] { Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN }),
            Throws.<IllegalArgumentException>TypeOf());
    }

    @Test
    public void emptyValuesInChartData() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add chart with default data.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();

        ChartSeriesCollection seriesColl = chart.getSeries();
        seriesColl.clear();

        // Create category names array, second category will be null.
        String[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

        // Adding new series with empty (double.NaN) values.
        seriesColl.add("AW Series 1", categories, new double[] { 1.0, 2.0, Double.NaN, 4.0, 5.0, 6.0 });
        seriesColl.add("AW Series 2", categories, new double[] { 2.0, 3.0, Double.NaN, 5.0, 6.0, 7.0 });
        seriesColl.add("AW Series 3", categories, new double[] { Double.NaN, 4.0, 5.0, Double.NaN, 7.0, 8.0 });
        seriesColl.add("AW Series 4", categories,
            new double[] { Double.NaN, Double.NaN, Double.NaN, Double.NaN, Double.NaN, 9.0 });

        doc.save(getArtifactsDir() + "Charts.EmptyValuesInChartData.docx");
    }

    @Test
    public void chartDefaultValues() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        builder.insertChart(ChartType.COLUMN_3_D, 432.0, 252.0);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Shape shapeNode = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        Chart chart = shapeNode.getChart();

        // Assert X axis
        msAssert.areEqual(ChartAxisType.CATEGORY, chart.getAxisX().getType());
        msAssert.areEqual(AxisCategoryType.AUTOMATIC, chart.getAxisX().getCategoryType());
        msAssert.areEqual(AxisCrosses.AUTOMATIC, chart.getAxisX().getCrosses());
        msAssert.areEqual(false, chart.getAxisX().getReverseOrder());
        msAssert.areEqual(AxisTickMark.NONE, chart.getAxisX().getMajorTickMark());
        msAssert.areEqual(AxisTickMark.NONE, chart.getAxisX().getMinorTickMark());
        msAssert.areEqual(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisX().getTickLabelPosition());
        msAssert.areEqual(1, chart.getAxisX().getMajorUnit());
        msAssert.areEqual(true, chart.getAxisX().getMajorUnitIsAuto());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisX().getMajorUnitScale());
        msAssert.areEqual(0.5, chart.getAxisX().getMinorUnit());
        msAssert.areEqual(true, chart.getAxisX().getMinorUnitIsAuto());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisX().getMinorUnitScale());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisX().getBaseTimeUnit());
        msAssert.areEqual("General", chart.getAxisX().getNumberFormat().getFormatCode());
        msAssert.areEqual(100, chart.getAxisX().getTickLabelOffset());
        msAssert.areEqual(AxisBuiltInUnit.NONE, chart.getAxisX().getDisplayUnit().getUnit());
        msAssert.areEqual(true, chart.getAxisX().getAxisBetweenCategories());
        msAssert.areEqual(AxisScaleType.LINEAR, chart.getAxisX().getScaling().getType());
        msAssert.areEqual(1, chart.getAxisX().getTickLabelSpacing());
        msAssert.areEqual(true, chart.getAxisX().getTickLabelSpacingIsAuto());
        msAssert.areEqual(1, chart.getAxisX().getTickMarkSpacing());
        msAssert.areEqual(false, chart.getAxisX().getHidden());

        // Assert Y axis
        msAssert.areEqual(ChartAxisType.VALUE, chart.getAxisY().getType());
        msAssert.areEqual(AxisCategoryType.CATEGORY, chart.getAxisY().getCategoryType());
        msAssert.areEqual(AxisCrosses.AUTOMATIC, chart.getAxisY().getCrosses());
        msAssert.areEqual(false, chart.getAxisY().getReverseOrder());
        msAssert.areEqual(AxisTickMark.NONE, chart.getAxisY().getMajorTickMark());
        msAssert.areEqual(AxisTickMark.NONE, chart.getAxisY().getMinorTickMark());
        msAssert.areEqual(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisY().getTickLabelPosition());
        msAssert.areEqual(1, chart.getAxisY().getMajorUnit());
        msAssert.areEqual(true, chart.getAxisY().getMajorUnitIsAuto());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisY().getMajorUnitScale());
        msAssert.areEqual(0.5, chart.getAxisY().getMinorUnit());
        msAssert.areEqual(true, chart.getAxisY().getMinorUnitIsAuto());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisY().getMinorUnitScale());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisY().getBaseTimeUnit());
        msAssert.areEqual("General", chart.getAxisY().getNumberFormat().getFormatCode());
        msAssert.areEqual(100, chart.getAxisY().getTickLabelOffset());
        msAssert.areEqual(AxisBuiltInUnit.NONE, chart.getAxisY().getDisplayUnit().getUnit());
        msAssert.areEqual(true, chart.getAxisY().getAxisBetweenCategories());
        msAssert.areEqual(AxisScaleType.LINEAR, chart.getAxisY().getScaling().getType());
        msAssert.areEqual(1, chart.getAxisY().getTickLabelSpacing());
        msAssert.areEqual(true, chart.getAxisY().getTickLabelSpacingIsAuto());
        msAssert.areEqual(1, chart.getAxisY().getTickMarkSpacing());
        msAssert.areEqual(false, chart.getAxisY().getHidden());

        // Assert Z axis
        msAssert.areEqual(ChartAxisType.SERIES, chart.getAxisZ().getType());
        msAssert.areEqual(AxisCategoryType.CATEGORY, chart.getAxisZ().getCategoryType());
        msAssert.areEqual(AxisCrosses.AUTOMATIC, chart.getAxisZ().getCrosses());
        msAssert.areEqual(false, chart.getAxisZ().getReverseOrder());
        msAssert.areEqual(AxisTickMark.NONE, chart.getAxisZ().getMajorTickMark());
        msAssert.areEqual(AxisTickMark.NONE, chart.getAxisZ().getMinorTickMark());
        msAssert.areEqual(AxisTickLabelPosition.NEXT_TO_AXIS, chart.getAxisZ().getTickLabelPosition());
        msAssert.areEqual(1, chart.getAxisZ().getMajorUnit());
        msAssert.areEqual(true, chart.getAxisZ().getMajorUnitIsAuto());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisZ().getMajorUnitScale());
        msAssert.areEqual(0.5, chart.getAxisZ().getMinorUnit());
        msAssert.areEqual(true, chart.getAxisZ().getMinorUnitIsAuto());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisZ().getMinorUnitScale());
        msAssert.areEqual(AxisTimeUnit.AUTOMATIC, chart.getAxisZ().getBaseTimeUnit());
        msAssert.areEqual("", chart.getAxisZ().getNumberFormat().getFormatCode());
        msAssert.areEqual(100, chart.getAxisZ().getTickLabelOffset());
        msAssert.areEqual(AxisBuiltInUnit.NONE, chart.getAxisZ().getDisplayUnit().getUnit());
        msAssert.areEqual(true, chart.getAxisZ().getAxisBetweenCategories());
        msAssert.areEqual(AxisScaleType.LINEAR, chart.getAxisZ().getScaling().getType());
        msAssert.areEqual(1, chart.getAxisZ().getTickLabelSpacing());
        msAssert.areEqual(true, chart.getAxisZ().getTickLabelSpacingIsAuto());
        msAssert.areEqual(1, chart.getAxisZ().getTickMarkSpacing());
        msAssert.areEqual(false, chart.getAxisZ().getHidden());
    }

    @Test
    public void insertChartUsingAxisProperties() throws Exception
    {
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
        //ExFor:Charts.AxisCategoryType
        //ExFor:Charts.AxisCrosses
        //ExFor:Charts.Chart.AxisX
        //ExFor:Charts.Chart.AxisY
        //ExFor:Charts.Chart.AxisZ
        //ExSummary:Shows how to insert chart using the axis options for detailed configuration.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();
        chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 640.0, 320.0, 280.0, 120.0, 150.0 });

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
        //ExEnd

        doc.save(getArtifactsDir() + "Charts.InsertChartUsingAxisProperties.docx");
        doc.save(getArtifactsDir() + "Charts.InsertChartUsingAxisProperties.pdf");
    }

    @Test
    public void insertChartWithDateTimeValues() throws Exception
    {
        //ExStart
        //ExFor:AxisBound
        //ExFor:AxisBound.#ctor(Double)
        //ExFor:AxisBound.#ctor(DateTime)
        //ExFor:AxisScaling.Minimum
        //ExFor:AxisScaling.Maximum
        //ExFor:ChartAxis.Scaling
        //ExFor:Charts.AxisTickMark
        //ExFor:Charts.AxisTickLabelPosition
        //ExFor:Charts.AxisTimeUnit
        //ExFor:Charts.ChartAxis.BaseTimeUnit
        //ExSummary:Shows how to insert chart with date/time values
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();
        
        // Clear demo data.
        chart.getSeries().clear();

        // Fill data.
        chart.getSeries().addInternal("Aspose Test Series",
            new DateTime[]
            {
                new DateTime(2017, 11, 6), new DateTime(2017, 11, 9), new DateTime(2017, 11, 15),
                new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
            },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });

        ChartAxis xAxis = chart.getAxisX();
        ChartAxis yAxis = chart.getAxisY();

        // Set X axis bounds.
        xAxis.getScaling().setMinimum(new AxisBound(new DateTime(2017, 11, 5).toOADate()));
        xAxis.getScaling().setMaximum(new AxisBound(new DateTime(2017, 12, 3)));

        // Set major units to a week and minor units to a day.
        xAxis.setBaseTimeUnit(AxisTimeUnit.DAYS);
        xAxis.setMajorUnit(7.0);
        xAxis.setMinorUnit(1.0);
        xAxis.setMajorTickMark(AxisTickMark.CROSS);
        xAxis.setMinorTickMark(AxisTickMark.OUTSIDE);

        // Define Y axis properties.
        yAxis.setTickLabelPosition(AxisTickLabelPosition.HIGH);
        yAxis.setMajorUnit(100.0);
        yAxis.setMinorUnit(50.0);
        yAxis.getDisplayUnit().setUnit(AxisBuiltInUnit.HUNDREDS);
        yAxis.getScaling().setMinimum(new AxisBound(100.0));
        yAxis.getScaling().setMaximum(new AxisBound(700.0));

        doc.save(getArtifactsDir() + "Charts.ChartAxisProperties.docx");
        //ExEnd
    }

    @Test
    public void hideChartAxis() throws Exception
    {
        //ExStart
        //ExFor:ChartAxis.Hidden
        //ExSummary:Shows how to hide chart axises.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.LINE, 432.0, 252.0);
        Chart chart = shape.getChart();
        chart.getAxisX().setHidden(true);
        chart.getAxisY().setHidden(true);

        // Clear demo data.
        chart.getSeries().clear();
        chart.getSeries().add("AW Series 1",
            new String[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
            new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

        MemoryStream stream = new MemoryStream();
        doc.save(stream, SaveFormat.DOCX);

        shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);
        chart = shape.getChart();

        msAssert.areEqual(true, chart.getAxisX().getHidden());
        msAssert.areEqual(true, chart.getAxisY().getHidden());
        //ExEnd
    }

    @Test
    public void setNumberFormatToChartAxis() throws Exception
    {
        //ExStart
        //ExFor:ChartAxis.NumberFormat
        //ExFor:Charts.ChartNumberFormat
        //ExFor:ChartNumberFormat.FormatCode
        //ExFor:Charts.ChartNumberFormat.IsLinkedToSource
        //ExSummary:Shows how to set formatting for chart values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.COLUMN, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });

        // Set number format.
        chart.getAxisY().getNumberFormat().setFormatCode("#,##0");

        // Set this to override the above value and draw the number format from the source cell
        Assert.assertFalse(chart.getAxisY().getNumberFormat().isLinkedToSource());
        //ExEnd

        doc.save(getArtifactsDir() + "Charts.SetNumberFormatToChartAxis.docx");
        doc.save(getArtifactsDir() + "Charts.SetNumberFormatToChartAxis.pdf");
    }

    // Note: Tests below used for verification conversion docx to pdf and the correct display.
    // For now, the results check manually.
    @Test (dataProvider = "testDisplayChartsWithConversionDataProvider")
    public void testDisplayChartsWithConversion(/*ChartType*/int chartType) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(chartType, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });

        doc.save(getArtifactsDir() + "Charts.TestDisplayChartsWithConversion.docx");
        doc.save(getArtifactsDir() + "Charts.TestDisplayChartsWithConversion.pdf");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "testDisplayChartsWithConversionDataProvider")
	public static Object[][] testDisplayChartsWithConversionDataProvider() throws Exception
	{
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
    public void surface3DChart() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.SURFACE_3_D, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series 1",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 });

        chart.getSeries().add("Aspose Test Series 2",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 900000.0, 50000.0, 1100000.0, 400000.0, 2500000.0 });

        chart.getSeries().add("Aspose Test Series 3",
            new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
            new double[] { 500000.0, 820000.0, 1500000.0, 400000.0, 100000.0 });

        doc.save(getArtifactsDir() + "Charts.SurfaceChart.docx");
        doc.save(getArtifactsDir() + "Charts.SurfaceChart.pdf");
    }

    @Test
    public void bubbleChart() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert chart.
        Shape shape = builder.insertChart(ChartType.BUBBLE, 432.0, 252.0);
        Chart chart = shape.getChart();

        // Clear demo data.
        chart.getSeries().clear();

        chart.getSeries().add("Aspose Test Series",
            new double[] { 2900000.0, 350000.0, 1100000.0, 400000.0, 400000.0 },
            new double[] { 1900000.0, 850000.0, 2100000.0, 600000.0, 1500000.0 },
            new double[] { 900000.0, 450000.0, 2500000.0, 800000.0, 500000.0 });

        doc.save(getArtifactsDir() + "Charts.BubbleChart.docx");
        doc.save(getArtifactsDir() + "Charts.BubbleChart.pdf");
    }

    //ExStart
    //ExFor:Charts.ChartSeries
    //ExFor:Charts.ChartSeries.DataLabels
    //ExFor:Charts.ChartSeries.DataPoints
    //ExFor:Charts.ChartSeries.Name
    //ExFor:Charts.ChartDataLabel
    //ExFor:Charts.ChartDataLabel.Index
    //ExFor:Charts.ChartDataLabel.IsVisible
    //ExFor:Charts.ChartDataLabel.NumberFormat
    //ExFor:Charts.ChartDataLabel.Separator
    //ExFor:Charts.ChartDataLabel.ShowCategoryName
    //ExFor:Charts.ChartDataLabel.ShowDataLabelsRange
    //ExFor:Charts.ChartDataLabel.ShowLeaderLines
    //ExFor:Charts.ChartDataLabel.ShowLegendKey
    //ExFor:Charts.ChartDataLabel.ShowPercentage
    //ExFor:Charts.ChartDataLabel.ShowSeriesName
    //ExFor:Charts.ChartDataLabel.ShowValue
    //ExFor:Charts.ChartDataLabelCollection
    //ExFor:Charts.ChartDataLabelCollection.Add(System.Int32)
    //ExFor:Charts.ChartDataLabelCollection.Clear
    //ExFor:Charts.ChartDataLabelCollection.Count
    //ExFor:Charts.ChartDataLabelCollection.GetEnumerator
    //ExFor:Charts.ChartDataLabelCollection.Item(System.Int32)
    //ExFor:Charts.ChartDataLabelCollection.RemoveAt(System.Int32)
    //ExSummary:Shows how to apply labels to data points in a chart.
    @Test //ExSkip
    public void chartDataLabels() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // Use a document builder to insert a bar chart
        Shape chartShape = builder.insertChart(ChartType.LINE, 400.0, 300.0);

        // Get the chart object from the containing shape
        Chart chart = chartShape.getChart();

        // The chart already contains demo data comprised of 3 series each with 4 categories
        msAssert.areEqual(3, chart.getSeries().getCount());
        msAssert.areEqual("Series 1", chart.getSeries().get(0).getName());

        // Apply data labels to every series in the graph
        for (ChartSeries series : chart.getSeries())
        {
            applyDataLabels(series, 4, "000.0", ", ");
            msAssert.areEqual(4, series.getDataLabels().getCount());
        }

        // Get the enumerator for a data label collection
        Iterator<ChartDataLabel> enumerator = chart.getSeries().get(0).getDataLabels().iterator();
        try /*JAVA: was using*/
        {
            // And use it to go over all the data labels in one series and change their separator
            while (enumerator.hasNext())
            {
                msAssert.areEqual(", ", enumerator.next().getSeparator());
                enumerator.next().setSeparator(" & ");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // If the chart looks too busy, we can remove data labels one by one
        chart.getSeries().get(1).getDataLabels().removeAt(2);

        // We can also clear an entire data label collection for one whole series
        chart.getSeries().get(2).getDataLabels().clear();

        doc.save(getArtifactsDir() + "Charts.ChartDataLabels.docx");
    }

    /// <summary>
    /// Apply uniform data labels with custom number format and separator to a number (determined by labelsCount) of data points in a series
    /// </summary>
    private void applyDataLabels(ChartSeries series, int labelsCount, String numberFormat, String separator)
    {
        for (int i = 0; i < labelsCount; i++)
        {
            ChartDataLabel label = series.getDataLabels().add(i);
            Assert.assertFalse(label.isVisible());

            // Edit the appearance of the new data label
            label.setShowCategoryName(true);
            label.setShowSeriesName(true);
            label.setShowValue(true);
            label.setShowLeaderLines(true);
            label.setShowLegendKey(true);
            label.setShowPercentage(false);
            Assert.assertFalse(label.getShowDataLabelsRange());

            // Apply number format and separator
            label.getNumberFormat().setFormatCode(numberFormat);
            label.setSeparator(separator);

            // The label automatically becomes visible
            Assert.assertTrue(label.isVisible());
        }
    }
    //ExEnd

    //ExStart
    //ExFor:Charts.ChartSeries.Smooth
    //ExFor:Charts.ChartDataPoint
    //ExFor:Charts.ChartDataPoint.Index
    //ExFor:Charts.ChartDataPointCollection
    //ExFor:Charts.ChartDataPointCollection.Add(System.Int32)
    //ExFor:Charts.ChartDataPointCollection.Clear
    //ExFor:Charts.ChartDataPointCollection.Count
    //ExFor:Charts.ChartDataPointCollection.GetEnumerator
    //ExFor:Charts.ChartDataPointCollection.Item(System.Int32)
    //ExFor:Charts.ChartDataPointCollection.RemoveAt(System.Int32)
    //ExFor:Charts.ChartMarker
    //ExFor:Charts.ChartMarker.Size
    //ExFor:Charts.ChartMarker.Symbol
    //ExFor:Charts.IChartDataPoint
    //ExFor:Charts.IChartDataPoint.InvertIfNegative
    //ExFor:Charts.IChartDataPoint.Marker
    //ExFor:Charts.MarkerSymbol
    //ExSummary:Shows how to customize chart data points.
    @Test
    public void chartDataPoint() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a line chart, which will have default data that we will use
        Shape shape = builder.insertChart(ChartType.LINE, 500.0, 350.0);
        Chart chart = shape.getChart();

        // Apply diamond-shaped data points to the line of the first series
        for (ChartSeries series : chart.getSeries())
        {
            applyDataPoints(series, 4, MarkerSymbol.DIAMOND, 15);
        }

        // We can further decorate a series line by smoothing it
        chart.getSeries().get(0).setSmooth(true);

        // Get the enumerator for the data point collection from one series
        Iterator<ChartDataPoint> enumerator = chart.getSeries().get(0).getDataPoints().iterator();
        try /*JAVA: was using*/
        {
            // And use it to go over all the data labels in one series and change their separator
            while (enumerator.hasNext())
            {
                Assert.assertFalse(enumerator.next().getInvertIfNegative());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // If the chart looks too busy, we can remove data points one by one
        chart.getSeries().get(1).getDataPoints().removeAt(2);

        // We can also clear an entire data point collection for one whole series
        chart.getSeries().get(2).getDataPoints().clear();

        doc.save(getArtifactsDir() + "Charts.ChartDataPoint.docx");
    }

    /// <summary>
    /// Applies a number of data points to a series
    /// </summary>
    private void applyDataPoints(ChartSeries series, int dataPointsCount, /*MarkerSymbol*/int markerSymbol, int dataPointSize)
    {
        for (int i = 0; i < dataPointsCount; i++)
        {
            ChartDataPoint point = series.getDataPoints().add(i);
            point.getMarker().setSymbol(markerSymbol);
            point.getMarker().setSize(dataPointSize);

            msAssert.areEqual(i, point.getIndex());
        }
    }
    //ExEnd

    @Test
    public void pieChartExplosion() throws Exception
    {
        //ExStart
        //ExFor:Charts.IChartDataPoint.Explosion
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
    }

    @Test
    public void bubble3D() throws Exception
    {
        //ExStart
        //ExFor:Charts.ChartDataLabel.ShowBubbleSize
        //ExFor:Charts.IChartDataPoint.Bubble3D
        //ExSummary:Demonstrates bubble chart-exclusive features.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bubble chart with 3D effects on each bubble
        Shape shape = builder.insertChart(ChartType.BUBBLE_3_D, 500.0, 350.0);
        Chart chart = shape.getChart();

        Assert.assertTrue(chart.getSeries().get(0).getBubble3D());

        // Apply a data label to each bubble that displays the size of its bubble
        for (int i = 0; i < 3; i++)
        {
            ChartDataLabel cdl = chart.getSeries().get(0).getDataLabels().add(i);
            cdl.setShowBubbleSize(true);
        }
        
        doc.save(getArtifactsDir() + "Charts.Bubble3D.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:Charts.ChartAxis.Type
    //ExFor:Charts.ChartAxisType
    //ExFor:Charts.ChartType
    //ExFor:Charts.Chart.Series
    //ExFor:Charts.ChartSeriesCollection.Add(String,DateTime[],Double[])
    //ExFor:Charts.ChartSeriesCollection.Add(String,Double[],Double[])
    //ExFor:Charts.ChartSeriesCollection.Add(String,Double[],Double[],Double[])
    //ExFor:Charts.ChartSeriesCollection.Add(String,String[],Double[])
    //ExSummary:Shows an appropriate graph type for each chart series.
    @Test //ExSkip
    public void chartSeriesCollection() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // There are 4 ways of populating a chart's series collection
        // 1: Each series has a string array of categories, each with a corresponding data value
        // Some of the other possible applications are bar, column, line and surface charts
        Chart chart = appendChart(builder, ChartType.COLUMN, 300.0, 300.0);

        // Create and name 3 categories with a string array
        String[] categories = { "Category 1", "Category 2", "Category 3" };

        // Create 2 series of data, each with one point for every category
        // This will generate a column graph with 3 clusters of 2 bars
        chart.getSeries().add("Series 1", categories, new double[] { 76.6, 82.1, 91.6 });
        chart.getSeries().add("Series 2", categories, new double[] { 64.2, 79.5, 94.0 });

        // Categories are distributed along the X-axis while values are distributed along the Y-axis
        msAssert.areEqual(ChartAxisType.CATEGORY, chart.getAxisX().getType());
        msAssert.areEqual(ChartAxisType.VALUE, chart.getAxisY().getType());

        // 2: Each series will have a collection of dates with a corresponding value for each date
        // Area, radar and stock charts are some of the appropriate chart types for this
        chart = appendChart(builder, ChartType.AREA, 300.0, 300.0);

        // Create a collection of dates to serve as categories
        DateTime[] dates = { new DateTime(2014, 3, 31),
            new DateTime(2017, 1, 23),
            new DateTime(2017, 6, 18),
            new DateTime(2019, 11, 22),
            new DateTime(2020, 9, 7)
        };

        // Add one series with one point for each date
        // Our sporadic dates will be distributed along the X-axis in a linear fashion 
        chart.getSeries().addInternal("Series 1", dates, new double[] { 15.8, 21.5, 22.9, 28.7, 33.1 });

        // 3: Each series will take two data arrays
        // Appropriate for scatter plots
        chart = appendChart(builder, ChartType.SCATTER, 300.0, 300.0);

        // In each series, the first array contains the X-coordinates and the second contains respective Y-coordinates of points
        chart.getSeries().add("Series 1", new double[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 }, new double[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });
        chart.getSeries().add("Series 2", new double[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 }, new double[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

        // Both axes are value axes in this case
        msAssert.areEqual(ChartAxisType.VALUE, chart.getAxisX().getType());
        msAssert.areEqual(ChartAxisType.VALUE, chart.getAxisY().getType());

        // 4: Each series will be built from three data arrays, used for bubble charts
        chart = appendChart(builder, ChartType.BUBBLE, 300.0, 300.0);

        // The first two arrays contain X/Y coordinates like above and the third determines the thickness of each point
        chart.getSeries().add("Series 1", new double[] { 1.1, 5.0, 9.8 }, new double[] { 1.2, 4.9, 9.9 }, new double[] { 2.0, 4.0, 8.0 });

        doc.save(getArtifactsDir() + "Charts.ChartSeriesCollection.docx");
    }
    
    /// <summary>
    /// Get the DocumentBuilder to insert a chart of a specified ChartType, width and height and clean out its default data
    /// </summary>
    private Chart appendChart(DocumentBuilder builder, /*ChartType*/int chartType, double width, double height) throws Exception
    {
        Shape chartShape = builder.insertChart(chartType, width, height);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();

        msAssert.areEqual(0, chart.getSeries().getCount());

        return chart;
    }
    //ExEnd

    @Test
    public void chartSeriesCollectionModify() throws Exception
    {
        //ExStart
        //ExFor:Charts.ChartSeriesCollection
        //ExFor:Charts.ChartSeriesCollection.Clear
        //ExFor:Charts.ChartSeriesCollection.Count
        //ExFor:Charts.ChartSeriesCollection.GetEnumerator
        //ExFor:Charts.ChartSeriesCollection.Item(Int32)
        //ExFor:Charts.ChartSeriesCollection.RemoveAt(Int32)
        //ExSummary:Shows how to work with a chart's data collection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a bar chart
        Shape chartShape = builder.insertChart(ChartType.COLUMN, 400.0, 300.0);
        Chart chart = chartShape.getChart();

        // All charts come with demo data
        // This column chart currently has 3 series with 4 categories, which means 4 clusters, 3 columns in each
        ChartSeriesCollection chartData = chart.getSeries();
        msAssert.areEqual(3, chartData.getCount());

        // Iterate through the series with an enumerator and print their names
        Iterator<ChartSeries> enumerator = chart.getSeries().iterator();
        try /*JAVA: was using*/
        {
            // And use it to go over all the data labels in one series and change their separator
            while (enumerator.hasNext())
            {
                msConsole.writeLine(enumerator.next().getName());
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can add new data by adding a new series to the collection, with categories and data
        // We will match the existing category/series names in the demo data and add a 4th column to each column cluster
        String[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };
        chart.getSeries().add("Series 4", categories, new double[] { 4.4, 7.0, 3.5, 2.1 });

        msAssert.areEqual(4, chartData.getCount());
        msAssert.areEqual("Series 4", chartData.get(3).getName());

        // We can remove series by index
        chartData.removeAt(2);

        msAssert.areEqual(3, chartData.getCount());
        msAssert.areEqual("Series 4", chartData.get(2).getName());

        // We can also remove out all the series
        // This leaves us with an empty graph and is a convenient way of wiping out demo data
        chartData.clear();

        msAssert.areEqual(0, chartData.getCount());
        //ExEnd
    }

    @Test
    public void axisScaling() throws Exception
    {
        //ExStart
        //ExFor:Charts.AxisScaleType
        //ExFor:Charts.AxisScaling
        //ExFor:Charts.AxisScaling.LogBase
        //ExFor:Charts.AxisScaling.Type
        //ExSummary:Shows how to set up logarithmic axis scaling.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart and clear its default data series
        Shape chartShape = builder.insertChart(ChartType.SCATTER, 450.0, 300.0);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();

        // Insert a series with X/Y coordinates for 5 points
        chart.getSeries().add("Series 1", new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, new double[] { 1.0, 20.0, 400.0, 8000.0, 160000.0 });

        // The scaling of the X axis is linear by default, which means it will display "0, 1, 2, 3..."
        msAssert.areEqual(AxisScaleType.LINEAR, chart.getAxisX().getScaling().getType());

        // Linear axis scaling is suitable for our X-values, but not our erratic Y-values 
        // We can set the scaling of the Y-axis to Logarithmic with a base of 20
        // The Y-axis will now display "1, 20, 400, 8000...", which is ideal for accurate representation of this set of Y-values
        chart.getAxisY().getScaling().setType(AxisScaleType.LOGARITHMIC);
        chart.getAxisY().getScaling().setLogBase(20.0);

        doc.save(getArtifactsDir() + "Charts.AxisScaling.docx");
        //ExEnd
    }

    @Test
    public void axisBound() throws Exception
    {
        //ExStart
        //ExFor:Charts.AxisBound.#ctor
        //ExFor:Charts.AxisBound.IsAuto
        //ExFor:Charts.AxisBound.Value
        //ExFor:Charts.AxisBound.ValueAsDate
        //ExSummary:Shows how to set custom axis bounds.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a scatter chart, remove default data and populate it with data from a ChartSeries
        Shape chartShape = builder.insertChart(ChartType.SCATTER, 450.0, 300.0);
        Chart chart = chartShape.getChart();
        chart.getSeries().clear();
        chart.getSeries().add("Series 1", new double[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 }, new double[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

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
        DateTime[] dates = { new DateTime(1973, 5, 11),
            new DateTime(1981, 2, 4),
            new DateTime(1985, 9, 23),
            new DateTime(1989, 6, 28),
            new DateTime(1994, 12, 15)
        };

        // Assign a Y-value for each date 
        chart.getSeries().addInternal("Series 1", dates, new double[] { 3.0, 4.7, 5.9, 7.1, 8.9 });

        // These particular bounds will cut off categories from before 1980 and from 1990 and onwards
        // This narrows the amount of categories and values in the viewport from 5 to 3
        // Note that the graph still contains the out-of-range data because we can see the line tend towards it
        chart.getAxisX().getScaling().setMinimum(new AxisBound(new DateTime(1980, 1, 1)));
        chart.getAxisX().getScaling().setMaximum(new AxisBound(new DateTime(1990, 1, 1)));

        doc.save(getArtifactsDir() + "Charts.AxisBound.docx");
        //ExEnd
    }

    @Test
    public void chartLegend() throws Exception
    {
        //ExStart
        //ExFor:Charts.Chart.Legend
        //ExFor:Charts.ChartLegend
        //ExFor:Charts.ChartLegend.Overlay
        //ExFor:Charts.ChartLegend.Position
        //ExFor:Charts.LegendPosition
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
    }

    @Test
    public void axisCross() throws Exception
    {
        //ExStart
        //ExFor:Charts.ChartAxis.AxisBetweenCategories
        //ExFor:Charts.ChartAxis.CrossesAt
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
    }

    @Test
    public void chartAxisDisplayUnit() throws Exception
    {
        //ExStart
        //ExFor:Charts.AxisBuiltInUnit
        //ExFor:Charts.ChartAxis.DisplayUnit
        //ExFor:Charts.ChartAxis.MajorUnitIsAuto
        //ExFor:Charts.ChartAxis.MajorUnitScale
        //ExFor:Charts.ChartAxis.MinorUnitIsAuto
        //ExFor:Charts.ChartAxis.MinorUnitScale
        //ExFor:Charts.ChartAxis.TickLabelSpacing
        //ExFor:Charts.ChartAxis.TickLabelAlignment
        //ExFor:Charts.AxisDisplayUnit
        //ExFor:Charts.AxisDisplayUnit.CustomUnit
        //ExFor:Charts.AxisDisplayUnit.Unit
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

        // Besides the built-in axis units we can choose from,
        // we can also set the axis to display values in some custom denomination, using the following attribute
        // The statement below is equivalent to the one above
        axis.getDisplayUnit().setCustomUnit(1000000.0);

        doc.save(getArtifactsDir() + "Charts.ChartAxisDisplayUnit.docx");
        //ExEnd
    }
}
