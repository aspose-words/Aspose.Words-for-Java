// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import ApiExamples.TestData.TestClasses.MessageTestClass;
import com.aspose.words.ReportBuildOptions;
import org.testng.Assert;
import ApiExamples.TestData.TestClasses.NumericTestClass;
import ApiExamples.TestData.TestBuilders.NumericTestBuilder;
import com.aspose.ms.System.DateTime;
import ApiExamples.TestData.Common;
import ApiExamples.TestData.TestClasses.DocumentTestClass;
import ApiExamples.TestData.TestBuilders.DocumentTestBuilder;
import java.util.ArrayList;
import ApiExamples.TestData.TestClasses.ColorItemTestClass;
import ApiExamples.TestData.TestBuilders.ColorItemTestBuilder;
import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.ReportingEngine;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.FileAccess;
import com.aspose.ms.System.IO.File;
import com.aspose.words.ShapeType;
import ApiExamples.TestData.TestClasses.ImageTestClass;
import ApiExamples.TestData.TestBuilders.ImageTestBuilder;
import java.awt.image.BufferedImage;
import com.aspose.words.DocumentBuilder;
import ApiExamples.TestData.TestClasses.ClientTestClass;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Shape;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.ControlChar;
import com.aspose.ms.System.msString;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.SaveFormat;
import com.aspose.words.XmlDataSource;
import java.io.FileInputStream;
import com.aspose.words.JsonDataLoadOptions;
import com.aspose.words.JsonDataSource;
import com.aspose.words.JsonSimpleValueParseMode;
import com.aspose.words.CsvDataLoadOptions;
import com.aspose.words.CsvDataSource;
import com.aspose.words.SdtType;
import com.aspose.words.SdtListItem;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.MarkupLevel;
import java.lang.Class;
import org.testng.annotations.DataProvider;


@Test
public class ExReportingEngine extends ApiExampleBase
{
    private /*final*/ String mImage = getImageDir() + "Logo.jpg";
    private /*final*/ String mDocument = getMyDir() + "Reporting engine template - Data table.docx";

    @Test
    public void simpleCase() throws Exception
    {
        Document doc = DocumentHelper.createSimpleDocument("<<[s.Name]>> says: <<[s.Message]>>");

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
        buildReport(doc, sender, "s", ReportBuildOptions.INLINE_ERROR_MESSAGES);

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("LINQ Reporting Engine says: Hello World\f", doc.getText());
    }

    @Test
    public void stringFormat() throws Exception
    {
        Document doc = DocumentHelper.createSimpleDocument(
            "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
        buildReport(doc, sender, "s");

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f", doc.getText());
    }

    @Test
    public void numberFormat() throws Exception
    {
        Document doc = DocumentHelper.createSimpleDocument(
            "<<[s.Value1]:alphabetic>> : <<[s.Value2]:roman:lower>>, <<[s.Value3]:ordinal>>, <<[s.Value1]:ordinalText:upper>>" +
            ", <<[s.Value2]:cardinal>>, <<[s.Value3]:hex>>, <<[s.Value3]:arabicDash>>");

        NumericTestClass sender = new NumericTestBuilder()
            .withValuesAndDate(1, 2.2, 200, null, DateTime.parse("10.09.2016 10:00:00")).build();
        buildReport(doc, sender, "s");

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("A : ii, 200th, FIRST, Two, C8, - 200 -\f", doc.getText());
    }

    @Test
    public void testDataTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Data table.docx");

        buildReport(doc, Common.getContracts(), "Contracts");

        doc.save(getArtifactsDir() + "ReportingEngine.TestDataTable.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestDataTable.docx", getGoldsDir() + "ReportingEngine.TestDataTable Gold.docx"));
    }

    @Test
    public void total() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Total.docx");

        buildReport(doc, Common.getContracts(), "Contracts");

        doc.save(getArtifactsDir() + "ReportingEngine.Total.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.Total.docx", getGoldsDir() + "ReportingEngine.Total Gold.docx"));
    }

    @Test
    public void testNestedDataTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Nested data table.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.TestNestedDataTable.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestNestedDataTable.docx", getGoldsDir() + "ReportingEngine.TestNestedDataTable Gold.docx"));
    }

    @Test
    public void restartingListNumberingDynamically() throws Exception
    {
        Document template = new Document(getMyDir() + "Reporting engine template - List numbering.docx");

        buildReport(template, Common.getManagers(), "Managers", ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        template.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamically Gold.docx"));
    }

    @Test
    public void restartingListNumberingDynamicallyWhileInsertingDocumentDynamically() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.Document] -build>>");
        
        DocumentTestClass doc = new DocumentTestBuilder()
            .withDocument(new Document(getMyDir() + "Reporting engine template - List numbering.docx")).build();

        buildReport(template, new Object[] {doc, Common.getManagers()} , new String[] {"src", "Managers"}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        template.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
    }

    @Test
    public void restartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically() throws Exception
    {
        Document mainTemplate = DocumentHelper.createSimpleDocument("<<doc [src] -build>>");
        Document template1 = DocumentHelper.createSimpleDocument("<<doc [src1] -build>>");
        Document template2 = DocumentHelper.createSimpleDocument("<<doc [src2.Document] -build>>");
        
        DocumentTestClass doc = new DocumentTestBuilder()
            .withDocument(new Document(getMyDir() + "Reporting engine template - List numbering.docx")).build();

        buildReport(mainTemplate, new Object[] {template1, template2, doc, Common.getManagers()} , new String[] {"src", "src1", "src2", "Managers"}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        mainTemplate.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
     }

    @Test
    public void chartTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart.docx");

        buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.TestChart.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestChart.docx", getGoldsDir() + "ReportingEngine.TestChart Gold.docx"));
    }

    @Test
    public void bubbleChartTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Bubble chart.docx");

        buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.TestBubbleChart.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestBubbleChart.docx", getGoldsDir() + "ReportingEngine.TestBubbleChart Gold.docx"));
    }

    @Test
    public void setChartSeriesColorsDynamically() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart series color.docx");

        buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SetChartSeriesColorDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SetChartSeriesColorDynamically.docx", getGoldsDir() + "ReportingEngine.SetChartSeriesColorDynamically Gold.docx"));
    }

    @Test
    public void setPointColorsDynamically() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Point color.docx");

        ArrayList<ColorItemTestClass> colors = new ArrayList<ColorItemTestClass>();
        {
            colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Black", Color.BLACK.getRGB(), 1.0, 2.5, 3.5).build());
            colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Red", Color.RED.getRGB(), 2.0, 4.0, 2.5).build());
            colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Green", msColor.getGreen().getRGB(), 0.5, 1.5, 2.5).build());
            colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Blue", Color.BLUE.getRGB(), 4.5, 3.5, 1.5).build());
            colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Yellow", Color.YELLOW.getRGB(), 5.0, 2.5, 1.5)
                .build());
        }

        buildReport(doc, colors, "colorItems", new Class[] { ColorItemTestClass.class });

        doc.save(getArtifactsDir() + "ReportingEngine.SetPointColorDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SetPointColorDynamically.docx", getGoldsDir() + "ReportingEngine.SetPointColorDynamically Gold.docx"));
    }

    @Test
    public void conditionalExpressionForLeaveChartSeries() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart series.docx");
        
        int condition = 3;
        buildReport(doc, new Object[] { Common.getManagers(), condition }, new String[] { "managers", "condition" });

        doc.save(getArtifactsDir() + "ReportingEngine.TestLeaveChartSeries.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestLeaveChartSeries.docx", getGoldsDir() + "ReportingEngine.TestLeaveChartSeries Gold.docx"));
    }

    @Test (enabled = false, description = "WORDSNET-20810")
    public void conditionalExpressionForRemoveChartSeries() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart series.docx");
        
        int condition = 2;
        buildReport(doc, new Object[] { Common.getManagers(), condition }, new String[] { "managers", "condition" });

        doc.save(getArtifactsDir() + "ReportingEngine.TestRemoveChartSeries.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestRemoveChartSeries.docx", getGoldsDir() + "ReportingEngine.TestRemoveChartSeries Gold.docx"));
    }

    @Test
    public void indexOf() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Index of.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("The names are: John Smith, Tony Anderson, July James\f", doc.getText());
    }

    @Test
    public void ifElse() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - If-else.docx");

        buildReport(doc, Common.getManagers(), "m");

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("You have chosen 3 item(s).\f", doc.getText());
    }

    @Test
    public void ifElseWithoutData() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - If-else.docx");

        buildReport(doc, Common.getEmptyManagers(), "m");

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("You have chosen no items.\f", doc.getText());
    }

    @Test
    public void extensionMethods() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Extension methods.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.ExtensionMethods.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.ExtensionMethods.docx", getGoldsDir() + "ReportingEngine.ExtensionMethods Gold.docx"));
    }

    @Test
    public void operators() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Operators.docx");

        NumericTestClass testData = new NumericTestBuilder().withValuesAndLogical(1, 2.0, 3, null, true).build();

        ReportingEngine report = new ReportingEngine();
        report.getKnownTypes().add(NumericTestBuilder.class);
        report.buildReport(doc, testData, "ds");

        doc.save(getArtifactsDir() + "ReportingEngine.Operators.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.Operators.docx", getGoldsDir() + "ReportingEngine.Operators Gold.docx"));
    }

    @Test
    public void contextualObjectMemberAccess() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Contextual object member access.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.ContextualObjectMemberAccess.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.ContextualObjectMemberAccess.docx", getGoldsDir() + "ReportingEngine.ContextualObjectMemberAccess Gold.docx"));
    }

    @Test
    public void insertDocumentDynamicallyWithAdditionalTemplateChecking() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.Document] -build>>");

        DocumentTestClass doc = new DocumentTestBuilder()
            .withDocument(new Document(getMyDir() + "Reporting engine template - Data table.docx")).build();

        buildReport(template, new Object[] { doc, Common.getContracts() }, new String[] { "src", "Contracts" }, 
            ReportBuildOptions.NONE);
        template.save(
            getArtifactsDir() + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking.docx");

        msAssert.isTrue(
            DocumentHelper.compareDocs(
                getArtifactsDir() + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking.docx",
                getGoldsDir() + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking Gold.docx"),
            "Fail inserting document by document");
    }

    @Test
    public void insertDocumentDynamicallyWithStyles() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.Document] -sourceStyles>>");

        DocumentTestClass doc = new DocumentTestBuilder()
            .withDocument(new Document(getMyDir() + "Reporting engine template - Data table.docx")).build();

        buildReport(template, doc, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by document");
    }

    @Test
    public void insertDocumentDynamicallyByStream() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentStream]>>");

        DocumentTestClass docStream = new DocumentTestBuilder()
            .withDocumentStream(new FileStream(mDocument, FileMode.OPEN, FileAccess.READ)).build();

        buildReport(template, docStream, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by stream");
    }

    @Test
    public void insertDocumentDynamicallyByBytes() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentBytes]>>");

        DocumentTestClass docBytes = new DocumentTestBuilder()
            .withDocumentBytes(File.readAllBytes(getMyDir() + "Reporting engine template - Data table.docx")).build();

        buildReport(template, docBytes, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertDocumentDynamicallyByUri() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentString]>>");

        DocumentTestClass docUri = new DocumentTestBuilder()
            .withDocumentString("http://www.snee.com/xml/xslt/sample.doc").build();

        buildReport(template, docUri, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDynamically(uri) Gold.docx"), "Fail inserting document by uri");
    }

    @Test
    public void insertDocumentDynamicallyByBase64() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentString]>>");
        String base64Template = File.readAllText(getMyDir() + "Reporting engine template - Data table (base64).txt");

        DocumentTestClass docBase64 = new DocumentTestBuilder().withDocumentString(base64Template).build();

        buildReport(template, docBase64, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by uri");
    }

    @Test
    public void insertImageDynamically() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.Image]>>", ShapeType.TEXT_BOX);
        
                ImageTestClass image = new ImageTestBuilder().withImage(BufferedImage.FromFile(mImage, true)).build();
                            
        buildReport(template, image, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx", getGoldsDir() + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDynamicallyByStream() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageStream]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageStream = new ImageTestBuilder()
            .withImageStream(new FileStream(mImage, FileMode.OPEN, FileAccess.READ)).build();

        buildReport(template, imageStream, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx", getGoldsDir() + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDynamicallyByBytes() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageBytes]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageBytes = new ImageTestBuilder().withImageBytes(File.readAllBytes(mImage)).build();

        buildReport(template, imageBytes, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx", getGoldsDir() + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDynamicallyByUri() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageString]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageUri = new ImageTestBuilder()
            .withImageString(
                "http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png")
            .build();

        buildReport(template, imageUri, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");

        msAssert.isTrue(
            DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx",
                getGoldsDir() + "ReportingEngine.InsertImageDynamically(uri) Gold.docx"),
            "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDynamicallyByBase64() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageString]>>", ShapeType.TEXT_BOX);
        String base64Template = File.readAllText(getMyDir() + "Reporting engine template - base64 image.txt");

        ImageTestClass imageBase64 = new ImageTestBuilder().withImageString(base64Template).build();

        buildReport(template, imageBase64, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");

        msAssert.isTrue(
            DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx",
                getGoldsDir() + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"),
            "Fail inserting document by bytes");

    }
    
    @Test
    public void dynamicStretchingImageWithinTextBox() throws Exception
    {
        Document template = new Document(getMyDir() + "Reporting engine template - Dynamic stretching.docx");
        
        ImageTestClass image = new ImageTestBuilder().withImage(BufferedImage.FromFile(mImage, true)).build();
        buildReport(template, image, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.DynamicStretchingImageWithinTextBox.docx");

        Assert.assertTrue(
            DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.DynamicStretchingImageWithinTextBox.docx",
                getGoldsDir() + "ReportingEngine.DynamicStretchingImageWithinTextBox Gold.docx"));
    }

    @Test (dataProvider = "insertHyperlinksDynamicallyDataProvider")
    public void insertHyperlinksDynamically(String link) throws Exception
    {
        Document template = new Document(getMyDir() + "Reporting engine template - Inserting hyperlinks.docx");
        buildReport(template, 
            new Object[]
            {
                link, // Use URI or the name of a bookmark within the same document for a hyperlink
                "Aspose"
            },
            new String[]
            {
                "uri_or_bookmark_expression", 
                "display_text_expression"
            });

        template.save(getArtifactsDir() + "ReportingEngine.InsertHyperlinksDynamically.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "insertHyperlinksDynamicallyDataProvider")
	public static Object[][] insertHyperlinksDynamicallyDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"https://auckland.dynabic.com/wiki/display/org/Supported+dynamic+insertion+of+hyperlinks+for+LINQ+Reporting+Engine"},
			{"Bookmark"},
		};
	}

    @Test
    public void insertBookmarksDynamically() throws Exception
    {
        Document doc =
            DocumentHelper.createSimpleDocument(
                "<<bookmark [bookmark_expression]>><<foreach [m in Contracts]>><<[m.Client.Name]>><</foreach>><</bookmark>>");

        buildReport(doc, new Object[] { "BookmarkOne", Common.getContracts() },
            new String[] { "bookmark_expression", "Contracts" });

        doc.save(getArtifactsDir() + "ReportingEngine.InsertBookmarksDynamically.docx");
    }

    @Test
    public void withoutKnownType() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("<<[new DateTime()]:”dd.MM.yyyy”>>");

        ReportingEngine engine = new ReportingEngine();
        Assert.That(() => engine.buildReport(doc, ""), Throws.<IllegalStateException>TypeOf());
    }

    @Test
    public void workWithKnownTypes() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("<<[new DateTime(2016, 1, 20)]:”dd.MM.yyyy”>>");
        builder.writeln("<<[new DateTime(2016, 1, 20)]:”dd”>>");
        builder.writeln("<<[new DateTime(2016, 1, 20)]:”MM”>>");
        builder.writeln("<<[new DateTime(2016, 1, 20)]:”yyyy”>>");
        builder.writeln("<<[new DateTime(2016, 1, 20).Month]>>");

        buildReport(doc, "", new Class[]{ DateTime.class });

        doc.save(getArtifactsDir() + "ReportingEngine.KnownTypes.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.KnownTypes.docx", getGoldsDir() + "ReportingEngine.KnownTypes Gold.docx"));
    }

    @Test
    public void workWithContentControls() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - CheckBox Content Control.docx");
        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.WorkWithContentControls.docx");
    }

    @Test
    public void workWithSingleColumnTableRow() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row.docx");
        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRow.docx");
    }

    @Test
    public void workWithSingleColumnTableRowGreedy() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row greedy.docx");
        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRowGreedy.docx");
    }

    @Test
    public void tableRowConditionalBlocks() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row conditional blocks.docx");

        ArrayList<ClientTestClass> clients = new ArrayList<ClientTestClass>();
        {
            clients.add(new ClientTestClass());
                {
                    ((ClientTestClass)clients.get(0)).setName("John Monrou");
                    ((ClientTestClass)clients.get(0)).setCountry("France");
                    ((ClientTestClass)clients.get(0)).setLocalAddress("27 RUE PASTEUR");
                }
            clients.add(new ClientTestClass());
                {
                    ((ClientTestClass)clients.get(1)).setName("James White");
                    ((ClientTestClass)clients.get(1)).setCountry("England");
                    ((ClientTestClass)clients.get(1)).setLocalAddress("14 Tottenham Court Road");
                }
            clients.add(new ClientTestClass());
                {
                    ((ClientTestClass)clients.get(2)).setName("Kate Otts");
                    ((ClientTestClass)clients.get(2)).setCountry("New Zealand");
                    ((ClientTestClass)clients.get(2)).setLocalAddress("Wellington 6004");
                }
        }

        buildReport(doc, clients, "clients");

        doc.save(getArtifactsDir() + "ReportingEngine.TableRowConditionalBlocks.docx");
    }

    @Test
    public void ifGreedy() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - If greedy.docx");

        AsposeData obj = new AsposeData();
        {
            obj.setList(new ArrayList<String>());
                {
                    obj.getList().add("abc");
                }
        }

        buildReport(doc, obj);

        doc.save(getArtifactsDir() + "ReportingEngine.IfGreedy.docx");
    }

    public static class AsposeData
    {
        public ArrayList<String> getList() { return mList; }; public void setList(ArrayList<String> value) { mList = value; };

        private ArrayList<String> mList;
    }

    @Test
    public void stretchImagefitHeight() throws Exception
    {
        Document doc =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitHeight>>",
                ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
            .withImageStream(new FileStream(mImage, FileMode.OPEN, FileAccess.READ)).build();
        buildReport(doc, imageStream, "src", ReportBuildOptions.NONE);

        doc = DocumentHelper.saveOpen(doc);

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that the width is preserved, and the height is changed
            Assert.assertNotEquals(346.35, shape.getHeight());
            Assert.assertEquals(431.5, shape.getWidth());
        }
    }

    @Test
    public void stretchImagefitWidth() throws Exception
    {
        Document doc =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitWidth>>",
                ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
            .withImageStream(new FileStream(mImage, FileMode.OPEN, FileAccess.READ)).build();
        buildReport(doc, imageStream, "src", ReportBuildOptions.NONE);

        doc = DocumentHelper.saveOpen(doc);

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that the height is preserved, and the width is changed
            Assert.assertNotEquals(431.5, shape.getWidth());
            Assert.assertEquals(346.35, shape.getHeight());
        }
    }

    @Test
    public void stretchImagefitSize() throws Exception
    {
        Document doc =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitSize>>",
                ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
            .withImageStream(new FileStream(mImage, FileMode.OPEN, FileAccess.READ)).build();
        buildReport(doc, imageStream, "src", ReportBuildOptions.NONE);

        doc = DocumentHelper.saveOpen(doc);

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that the height and the width are changed
            Assert.assertNotEquals(346.35, shape.getHeight());
            Assert.assertNotEquals(431.5, shape.getWidth());
        }
    }

    @Test
    public void stretchImagefitSizeLim() throws Exception
    {
        Document doc =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitSizeLim>>",
                ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
            .withImageStream(new FileStream(mImage, FileMode.OPEN, FileAccess.READ)).build();
        buildReport(doc, imageStream, "src", ReportBuildOptions.NONE);

        doc = DocumentHelper.saveOpen(doc);

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that textbox size are equal image size
            Assert.assertEquals(300.0d, shape.getHeight());
            Assert.assertEquals(300.0d, shape.getWidth());
        }
    }

    @Test
    public void withoutMissingMembers() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        //Add templete to the document for reporting engine
        DocumentHelper.insertBuilderText(builder,
            new String[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

        //Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
        Assert.That(() => buildReport(builder.getDocument(), new DataSet(), "", ReportBuildOptions.NONE),
            Throws.<IllegalStateException>TypeOf());
    }

    @Test
    public void withMissingMembers() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        //Add templete to the document for reporting engine
        DocumentHelper.insertBuilderText(builder,
            new String[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

        buildReport(builder.getDocument(), new DataSet(), "", ReportBuildOptions.ALLOW_MISSING_MEMBERS);

        //Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
        Assert.assertEquals(ControlChar.PARAGRAPH_BREAK + ControlChar.PARAGRAPH_BREAK + ControlChar.SECTION_BREAK,
            builder.getDocument().getText());
    }

    @Test (dataProvider = "inlineErrorMessagesDataProvider")
    public void inlineErrorMessages(String templateText, String result) throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();
        DocumentHelper.insertBuilderText(builder, new String[] { templateText });
        
        buildReport(builder.getDocument(), new DataSet(), "", ReportBuildOptions.INLINE_ERROR_MESSAGES);

        Assert.That(msString.trimEnd(builder.getDocument().getFirstSection().getBody().getParagraphs().get(0).getText()), Is.EqualTo(result));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "inlineErrorMessagesDataProvider")
	public static Object[][] inlineErrorMessagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"<<[missingObject.First().id]>>",  "<<[missingObject.First( Error! Can not get the value of member 'missingObject' on type 'System.Data.DataSet'. ).id]>>"},
			{"<<[new DateTime()]:\"dd.MM.yyyy\">>",  "<<[new DateTime( Error! A type identifier is expected. )]:\"dd.MM.yyyy\">>"},
			{"<<]>>",  "<<] Error! Character ']' is unexpected. >>"},
			{"<<[>>",  "<<[>> Error! An expression is expected."},
			{"<<>>",  "<<>> Error! Tag end is unexpected."},
		};
	}

    @Test
    public void setBackgroundColor() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Background color.docx");

        ArrayList<ColorItemTestClass> colors = new ArrayList<ColorItemTestClass>();
        {
            colors.add(new ColorItemTestBuilder().withColor("Black", Color.BLACK).build());
            colors.add(new ColorItemTestBuilder().withColor("Red", new Color((255), (0), (0))).build());
            colors.add(new ColorItemTestBuilder().withColor("Empty", msColor.Empty).build());
        }

        buildReport(doc, colors, "Colors");

        doc.save(getArtifactsDir() + "ReportingEngine.BackColor.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.BackColor.docx",
            getGoldsDir() + "ReportingEngine.BackColor Gold.docx"));
    }

    @Test
    public void doNotRemoveEmptyParagraphs() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Remove empty paragraphs.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx",
            getGoldsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs Gold.docx"));
    }

    @Test
    public void removeEmptyParagraphs() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Remove empty paragraphs.docx");

        buildReport(doc, Common.getManagers(), "Managers", ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        doc.save(getArtifactsDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RemoveEmptyParagraphs.docx",
            getGoldsDir() + "ReportingEngine.RemoveEmptyParagraphs Gold.docx"));
    }

    @Test (dataProvider = "mergingTableCellsDynamicallyDataProvider")
    public void mergingTableCellsDynamically(String value1, String value2, String resultDocumentName) throws Exception
    {
        String artifactPath = getArtifactsDir() + resultDocumentName +
                               FileFormatUtil.saveFormatToExtension(SaveFormat.DOCX);
        String goldPath = getGoldsDir() + resultDocumentName + " Gold" +
                          FileFormatUtil.saveFormatToExtension(SaveFormat.DOCX);
        
        Document doc = new Document(getMyDir() + "Reporting engine template - Merging table cells dynamically.docx");

        ArrayList<ClientTestClass> clients = new ArrayList<ClientTestClass>();
        {
            clients.add(new ClientTestClass());
                {
                    ((ClientTestClass)clients.get(0)).setName("John Monrou");
                    ((ClientTestClass)clients.get(0)).setCountry("France");
                    ((ClientTestClass)clients.get(0)).setLocalAddress("27 RUE PASTEUR");
                }
            clients.add(new ClientTestClass());
                {
                    ((ClientTestClass)clients.get(1)).setName("James White");
                    ((ClientTestClass)clients.get(1)).setCountry("New Zealand");
                    ((ClientTestClass)clients.get(1)).setLocalAddress("14 Tottenham Court Road");
                }
            clients.add(new ClientTestClass());
                {
                    ((ClientTestClass)clients.get(2)).setName("Kate Otts");
                    ((ClientTestClass)clients.get(2)).setCountry("New Zealand");
                    ((ClientTestClass)clients.get(2)).setLocalAddress("Wellington 6004");
                }
        }

        buildReport(doc, new Object[] { value1, value2, clients }, new String[] { "value1", "value2", "clients" });
        doc.save(artifactPath);

        Assert.assertTrue(DocumentHelper.compareDocs(artifactPath, goldPath));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "mergingTableCellsDynamicallyDataProvider")
	public static Object[][] mergingTableCellsDynamicallyDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"Hello",  "Hello",  "ReportingEngine.MergingTableCellsDynamically.Merged"},
			{"Hello",  "Name",  "ReportingEngine.MergingTableCellsDynamically.NotMerged"},
		};
	}

    @Test
    public void xmlDataStringWithoutSchema() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - XML data destination.docx");

        XmlDataSource dataSource = new XmlDataSource(getMyDir() + "List of people.xml");
        buildReport(doc, dataSource, "persons");

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataString.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataString.docx",
            getGoldsDir() + "ReportingEngine.DataSource Gold.docx"));
    }

    @Test
    public void xmlDataStreamWithoutSchema() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - XML data destination.docx");

        FileStream stream = new FileInputStream(getMyDir() + "List of people.xml");
        try /*JAVA: was using*/
        {
            XmlDataSource dataSource = new XmlDataSource(stream);
            buildReport(doc, dataSource, "persons");
        }
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataStream.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataStream.docx",
            getGoldsDir() + "ReportingEngine.DataSource Gold.docx"));
    }

    @Test
    public void xmlDataWithNestedElements() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Data destination with nested elements.docx");

        XmlDataSource dataSource = new XmlDataSource(getMyDir() + "Nested elements.xml");
        buildReport(doc, dataSource, "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataWithNestedElements.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataWithNestedElements.docx",
            getGoldsDir() + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
    }

    @Test
    public void jsonDataString() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - JSON data destination.docx");

        JsonDataLoadOptions options = new JsonDataLoadOptions();
        {
            options.setExactDateTimeParseFormats(new ArrayList<String>()); {options.getExactDateTimeParseFormats().add("MM/dd/yyyy"); options.getExactDateTimeParseFormats().add("MM.d.yy"); options.getExactDateTimeParseFormats().add("MM d yy");}
        }

        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "List of people.json", options);
        buildReport(doc, dataSource, "persons");
        
        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataString.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.JsonDataString.docx",
            getGoldsDir() + "ReportingEngine.JsonDataString Gold.docx"));
    }

    @Test
    public void jsonDataStringException() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - JSON data destination.docx");

        JsonDataLoadOptions options = new JsonDataLoadOptions();
        options.setSimpleValueParseMode(JsonSimpleValueParseMode.STRICT);
        
        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "List of people.json", options);
        Assert.<IllegalStateException>Throws(() => buildReport(doc, dataSource, "persons"));
    }

    @Test
    public void jsonDataStream() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - JSON data destination.docx");

        JsonDataLoadOptions options = new JsonDataLoadOptions();
        {
            options.setExactDateTimeParseFormats(new ArrayList<String>()); {options.getExactDateTimeParseFormats().add("MM/dd/yyyy"); options.getExactDateTimeParseFormats().add("MM.d.yy"); options.getExactDateTimeParseFormats().add("MM d yy");}
        }

        FileStream stream = new FileInputStream(getMyDir() + "List of people.json");
        try /*JAVA: was using*/
        {
            JsonDataSource dataSource = new JsonDataSource(stream, options);
            buildReport(doc, dataSource, "persons");
        }
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataStream.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.JsonDataStream.docx",
            getGoldsDir() + "ReportingEngine.JsonDataString Gold.docx"));
    }

    @Test
    public void jsonDataWithNestedElements() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Data destination with nested elements.docx");

        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "Nested elements.json");
        buildReport(doc, dataSource, "managers");
        
        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataWithNestedElements.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.JsonDataWithNestedElements.docx",
            getGoldsDir() + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
    }

    @Test
    public void csvDataString() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - CSV data destination.docx");
        
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.setDelimiter(';');
        loadOptions.setCommentChar('$');

        CsvDataSource dataSource = new CsvDataSource(getMyDir() + "List of people.csv", loadOptions);
        buildReport(doc, dataSource, "persons");
        
        doc.save(getArtifactsDir() + "ReportingEngine.CsvDataString.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.CsvDataString.docx",
            getGoldsDir() + "ReportingEngine.CsvData Gold.docx"));
    }

    @Test
    public void csvDataStream() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - CSV data destination.docx");
        
        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.setDelimiter(';');
        loadOptions.setCommentChar('$');

        FileStream stream = new FileInputStream(getMyDir() + "List of people.csv");
        try /*JAVA: was using*/
        {
            CsvDataSource dataSource = new CsvDataSource(stream, loadOptions);
            buildReport(doc, dataSource, "persons");
        }
        finally { if (stream != null) stream.close(); }
        
        doc.save(getArtifactsDir() + "ReportingEngine.CsvDataStream.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.CsvDataStream.docx",
            getGoldsDir() + "ReportingEngine.CsvData Gold.docx"));
    }

    @Test (dataProvider = "insertComboboxDropdownListItemsDynamicallyDataProvider")
    public void insertComboboxDropdownListItemsDynamically(/*SdtType*/int sdtType) throws Exception
    {
        final String TEMPLATE =
            "<<item[\"three\"] [\"3\"]>><<if [false]>><<item [\"four\"] [null]>><</if>><<item[\"five\"] [\"5\"]>>";

        SdtListItem[] staticItems =
        {
            new SdtListItem("1", "one"),
            new SdtListItem("2", "two")
        };

        Document doc = new Document();

        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, sdtType, MarkupLevel.BLOCK); { sdt.setTitle(TEMPLATE); }

        for (SdtListItem item : staticItems)
        {
            sdt.getListItems().add(item);
        }

        doc.getFirstSection().getBody().appendChild(sdt);

        buildReport(doc, new Object(), "");

        doc.save(getArtifactsDir() + $"ReportingEngine.InsertComboboxDropdownListItemsDynamically_{sdtType}.docx");

        doc = new Document(getArtifactsDir() +
                           $"ReportingEngine.InsertComboboxDropdownListItemsDynamically_{sdtType}.docx");

        sdt = (StructuredDocumentTag) doc.getChild(NodeType.STRUCTURED_DOCUMENT_TAG, 0, true);

        SdtListItem[] expectedItems =
        {
            new SdtListItem("1", "one"),
            new SdtListItem("2", "two"),
            new SdtListItem("3", "three"),
            new SdtListItem("5", "five")
        };

        Assert.assertEquals(expectedItems.length, sdt.getListItems().getCount());

        for (int i = 0; i < expectedItems.length; i++)
        {
            Assert.assertEquals(expectedItems[i].getValue(), sdt.getListItems().get(i).getValue());
            Assert.assertEquals(expectedItems[i].getDisplayText(), sdt.getListItems().get(i).getDisplayText());
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "insertComboboxDropdownListItemsDynamicallyDataProvider")
	public static Object[][] insertComboboxDropdownListItemsDynamicallyDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SdtType.COMBO_BOX},
			{SdtType.DROP_DOWN_LIST},
		};
	}

    private static void buildReport(Document document, Object dataSource, String dataSourceName,
        /*ReportBuildOptions*/int reportBuildOptions) throws Exception
    {
        ReportingEngine engine = new ReportingEngine(); { engine.setOptions(reportBuildOptions); }
        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(Document document, Object[] dataSource, String[] dataSourceName) throws Exception
    {
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(Document document, Object[] dataSource, String[] dataSourceName,
        /*ReportBuildOptions*/int reportBuildOptions) throws Exception
    {
        ReportingEngine engine = new ReportingEngine(); { engine.setOptions(reportBuildOptions); }
        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(Document document, Object dataSource, String dataSourceName, Class[] knownTypes) throws Exception
    {
        ReportingEngine engine = new ReportingEngine();

        for (Class knownType : knownTypes)
        {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(Document document, Object dataSource) throws Exception
    {
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(document, dataSource);
    }

    private static void buildReport(Document document, Object dataSource, String dataSourceName) throws Exception
    {
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(Document document, Object dataSource, Class[] knownTypes) throws Exception
    {
        ReportingEngine engine = new ReportingEngine();

        for (Class knownType : knownTypes)
        {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource);
    }
}
