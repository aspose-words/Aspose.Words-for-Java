// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import ApiExamples.TestData.TestClasses.NumericTestClass;
import ApiExamples.TestData.TestBuilders.NumericTestBuilder;
import com.aspose.ms.System.DateTime;
import ApiExamples.TestData.Common;
import java.util.ArrayList;
import ApiExamples.TestData.TestClasses.ColorItemTestClass;
import ApiExamples.TestData.TestBuilders.ColorItemTestBuilder;
import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.ReportingEngine;
import ApiExamples.TestData.TestClasses.DocumentTestClass;
import ApiExamples.TestData.TestBuilders.DocumentTestBuilder;
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
import java.lang.Class;
import org.testng.annotations.DataProvider;


@Test
public class ExReportingEngine extends ApiExampleBase
{
    private /*final*/ String mImage = getImageDir() + "Test_636_852.gif";
    private /*final*/ String mDocument = getMyDir() + "ReportingEngine.TestDataTable.docx";

    @Test
    public void simpleCase() throws Exception
    {
        Document doc = DocumentHelper.createSimpleDocument("<<[s.Name]>> says: <<[s.Message]>>");

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
        buildReport(doc, sender, "s", ReportBuildOptions.INLINE_ERROR_MESSAGES);

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual("LINQ Reporting Engine says: Hello World\f", doc.getText());
    }

    @Test
    public void stringFormat() throws Exception
    {
        Document doc = DocumentHelper.createSimpleDocument(
            "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
        buildReport(doc, sender, "s");

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f", doc.getText());
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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual("A : ii, 200th, FIRST, Two, C8, - 200 -\f", doc.getText());
    }

    @Test
    public void dataTableTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestDataTable.docx");

        buildReport(doc, Common.getContracts(), "Contracts");

        doc.save(getArtifactsDir() + "ReportingEngine.TestDataTable.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestDataTable.docx", getGoldsDir() + "ReportingEngine.TestDataTable Gold.docx"));
    }

    @Test
    public void progressiveTotal() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.Total.docx");

        buildReport(doc, Common.getContracts(), "Contracts");

        doc.save(getArtifactsDir() + "ReportingEngine.Total.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.Total.docx", getGoldsDir() + "ReportingEngine.Total Gold.docx"));
    }

    @Test
    public void nestedDataTableTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestNestedDataTable.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.TestNestedDataTable.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestNestedDataTable.docx", getGoldsDir() + "ReportingEngine.TestNestedDataTable Gold.docx"));
    }

    @Test
    public void chartTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestChart.docx");

        buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.TestChart.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestChart.docx", getGoldsDir() + "ReportingEngine.TestChart Gold.docx"));
    }

    @Test
    public void bubbleChartTest() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestBubbleChart.docx");

        buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.TestBubbleChart.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestBubbleChart.docx", getGoldsDir() + "ReportingEngine.TestBubbleChart Gold.docx"));
    }

    @Test
    public void setChartSeriesColorsDynamically() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.SetChartSeriesColorDinamically.docx");

        buildReport(doc, Common.getManagers(), "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SetChartSeriesColorDinamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SetChartSeriesColorDinamically.docx", getGoldsDir() + "ReportingEngine.SetChartSeriesColorDinamically Gold.docx"));
    }

    @Test
    public void setPointColorsDynamically() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.SetPointColorDinamically.docx");

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

        doc.save(getArtifactsDir() + "ReportingEngine.SetPointColorDinamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SetPointColorDinamically.docx", getGoldsDir() + "ReportingEngine.SetPointColorDinamically Gold.docx"));
    }

    @Test
    public void conditionalExpressionForLeaveChartSeries() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestRemoveChartSeries.docx");

        doc.save(getArtifactsDir() + "ReportingEngine.TestLeaveChartSeries.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestLeaveChartSeries.docx", getGoldsDir() + "ReportingEngine.TestLeaveChartSeries Gold.docx"));
    }

    @Test
    public void conditionalExpressionForRemoveChartSeries() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestRemoveChartSeries.docx");

        doc.save(getArtifactsDir() + "ReportingEngine.TestRemoveChartSeries.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.TestRemoveChartSeries.docx", getGoldsDir() + "ReportingEngine.TestRemoveChartSeries Gold.docx"));
    }

    @Test
    public void indexOf() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestIndexOf.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual("The names are: John Smith, Tony Anderson, July James\f", doc.getText());
    }

    @Test
    public void ifElse() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.IfElse.docx");

        buildReport(doc, Common.getManagers(), "m");

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual("You have chosen 3 item(s).\f", doc.getText());
    }

    @Test
    public void ifElseWithoutData() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.IfElse.docx");

        buildReport(doc, Common.getEmptyManagers(), "m");

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        msAssert.areEqual("You have chosen no items.\f", doc.getText());
    }

    @Test
    public void extensionMethods() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.ExtensionMethods.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.ExtensionMethods.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.ExtensionMethods.docx", getGoldsDir() + "ReportingEngine.ExtensionMethods Gold.docx"));
    }

    @Test
    public void operators() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.Operators.docx");

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
        Document doc = new Document(getMyDir() + "ReportingEngine.ContextualObjectMemberAccess.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.ContextualObjectMemberAccess.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.ContextualObjectMemberAccess.docx", getGoldsDir() + "ReportingEngine.ContextualObjectMemberAccess Gold.docx"));
    }

    @Test
    public void insertDocumentDinamicallyWithAdditionalTemplateChecking() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.Document] -build>>");

        DocumentTestClass doc = new DocumentTestBuilder()
            .withDocument(new Document(getMyDir() + "ReportingEngine.TestDataTable.docx")).build();

        buildReport(template, new Object[] { doc, Common.getContracts() }, new String[] { "src", "Contracts" }, 
            ReportBuildOptions.NONE);
        template.save(
            getArtifactsDir() + "ReportingEngine.InsertDocumentDinamicallyWithAdditionalTemplateChecking.docx");

        msAssert.isTrue(
            DocumentHelper.compareDocs(
                getArtifactsDir() + "ReportingEngine.InsertDocumentDinamicallyWithAdditionalTemplateChecking.docx",
                getGoldsDir() + "ReportingEngine.InsertDocumentDinamicallyWithAdditionalTemplateChecking Gold.docx"),
            "Fail inserting document by document");
    }

    @Test
    public void insertDocumentDinamically() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.Document]>>");

        DocumentTestClass doc = new DocumentTestBuilder()
            .withDocument(new Document(getMyDir() + "ReportingEngine.TestDataTable.docx")).build();

        buildReport(template, doc, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by document");
    }

    @Test
    public void insertDocumentDinamicallyByStream() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentStream]>>");

        DocumentTestClass docStream = new DocumentTestBuilder()
            .withDocumentStream(new FileStream(mDocument, FileMode.OPEN, FileAccess.READ)).build();

        buildReport(template, docStream, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by stream");
    }

    @Test
    public void insertDocumentDinamicallyByBytes() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentBytes]>>");

        DocumentTestClass docBytes = new DocumentTestBuilder()
            .withDocumentBytes(File.readAllBytes(getMyDir() + "ReportingEngine.TestDataTable.docx")).build();

        buildReport(template, docBytes, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertDocumentDinamicallyByUri() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.DocumentUri]>>");

        DocumentTestClass docUri = new DocumentTestBuilder()
            .withDocumentUri("http://www.snee.com/xml/xslt/sample.doc").build();

        buildReport(template, docUri, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertDocumentDinamically.docx", getGoldsDir() + "ReportingEngine.InsertDocumentDinamically(uri) Gold.docx"), "Fail inserting document by uri");
    }

    @Test
    public void insertImageDinamically() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.Image]>>", ShapeType.TEXT_BOX);
        ImageTestClass image = new ImageTestBuilder().withImage(BufferedImage.FromFile(mImage, true)).build();
        buildReport(template, image, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx", getGoldsDir() + "ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDinamicallyByStream() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageStream]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageStream = new ImageTestBuilder()
            .withImageStream(new FileStream(mImage, FileMode.OPEN, FileAccess.READ)).build();

        buildReport(template, imageStream, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx", getGoldsDir() + "ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDinamicallyByBytes() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageBytes]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageBytes = new ImageTestBuilder().withImageBytes(File.readAllBytes(mImage)).build();

        buildReport(template, imageBytes, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx");

        msAssert.isTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx", getGoldsDir() + "ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
    }

    @Test
    public void insertImageDinamicallyByUri() throws Exception
    {
        Document template =
            DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.ImageUri]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageUri = new ImageTestBuilder()
            .withImageUri(
                "http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png")
            .build();

        buildReport(template, imageUri, "src", ReportBuildOptions.NONE);
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx");

        msAssert.isTrue(
            DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertImageDinamically.docx",
                getGoldsDir() + "ReportingEngine.InsertImageDinamically(uri) Gold.docx"),
            "Fail inserting document by bytes");
    }

    @Test
    public void insertHyperlinksDinamically() throws Exception
    {
        Document template = new Document(getMyDir() + "ReportingEngine.InsertingHyperlinks.docx");
        buildReport(template, 
            new Object[]
            {
                "https://auckland.dynabic.com/wiki/display/org/Supported+dynamic+insertion+of+hyperlinks+for+LINQ+Reporting+Engine",
                "Aspose"
            },
            new String[]
            {
                "uri_expression", 
                "display_text_expression"
            });

        template.save(getArtifactsDir() + "ReportingEngine.InsertHyperlinksDinamically.docx");

        msAssert.isTrue(
            DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.InsertHyperlinksDinamically.docx",
                getGoldsDir() + "ReportingEngine.InsertHyperlinksDinamically Gold.docx"),
            "Fail inserting document by bytes");
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
    public void workWithSingleColumnTableRow() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.SingleColumnTableRow.docx");
        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRow.docx");
    }

    @Test
    public void workWithSingleColumnTableRowGreedy() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.SingleColumnTableRowGreedy.docx");
        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRowGreedy.docx");
    }

    @Test
    public void tableRowConditionalBlocks() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.TableRowConditionalBlocks.docx");

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
        Document doc = new Document(getMyDir() + "ReportingEngine.IfGreedy.docx");

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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        doc = new Document(dstStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that width is keeped and height is changed
            msAssert.areNotEqual(346.35, shape.getHeight());
            msAssert.areEqual(431.5, shape.getWidth());
        }

        dstStream.dispose();
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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        doc = new Document(dstStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            // Assert that the image is really insert in textbox and 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that height is keeped and width is changed
            msAssert.areNotEqual(431.5, shape.getWidth());
            msAssert.areEqual(346.35, shape.getHeight());
        }

        dstStream.dispose();
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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        doc = new Document(dstStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that height is changed and width is changed
            msAssert.areNotEqual(346.35, shape.getHeight());
            msAssert.areNotEqual(431.5, shape.getWidth());
        }

        dstStream.dispose();
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

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        doc = new Document(dstStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Shape shape : shapes.<Shape>OfType() !!Autoporter error: Undefined expression type )
        {
            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that textbox size are equal image size
            msAssert.areEqual(346.35, shape.getHeight());
            msAssert.areEqual(258.54, shape.getWidth());
        }

        dstStream.dispose();
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
        msAssert.areEqual(ControlChar.PARAGRAPH_BREAK + ControlChar.PARAGRAPH_BREAK + ControlChar.SECTION_BREAK,
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
        Document doc = new Document(getMyDir() + "ReportingEngine.BackColor.docx");

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
        Document doc = new Document(getMyDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");

        buildReport(doc, Common.getManagers(), "Managers");

        doc.save(getArtifactsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx",
            getGoldsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs Gold.docx"));
    }

    @Test
    public void removeEmptyParagraphs() throws Exception
    {
        Document doc = new Document(getMyDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");

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
        
        Document doc = new Document(getMyDir() + "ReportingEngine.MergingTableCellsDynamically.docx");

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
