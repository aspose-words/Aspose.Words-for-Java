package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.Common;
import TestData.TestBuilders.ColorItemTestBuilder;
import TestData.TestBuilders.DocumentTestBuilder;
import TestData.TestBuilders.ImageTestBuilder;
import TestData.TestBuilders.NumericTestBuilder;
import TestData.TestClasses.*;
import com.aspose.words.Shape;
import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataSet;
import org.apache.commons.io.FileUtils;
import org.omg.CORBA.Environment;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.time.LocalDate;
import java.util.*;
import java.util.List;

@Test
public class ExReportingEngine extends ApiExampleBase {
    private final String mImage = getImageDir() + "Logo.jpg";
    private final String mDocument = getMyDir() + "Reporting engine template - Data table (Java).docx";

    @Test
    public void simpleCase() throws Exception {
        Document doc = DocumentHelper.createSimpleDocument("<<[s.getName()]>> says: <<[s.getMessage()]>>");

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
        buildReport(doc, sender, "s", ReportBuildOptions.INLINE_ERROR_MESSAGES);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals(doc.getText(), "LINQ Reporting Engine says: Hello World\f");
    }

    @Test
    public void stringFormat() throws Exception {
        Document doc = DocumentHelper.createSimpleDocument(
                "<<[s.getName()]:lower>> says: <<[s.getMessage()]:upper>>, <<[s.getMessage()]:caps>>, <<[s.getMessage()]:firstCap>>");

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
        buildReport(doc, sender, "s");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals(doc.getText(), "linq reporting engine says: HELLO WORLD, Hello World, Hello world\f");
    }

    @Test
    public void numberFormat() throws Exception {
        Document doc = DocumentHelper.createSimpleDocument(
                "<<[s.getValue1()]:alphabetic>> : <<[s.getValue2()]:roman:lower>>, <<[s.getValue3()]:ordinal>>, <<[s.getValue1()]:ordinalText:upper>>"
                        + ", <<[s.getValue2()]:cardinal>>, <<[s.getValue3()]:hex>>, <<[s.getValue3()]:arabicDash>>");

        NumericTestClass sender = new NumericTestBuilder()
                .withValuesAndDate(1, 2.2, 200, null, LocalDate.of(2016, 9, 10)).build();
        buildReport(doc, sender, "s");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals(doc.getText(), "A : ii, 200th, FIRST, Two, C8, - 200 -\f");
    }

    @Test
    public void dataTableTest() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Data table (Java).docx");

        buildReport(doc, Common.getContracts(), "Contracts", new Class[]{ContractTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestDataTable.docx");
    }

    @Test
    public void progressiveTotal() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Total (Java).docx");

        buildReport(doc, Common.getContracts(), "Contracts", new Class[]{ContractTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.Total.docx");
    }

    @Test
    public void nestedDataTableTest() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Nested data table (Java).docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class, ContractTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestNestedDataTable.docx");
    }

    @Test
    public void restartingListNumberingDynamically() throws Exception {
        Document template = new Document(getMyDir() + "Reporting engine template - List numbering (Java).docx");

        buildReport(template, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class, ContractTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        template.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamically Gold.docx"));
    }

    @Test
    public void restartingListNumberingDynamicallyWhileInsertingDocumentDynamically() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()] -build>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(getMyDir() + "Reporting engine template - List numbering (Java).docx")).build();

        buildReport(template, new Object[]{doc, Common.getManagers()}, new String[]{"src", "Managers"}, new Class[]{ManagerTestClass.class, ContractTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        template.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
    }

    @Test
    public void restartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically() throws Exception {
        Document mainTemplate = DocumentHelper.createSimpleDocument("<<doc [src] -build>>");
        Document template1 = DocumentHelper.createSimpleDocument("<<doc [src1] -build>>");
        Document template2 = DocumentHelper.createSimpleDocument("<<doc [src2.getDocument()] -build>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(getMyDir() + "Reporting engine template - List numbering (Java).docx")).build();

        buildReport(mainTemplate, new Object[]{template1, template2, doc, Common.getManagers()}, new String[]{"src", "src1", "src2", "Managers"}, new Class[]{ManagerTestClass.class, ContractTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        mainTemplate.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
    }

    @Test
    public void chartTest() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart (Java).docx");

        buildReport(doc, Common.getManagers(), "managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestChart.docx");
    }

    @Test
    public void bubbleChartTest() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Bubble chart (Java).docx");

        buildReport(doc, Common.getManagers(), "managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestBubbleChart.docx");
    }

    @Test
    public void setChartSeriesColorsDynamically() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart series color (Java).docx");

        buildReport(doc, Common.getManagers(), "managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SetChartSeriesColorDynamically.docx");
    }

    @Test
    public void setPointColorsDynamically() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Point color (Java).docx");

        List<ColorItemTestClass> colors = new ArrayList<>();
        colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Black", Color.BLACK.getRGB(), 1.0, 2.5, 3.5).build());
        colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Red", Color.RED.getRGB(), 2.0, 4.0, 2.5).build());
        colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Green", Color.GREEN.getRGB(), 0.5, 1.5, 2.5).build());
        colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Blue", Color.BLUE.getRGB(), 4.5, 3.5, 1.5).build());
        colors.add(new ColorItemTestBuilder().withColorCodeAndValues("Yellow", Color.YELLOW.getRGB(), 5.0, 2.5, 1.5).build());

        buildReport(doc, colors, "colorItems", new Class[]{ColorItemTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SetPointColorDynamically.docx");
    }

    @Test(enabled = false, description = "WORDSNET-20810")
    public void conditionalExpressionRemoveChartSeries() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Chart series (Java)");

        int condition = 2;
        buildReport(doc, new Object[]{Common.getManagers(), condition}, new String[]{"managers", "condition"}, new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestRemoveChartSeries.docx");
    }

    @Test
    public void indexOf() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Index of (Java).docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals("The names are: John Smith, Tony Anderson, July James\f", doc.getText());
    }

    @Test
    public void ifElse() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - If-else (Java).docx");

        buildReport(doc, Common.getManagers(), "m", new Class[]{ManagerTestClass.class});

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals("You have chosen 3 item(s).\f", doc.getText());
    }

    @Test
    public void ifElseWithoutData() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - If-else (Java).docx");

        buildReport(doc, Common.getEmptyManagers(), "m", new Class[]{ManagerTestClass.class});

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals("You have chosen no items.\f", doc.getText());
    }

    @Test
    public void extensionMethods() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Extension methods (Java).docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});
        doc.save(getArtifactsDir() + "ReportingEngine.ExtensionMethods.docx");
    }

    @Test
    public void operators() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Operators (Java).docx");

        NumericTestClass testData = new NumericTestBuilder().withValuesAndLogical(1, 2.0, 3, null, true).build();

        buildReport(doc, testData, "ds", new Class[]{NumericTestBuilder.class});
        doc.save(getArtifactsDir() + "ReportingEngine.Operators.docx");
    }

    @Test
    public void headerVariable() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Header variable (Java).docx");

        buildReport(doc, new DataSet(), "", ReportBuildOptions.USE_LEGACY_HEADER_FOOTER_VISITING);

        doc.save(getArtifactsDir() + "ReportingEngine.HeaderVariable.docx");

        Assert.assertEquals("Value of myHeaderVariable is: I am header variable", doc.getFirstSection().getBody().getFirstParagraph().getText().trim());
    }

    @Test
    public void contextualObjectMemberAccess() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Contextual object member access (Java).docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.ContextualObjectMemberAccess.docx");
    }

    @Test
    public void insertDocumentDynamicallyWithAdditionalTemplateChecking() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()] -build>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(mDocument)).build();

        buildReport(template, new Object[]{doc, Common.getContracts()}, new String[]{"src", "Contracts"}, new Class[]{ContractTestClass.class});
        template.save(
                getArtifactsDir() + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking.docx");
    }

    @Test
    public void insertDocumentDynamicallyTrimLastParagraph() throws Exception
    {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()] -inline>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(mDocument)).build();

        buildReport(template, doc, "src", new Class[]{DocumentTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");

        template = new Document(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");
        Assert.assertEquals(1, template.getFirstSection().getBody().getParagraphs().getCount());
    }

    @Test
    public void insertDocumentDynamically() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()]>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(mDocument)).build();

        buildReport(template, doc, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");
    }

    @Test
    public void sourseListNumbering() throws Exception
    {
        //ExStart:SourseListNumbering
        //GistId:f99d87e10ab87a581c52206321d8b617
        //ExFor:ReportingEngine.BuildReport(Document, Object[], String[])
        //ExSummary:Shows how to keep inserted numbering as is.
        // By default, numbered lists from a template document are continued when their identifiers match those from a document being inserted.
        // With "-sourceNumbering" numbering should be separated and kept as is.
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()]>>" + System.lineSeparator() + "<<doc [src.getDocument()] -sourceNumbering>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(getMyDir() + "List item.docx")).build();

        ReportingEngine engine = new ReportingEngine(); { engine.setOptions(ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS); }
        engine.buildReport(template, new Object[] { doc }, new String[] { "src" });

        template.save(getArtifactsDir() + "ReportingEngine.SourseListNumbering.docx");
        //ExEnd:SourseListNumbering

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SourseListNumbering.docx", getGoldsDir() + "ReportingEngine.SourseListNumbering Gold.docx"));
    }

    @Test
    public void insertDocumentDynamicallyByStream() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocumentStream()]>>");

        DocumentTestClass docStream = new DocumentTestBuilder()
                .withDocumentStream(new FileInputStream(mDocument)).build();

        buildReport(template, docStream, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");
    }

    @Test
    public void insertDocumentDynamicallyByBytes() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocumentBytes()]>>");

        DocumentTestClass docBytes = new DocumentTestBuilder()
                .withDocumentBytes(Files.readAllBytes(Paths.get(mDocument))).build();

        buildReport(template, docBytes, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");
    }

    @Test
    public void insertDocumentDynamicallyByUri() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocumentString()]>>");

        DocumentTestClass docUri = new DocumentTestBuilder()
                .withDocumentString("http://www.snee.com/xml/xslt/sample.doc").build();

        buildReport(template, docUri, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");
    }

    @Test
    public void insertImageDynamically() throws Exception {
        Document template =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImage()]>>", ShapeType.TEXT_BOX);

        ImageTestClass image = new ImageTestBuilder().withImage(mImage).build();
        buildReport(template, image, "src");

        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");
    }

    @Test
    public void insertImageDynamicallyByStream() throws Exception {
        Document template =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageStream()]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageStream = new ImageTestBuilder()
                .withImageStream(new FileInputStream(mImage)).build();

        buildReport(template, imageStream, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");
    }

    @Test
    public void insertImageDynamicallyByBytes() throws Exception {
        Document template =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageBytes()]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageBytes = new ImageTestBuilder().withImageBytes(Files.readAllBytes(Paths.get(mImage))).build();

        buildReport(template, imageBytes, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");
    }

    @Test
    public void insertImageDynamicallyByUri() throws Exception {
        Document template =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageString()]>>", ShapeType.TEXT_BOX);
        ImageTestClass imageUri = new ImageTestBuilder().withImageString("https://metrics.aspose.com/img/headergraphics.svg").build();

        buildReport(template, imageUri, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");
    }

    @Test
    public void insertHyperlinksDynamically() throws Exception {
        Document template = new Document(getMyDir() + "Reporting engine template - Inserting hyperlinks (Java).docx");
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

        template.save(getArtifactsDir() + "ReportingEngine.InsertHyperlinksDynamically.docx");
    }

    @Test (dataProvider = "insertHtmlDinamicallyDataProvider")
    public void insertHtmlDinamically(String templateText) throws Exception
    {
        String html = FileUtils.readFileToString(new File(getMyDir() + "Reporting engine template - Html (Java).html"), "utf-8");

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln(templateText);

        buildReport(doc, html, "html_text");
        doc.save(getArtifactsDir() + "ReportingEngine.InsertHtmlDinamically.docx");
    }

    @DataProvider(name = "insertHtmlDinamicallyDataProvider")
    public static Object[][] insertHtmlDinamicallyDataProvider() {
        return new Object[][]
                {
                        {"<<[html_text] -html>>"},
                        {"<<html [html_text]>>"},
                        {"<<html [html_text] -sourceStyles>>"},
                };
    }

    @Test(expectedExceptions = IllegalStateException.class)
    public void withoutKnownType() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("<<[new Date()]:”dd.MM.yyyy”>>");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, "");
    }

    @Test
    public void workWithKnownTypes() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("<<[new GregorianCalendar(2016, 0, 20).getTime()]:”dd.MM.yyyy”>>");
        builder.writeln("<<[new GregorianCalendar(2016, 0, 20).getTime()]:”dd”>>");
        builder.writeln("<<[new GregorianCalendar(2016, 0, 20).getTime()]:”MM”>>");
        builder.writeln("<<[new GregorianCalendar(2016, 0, 20).getTime()]:”yyyy”>>");
        builder.writeln("<<[new GregorianCalendar(2016, 1, 20).get(Calendar.MONTH)]>>");

        buildReport(doc, "", new Class[]{GregorianCalendar.class, Calendar.class});

        doc.save(getArtifactsDir() + "ReportingEngine.KnownTypes.docx");
    }

    @Test
    public void workWithSingleColumnTableRow() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row (Java).docx");
        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRow.docx");
    }

    @Test
    public void workWithSingleColumnTableRowGreedy() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row greedy (Java).docx");
        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRowGreedy.docx");
    }

    @Test
    public void tableRowConditionalBlocks() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Table row conditional blocks (Java).docx");

        ArrayList<ClientTestClass> clients = new ArrayList<>();
        clients.add(new ClientTestClass("John Monrou", "France", "27 RUE PASTEUR"));
        clients.add(new ClientTestClass("James White", "England", "14 Tottenham Court Road"));
        clients.add(new ClientTestClass("Kate Otts", "New Zealand", "Wellington 6004"));

        buildReport(doc, clients, "clients", new Class[]{ClientTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TableRowConditionalBlocks.docx");
    }

    @Test
    public void ifGreedy() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - If greedy (Java).docx");

        AsposeData obj = new AsposeData();
        obj.setList(new ArrayList<>());
        obj.getList().add("abc");

        buildReport(doc, obj);

        doc.save(getArtifactsDir() + "ReportingEngine.IfGreedy.docx");
    }

    public static class AsposeData {
        public ArrayList<String> getList() {
            return mList;
        }

        public void setList(final ArrayList<String> value) {
            mList = value;
        }

        private ArrayList<String> mList;
    }

    @Test
    public void stretchImagefitHeight() throws Exception {
        Document doc =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageStream()] -fitHeight>>",
                        ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
                .withImageStream(new FileInputStream(mImage)).build();
        buildReport(doc, imageStream, "src");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        ByteArrayInputStream byteStream = new ByteArrayInputStream(dstStream.toByteArray());

        doc = new Document(byteStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Object shapeNode : shapes) {
            Shape shape = (Shape) shapeNode;
            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that width is keeped and height is changed
            Assert.assertNotEquals(shape.getHeight(), 346.35);
            Assert.assertEquals(shape.getWidth(), 431.5);
        }
    }

    @Test
    public void stretchImagefitWidth() throws Exception {
        Document doc =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageStream()] -fitWidth>>",
                        ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
                .withImageStream(new FileInputStream(mImage)).build();
        buildReport(doc, imageStream, "src");

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        doc.save(baos, SaveFormat.DOCX);

        ByteArrayInputStream docStream = new ByteArrayInputStream(baos.toByteArray());

        doc = new Document(docStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Object shapeNode : shapes) {
            Shape shape = (Shape) shapeNode;

            // Assert that the image is really insert in textbox and 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that height is keeped and width is changed
            Assert.assertNotEquals(shape.getWidth(), 346.35);
            Assert.assertEquals(shape.getHeight(), 431.5);
        }
    }

    @Test
    public void stretchImagefitSize() throws Exception {
        Document doc =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageStream()] -fitSize>>",
                        ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
                .withImageStream(new FileInputStream(mImage)).build();
        buildReport(doc, imageStream, "src");

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        doc.save(baos, SaveFormat.DOCX);

        ByteArrayInputStream docStream = new ByteArrayInputStream(baos.toByteArray());

        doc = new Document(docStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Object shapeNode : shapes) {
            Shape shape = (Shape) shapeNode;

            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that height is changed and width is changed
            Assert.assertNotEquals(346.35, shape.getHeight());
            Assert.assertNotEquals(431.5, shape.getWidth());
        }
    }

    @Test
    public void stretchImagefitSizeLim() throws Exception {
        Document doc =
                DocumentHelper.createTemplateDocumentWithDrawObjects("<<image [src.getImageStream()] -fitSizeLim>>",
                        ShapeType.TEXT_BOX);

        ImageTestClass imageStream = new ImageTestBuilder()
                .withImageStream(new FileInputStream(mImage)).build();
        buildReport(doc, imageStream, "src");

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        doc.save(baos, SaveFormat.DOCX);

        ByteArrayInputStream docStream = new ByteArrayInputStream(baos.toByteArray());

        doc = new Document(docStream);
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);

        for (Object shapeNode : shapes) {
            Shape shape = (Shape) shapeNode;

            // Assert that the image is really insert in textbox 
            Assert.assertNotNull(shape.getFill().getImageBytes());

            // Assert that textbox size are equal image size
            Assert.assertEquals(shape.getHeight(), 300.0);
            Assert.assertEquals(shape.getWidth(), 300.0);
        }
    }

    @Test(expectedExceptions = IllegalStateException.class)
    public void withoutMissingMembers() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        // Add templete to the document for reporting engine
        DocumentHelper.insertBuilderText(builder,
                new String[]{"<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>"});

        // Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
        buildReport(builder.getDocument(), new DataSet(), "");
    }

    @Test
    public void missingMembers() throws Exception {
        //ExStart:MissingMembers
        //GistId:a76df4b18bee76d169e55cdf6af8129c
        //ExFor:ReportingEngine.BuildReport(Document, Object, String)
        //ExFor:ReportingEngine.MissingMemberMessage
        //ExFor:ReportingEngine.Options
        //ExSummary:Shows how to allow missinng members.
        DocumentBuilder builder = new DocumentBuilder();
        builder.writeln("<<[missingObject.First().id]>>");
        builder.writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        ReportingEngine engine = new ReportingEngine(); { engine.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS); }
        engine.setMissingMemberMessage("Missed");
        engine.buildReport(builder.getDocument(), new DataSet(), "");
        //ExEnd:MissingMembers

        // Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
        Assert.assertEquals("Missed", builder.getDocument().getText().trim());
    }

    @Test(dataProvider = "inlineErrorMessagesDataProvider")
    public void inlineErrorMessages(String templateText, String result) throws Exception {
        DocumentBuilder builder = new DocumentBuilder();
        DocumentHelper.insertBuilderText(builder, new String[]{templateText});

        buildReport(builder.getDocument(), new DataSet(), "", ReportBuildOptions.INLINE_ERROR_MESSAGES);

        Assert.assertEquals(builder.getDocument().getFirstSection().getBody().getParagraphs().get(0).getText().replaceAll("\\s+$", ""), result);
    }

    @DataProvider(name = "inlineErrorMessagesDataProvider")
    public static Object[][] inlineErrorMessagesDataProvider() {
        return new Object[][]
                {
                        {"<<[missingObject.First().id]>>", "<<[missingObject.First( Error! Can not get the value of member 'missingObject' on type 'class com.aspose.words.net.System.Data.DataSet'. ).id]>>"},
                        {"<<[new DateTime()]:\"dd.MM.yyyy\">>", "<<[new DateTime( Error! A type identifier is expected. )]:\"dd.MM.yyyy\">>"},
                        {"<<]>>", "<<] Error! Character ']' is unexpected. >>"},
                        {"<<[>>", "<<[>> Error! An expression is expected."},
                        {"<<>>", "<<>> Error! Tag end is unexpected."},
                };
    }

    @Test
    public void setBackgroundColorDynamically() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Background color (Java).docx");

        ArrayList<ColorItemTestClass> colors = new ArrayList<>();
        colors.add(new ColorItemTestBuilder().withColor("Black", Color.BLACK).build());
        colors.add(new ColorItemTestBuilder().withColor("Red", new Color((255), (0), (0))).build());
        colors.add(new ColorItemTestBuilder().withColor("Empty", null).build());

        buildReport(doc, colors, "Colors", new Class[]{ColorItemTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SetBackgroundColorDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SetBackgroundColorDynamically.docx",
                getGoldsDir() + "ReportingEngine.SetBackgroundColorDynamically Gold.docx"));
    }

    @Test
    public void setTextColorDynamically() throws Exception
    {
        Document doc = new Document(getMyDir() + "Reporting engine template - Text color (Java).docx");

        ArrayList<ColorItemTestClass> colors = new ArrayList<>();
        {
            colors.add(new ColorItemTestBuilder().withColor("Black", Color.BLUE).build());
            colors.add(new ColorItemTestBuilder().withColor("Red", new Color((255), (0), (0))).build());
            colors.add(new ColorItemTestBuilder().withColor("Empty", null).build());
        }

        buildReport(doc, colors, "Colors", new Class[]{ColorItemTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SetTextColorDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SetTextColorDynamically.docx",
                getGoldsDir() + "ReportingEngine.SetTextColorDynamically Gold.docx"));
    }

    @Test
    public void doNotRemoveEmptyParagraphs() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Remove empty paragraphs (Java).docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx");
    }

    @Test
    public void removeEmptyParagraphs() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Remove empty paragraphs (Java).docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        doc.save(getArtifactsDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");
    }

    @Test(dataProvider = "mergingTableCellsDynamicallyDataProvider")
    public void mergingTableCellsDynamically(final String value1, final String value2, final String resultDocumentName) throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Merging table cells dynamically (Java).docx");

        ArrayList<ClientTestClass> clients = new ArrayList<>();
        clients.add(new ClientTestClass("John Monrou", "France", "27 RUE PASTEUR"));
        clients.add(new ClientTestClass("James White", "New Zealand", "14 Tottenham Court Road"));
        clients.add(new ClientTestClass("Kate Otts", "New Zealand", "Wellington 6004"));

        buildReport(doc, new Object[]{value1, value2, clients}, new String[]{"value1", "value2", "clients"}, new Class[]{ClientTestClass.class});

        doc.save(getArtifactsDir() + resultDocumentName + FileFormatUtil.saveFormatToExtension(SaveFormat.DOCX));
    }

    @DataProvider(name = "mergingTableCellsDynamicallyDataProvider")
    public static Object[][] mergingTableCellsDynamicallyDataProvider() {
        return new Object[][]
                {
                        {"Hello", "Hello", "ReportingEngine.MergingTableCellsDynamically.Merged"},
                        {"Hello", "Name", "ReportingEngine.MergingTableCellsDynamically.NotMerged"}
                };
    }

    @Test
    public void xmlDataStringWithoutSchema() throws Exception
    {
        //ExStart
        //ExFor:XmlDataSource
        //ExFor:XmlDataSource.#ctor(String)
        //ExSummary:Show how to use XML as a data source (string).
        Document doc = new Document(getMyDir() + "Reporting engine template - XML data destination (Java).docx");

        XmlDataSource dataSource = new XmlDataSource(getMyDir() + "List of people.xml");
        buildReport(doc, dataSource, "persons");

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataString.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataString.docx",
                getGoldsDir() + "ReportingEngine.DataSource Gold.docx"));
    }

    @Test
    public void xmlDataStreamWithoutSchema() throws Exception
    {
        //ExStart
        //ExFor:XmlDataSource
        //ExFor:XmlDataSource.#ctor(Stream)
        //ExSummary:Show how to use XML as a data source (stream).
        Document doc = new Document(getMyDir() + "Reporting engine template - XML data destination (Java).docx");

        InputStream stream = new FileInputStream(getMyDir() + "List of people.xml");
        try {
            XmlDataSource dataSource = new XmlDataSource(stream);
            buildReport(doc, dataSource, "persons");
        } finally {
            stream.close();
        }

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataStream.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataStream.docx",
                getGoldsDir() + "ReportingEngine.DataSource Gold.docx"));
    }

    @Test
    public void xmlDataWithNestedElements() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Data destination with nested elements (Java).docx");

        XmlDataSource dataSource = new XmlDataSource(getMyDir() + "Nested elements.xml");
        buildReport(doc, dataSource, "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataWithNestedElements.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataWithNestedElements.docx",
                getGoldsDir() + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
    }

    @Test
    public void jsonDataString() throws Exception
    {
        //ExStart
        //ExFor:JsonDataLoadOptions
        //ExFor:JsonDataLoadOptions.#ctor
        //ExFor:JsonDataLoadOptions.ExactDateTimeParseFormats
        //ExFor:JsonDataLoadOptions.AlwaysGenerateRootObject
        //ExFor:JsonDataLoadOptions.PreserveSpaces
        //ExFor:JsonDataLoadOptions.SimpleValueParseMode
        //ExFor:JsonDataSource
        //ExFor:JsonDataSource.#ctor(String,JsonDataLoadOptions)
        //ExFor:JsonSimpleValueParseMode
        //ExSummary:Shows how to use JSON as a data source (string).
        Document doc = new Document(getMyDir() + "Reporting engine template - JSON data destination (Java).docx");

        JsonDataLoadOptions options = new JsonDataLoadOptions();
        {
            options.setExactDateTimeParseFormats(Arrays.asList(new String[]{"MM/dd/yyyy", "MM.d.yy", "MM d yy"}));
        }

        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "List of people.json", options);
        buildReport(doc, dataSource, "persons");

        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataString.docx");
        //ExEnd
    }

    @Test
    public void jsonDataStream() throws Exception
    {
        //ExStart
        //ExFor:JsonDataSource.#ctor(Stream,JsonDataLoadOptions)
        //ExSummary:Shows how to use JSON as a data source (stream).
        Document doc = new Document(getMyDir() + "Reporting engine template - JSON data destination (Java).docx");

        JsonDataLoadOptions options = new JsonDataLoadOptions();
        {
            options.setExactDateTimeParseFormats(Arrays.asList(new String[]{"MM/dd/yyyy", "MM.d.yy", "MM d yy"}));
        }

        InputStream stream = new FileInputStream(getMyDir() + "List of people.json");
        try {
            JsonDataSource dataSource = new JsonDataSource(stream, options);
            buildReport(doc, dataSource, "persons");
        } finally {
            stream.close();
        }

        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataStream.docx");
        //ExEnd
    }

    @Test
    public void jsonDataWithNestedElements() throws Exception {
        Document doc = new Document(getMyDir() + "Reporting engine template - Data destination with nested elements (Java).docx");

        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "Nested elements.json");
        buildReport(doc, dataSource, "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataWithNestedElements.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.JsonDataWithNestedElements.docx",
                getGoldsDir() + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
    }

    @Test
    public void jsonDataPreserveSpaces() throws Exception
    {
        final String TEMPLATE = "LINE BEFORE\r<<[LineWhitespace]>>\r<<[BlockWhitespace]>>LINE AFTER";
        final String EXPECTED_RESULT = "LINE BEFORE\r    \r\r\r\r\rLINE AFTER";
        final String JSON =
                "{" +
                        "    \"LineWhitespace\" : \"    \"," +
                        "    \"BlockWhitespace\" : \"\r\n\r\n\r\n\r\n\"" +
                        "}";

        ByteArrayInputStream stream = new ByteArrayInputStream(JSON.getBytes(StandardCharsets.UTF_8));
        try
        {
            JsonDataLoadOptions options = new JsonDataLoadOptions();
            options.setPreserveSpaces(true);
            options.setSimpleValueParseMode(JsonSimpleValueParseMode.STRICT);

            JsonDataSource dataSource = new JsonDataSource(stream, options);

            DocumentBuilder builder = new DocumentBuilder();
            builder.write(TEMPLATE);

            buildReport(builder.getDocument(), dataSource, "ds");

            Assert.assertEquals(EXPECTED_RESULT + ControlChar.SECTION_BREAK, builder.getDocument().getText());
        }
        finally { if (stream != null) stream.close(); }
    }

    @Test
    public void csvDataString() throws Exception
    {
        //ExStart
        //ExFor:CsvDataLoadOptions
        //ExFor:CsvDataLoadOptions.#ctor
        //ExFor:CsvDataLoadOptions.#ctor(Boolean)
        //ExFor:CsvDataLoadOptions.Delimiter
        //ExFor:CsvDataLoadOptions.CommentChar
        //ExFor:CsvDataLoadOptions.HasHeaders
        //ExFor:CsvDataLoadOptions.QuoteChar
        //ExFor:CsvDataSource
        //ExFor:CsvDataSource.#ctor(String,CsvDataLoadOptions)
        //ExSummary:Shows how to use CSV as a data source (string).
        Document doc = new Document(getMyDir() + "Reporting engine template - CSV data destination (Java).docx");

        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.setDelimiter(';');
        loadOptions.setCommentChar('$');
        loadOptions.hasHeaders(true);
        loadOptions.setQuoteChar('"');

        CsvDataSource dataSource = new CsvDataSource(getMyDir() + "List of people.csv", loadOptions);
        buildReport(doc, dataSource, "persons");

        doc.save(getArtifactsDir() + "ReportingEngine.CsvDataString.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.CsvDataString.docx",
                getGoldsDir() + "ReportingEngine.CsvData Gold.docx"));
    }

    @Test
    public void csvDataStream() throws Exception
    {
        //ExStart
        //ExFor:CsvDataSource.#ctor(Stream,CsvDataLoadOptions)
        //ExSummary:Shows how to use CSV as a data source (stream).
        Document doc = new Document(getMyDir() + "Reporting engine template - CSV data destination (Java).docx");

        CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
        loadOptions.setDelimiter(';');
        loadOptions.setCommentChar('$');

        InputStream stream = new FileInputStream(getMyDir() + "List of people.csv");
        try {
            CsvDataSource dataSource = new CsvDataSource(stream, loadOptions);
            buildReport(doc, dataSource, "persons");
        } finally {
            stream.close();
        }

        doc.save(getArtifactsDir() + "ReportingEngine.CsvDataStream.docx");
        //ExEnd

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.CsvDataStream.docx",
                getGoldsDir() + "ReportingEngine.CsvData Gold.docx"));
    }

    @Test (dataProvider = "insertComboboxDropdownListItemsDynamicallyDataProvider")
    public void insertComboboxDropdownListItemsDynamically(int sdtType) throws Exception
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

        doc.save(getArtifactsDir() + MessageFormat.format("ReportingEngine.InsertComboboxDropdownListItemsDynamically_{0}.docx", sdtType));

        doc = new Document(getArtifactsDir() +
                           MessageFormat.format("ReportingEngine.InsertComboboxDropdownListItemsDynamically_{0}.docx", sdtType));

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

	@DataProvider(name = "insertComboboxDropdownListItemsDynamicallyDataProvider")
	public static Object[][] insertComboboxDropdownListItemsDynamicallyDataProvider() {
		return new Object[][]
		{
			{SdtType.COMBO_BOX},
			{SdtType.DROP_DOWN_LIST},
		};
	}

    @Test
    public void updateFieldsSyntaxAware() throws Exception
    {
        //ExStart:UpdateFieldsSyntaxAware
        //ExFor:ReportingEngine.Options
        //ExSummary:Shows how to set options for Reporting Engine
        //GistId:66dd22f0854357e394a013b536e2181b
        Document doc = new Document(getMyDir() + "Reporting engine template - Fields (Java).docx");

        // Note that enabling of the option makes the engine to update fields while building a report,
        // so there is no need to update fields separately after that.
        ReportingEngine engine = new ReportingEngine();
        engine.setOptions(ReportBuildOptions.UPDATE_FIELDS_SYNTAX_AWARE);
        engine.buildReport(doc, new String[] { "First topic", "Second topic", "Third topic" }, "topics");

        doc.save(getArtifactsDir() + "ReportingEngine.UpdateFieldsSyntaxAware.docx");
        //ExEnd:UpdateFieldsSyntaxAware
    }

    @Test
    public void dollarTextFormat() throws Exception
    {
        //ExStart:DollarTextFormat
        //GistId:f0964b777330b758f6b82330b040b24c
        //ExFor:ReportingEngine.BuildReport(Document, Object, String)
        //ExSummary:Shows how to display values as dollar text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("<<[ds.getValue1()]:dollarText>>\r<<[ds.getValue2()]:dollarText>>");

        NumericTestClass testData = new NumericTestBuilder().withValues(1234, 5621718.589).build();

        ReportingEngine report = new ReportingEngine();
        report.getKnownTypes().add(NumericTestClass.class);
        report.buildReport(doc, testData, "ds");

        doc.save(getArtifactsDir() + "ReportingEngine.DollarTextFormat.docx");
        //ExEnd:DollarTextFormat

        Assert.assertEquals("one thousand two hundred thirty-four and 00/100\rfive million six hundred twenty-one thousand seven hundred eighteen and 59/100\r\f", doc.getText());
    }

    @Test (enabled = false, description = "To avoid exception with 'SetRestrictedTypes' after execution other tests.")
    public void restrictedTypes() throws Exception
    {
        //ExStart:RestrictedTypes
        //GistId:b9e728d2381f759edd5b31d64c1c4d3f
        //ExFor:ReportingEngine.SetRestrictedTypes(Type[])
        //ExSummary:Shows how to deny access to members of types considered insecure.
        Document doc =
                DocumentHelper.createSimpleDocument(
                        "<<var [typeVar = \"\".getClass().getName()]>><<[typeVar]>>");

        // Note, that you can't set restricted types during or after building a report.
        ReportingEngine.setRestrictedTypes(Class.class);
        // We set "AllowMissingMembers" option to avoid exceptions during building a report.
        ReportingEngine engine = new ReportingEngine();
        engine.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
        engine.buildReport(doc, new Object());

        // We get an empty string because we can't access the GetType() method.
        Assert.assertEquals(doc.getText().trim(), "");
        //ExEnd:RestrictedTypes
    }

    @Test
    public void word2016Charts() throws Exception
    {
        //ExStart:Word2016Charts
        //GistId:9c17d666c47318436785490829a3984f
        //ExFor:ReportingEngine.BuildReport(Document, Object[], String[])
        //ExSummary:Shows how to work with charts from word 2016.
        Document doc = new Document(getMyDir() + "Reporting engine template - Word 2016 Charts (Java).docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, new Object[] { Common.getShares(), Common.getShareQuotes() },
                new String[] { "shares", "quotes" });

        doc.save(getArtifactsDir() + "ReportingEngine.Word2016Charts.docx");
        //ExEnd:Word2016Charts
    }

    @Test
    public void removeParagraphsSelectively() throws Exception
    {
        //ExStart:RemoveParagraphsSelectively
        //GistId:a76df4b18bee76d169e55cdf6af8129c
        //ExFor:ReportingEngine.BuildReport(Document, Object, String)
        //ExSummary:Shows how to remove paragraphs selectively.
        // Template contains tags with an exclamation mark. For such tags, empty paragraphs will be removed.
        Document doc = new Document(getMyDir() + "Reporting engine template - Selective remove paragraphs.docx");

        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(doc, false, "value");

        doc.save(getArtifactsDir() + "ReportingEngine.SelectiveDeletionOfParagraphs.docx");
        //ExEnd:RemoveParagraphsSelectively

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.SelectiveDeletionOfParagraphs.docx", getGoldsDir() + "ReportingEngine.SelectiveDeletionOfParagraphs Gold.docx"));
    }

    private static void buildReport(final Document document, final Object dataSource) throws Exception {
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(document, dataSource);
    }

    private static void buildReport(final Document document, final Object dataSource, final String dataSourceName) throws Exception {
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(final Document document, final Object dataSource, final String dataSourceName,
                                    final int reportBuildOptions) throws Exception {
        ReportingEngine engine = new ReportingEngine();
        engine.setOptions(reportBuildOptions);

        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(final Document document, final Object dataSource, final Class[] knownTypes) throws Exception {
        ReportingEngine engine = new ReportingEngine();

        for (Class knownType : knownTypes) {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource);
    }

    private static void buildReport(final Document document, final Object[] dataSource, final String[] dataSourceName) throws Exception {
        ReportingEngine engine = new ReportingEngine();
        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(final Document document, final Object[] dataSource, final String[] dataSourceName,
                                    final Class[] knownTypes) throws Exception {
        ReportingEngine engine = new ReportingEngine();

        for (Class knownType : knownTypes) {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(final Document document, final Object dataSource, final String dataSourceName,
                                    final Class[] knownTypes, final int reportBuildOptions) throws Exception {
        ReportingEngine engine = new ReportingEngine();
        engine.setOptions(reportBuildOptions);

        for (Class knownType : knownTypes) {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(final Document document, final Object[] dataSource, final String[] dataSourceName,
                                    final Class[] knownTypes, final int reportBuildOptions) throws Exception {
        ReportingEngine engine = new ReportingEngine();
        engine.setOptions(reportBuildOptions);

        for (Class knownType : knownTypes) {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource, dataSourceName);
    }

    private static void buildReport(final Document document, final Object dataSource, final String dataSourceName,
                                    final Class[] knownTypes) throws Exception {
        ReportingEngine engine = new ReportingEngine();

        for (Class knownType : knownTypes) {
            engine.getKnownTypes().add(knownType);
        }

        engine.buildReport(document, dataSource, dataSourceName);
    }
}