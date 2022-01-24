package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.time.LocalDate;
import java.util.*;
import java.util.List;

@Test
public class ExReportingEngine extends ApiExampleBase {
    private final String mImage = getImageDir() + "Logo.jpg";
    private final String mDocument = getMyDir() + "ReportingEngine.TestDataTable.Java.docx";

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
        Document doc = new Document(getMyDir() + "ReportingEngine.TestDataTable.Java.docx");

        buildReport(doc, Common.getContracts(), "Contracts", new Class[]{ContractTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestDataTable.docx");
    }

    @Test
    public void progressiveTotal() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.Total.Java.docx");

        buildReport(doc, Common.getContracts(), "Contracts", new Class[]{ContractTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.Total.docx");
    }

    @Test
    public void nestedDataTableTest() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestNestedDataTable.Java.docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class, ContractTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestNestedDataTable.docx");
    }

    @Test
    public void restartingListNumberingDynamically() throws Exception {
        Document template = new Document(getMyDir() + "ReportingEngine.RestartingListNumberingDynamically.Java.docx");

        buildReport(template, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class, ContractTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        template.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamically Gold.docx"));
    }

    @Test
    public void restartingListNumberingDynamicallyWhileInsertingDocumentDynamically() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()] -build>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(getMyDir() + "ReportingEngine.RestartingListNumberingDynamically.Java.docx")).build();

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
                .withDocument(new Document(getMyDir() + "ReportingEngine.RestartingListNumberingDynamically.Java.docx")).build();

        buildReport(mainTemplate, new Object[]{template1, template2, doc, Common.getManagers()}, new String[]{"src", "src1", "src2", "Managers"}, new Class[]{ManagerTestClass.class, ContractTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        mainTemplate.save(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx", getGoldsDir() + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
    }

    @Test
    public void chartTest() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestChart.Java.docx");

        buildReport(doc, Common.getManagers(), "managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestChart.docx");
    }

    @Test
    public void bubbleChartTest() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestBubbleChart.Java.docx");

        buildReport(doc, Common.getManagers(), "managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestBubbleChart.docx");
    }

    @Test
    public void setChartSeriesColorsDynamically() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.SetChartSeriesColorDynamically.Java.docx");

        buildReport(doc, Common.getManagers(), "managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SetChartSeriesColorDynamically.docx");
    }

    @Test
    public void setPointColorsDynamically() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.SetPointColorDynamically.Java.docx");

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
        Document doc = new Document(getMyDir() + "ReportingEngine.TestRemoveChartSeries.Java.docx");

        int condition = 2;
        buildReport(doc, new Object[]{Common.getManagers(), condition}, new String[]{"managers", "condition"}, new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TestRemoveChartSeries.docx");
    }

    @Test
    public void indexOf() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.TestIndexOf.Java.docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals("The names are: John Smith, Tony Anderson, July James\f", doc.getText());
    }

    @Test
    public void ifElse() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.IfElse.Java.docx");

        buildReport(doc, Common.getManagers(), "m", new Class[]{ManagerTestClass.class});

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals("You have chosen 3 item(s).\f", doc.getText());
    }

    @Test
    public void ifElseWithoutData() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.IfElse.Java.docx");

        buildReport(doc, Common.getEmptyManagers(), "m", new Class[]{ManagerTestClass.class});

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        Assert.assertEquals("You have chosen no items.\f", doc.getText());
    }

    @Test
    public void extensionMethods() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.ExtensionMethods.Java.docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});
        doc.save(getArtifactsDir() + "ReportingEngine.ExtensionMethods.docx");
    }

    @Test
    public void operators() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.Operators.Java.docx");

        NumericTestClass testData = new NumericTestBuilder().withValuesAndLogical(1, 2.0, 3, null, true).build();

        buildReport(doc, testData, "ds", new Class[]{NumericTestBuilder.class});
        doc.save(getArtifactsDir() + "ReportingEngine.Operators.docx");
    }

    @Test
    public void contextualObjectMemberAccess() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.ContextualObjectMemberAccess.Java.docx");

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
    public void insertDocumentDynamically() throws Exception {
        Document template = DocumentHelper.createSimpleDocument("<<doc [src.getDocument()]>>");

        DocumentTestClass doc = new DocumentTestBuilder()
                .withDocument(new Document(mDocument)).build();

        buildReport(template, doc, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertDocumentDynamically.docx");
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

        ImageTestClass image = new ImageTestBuilder().withImage(ImageIO.read(new File(mImage))).build();
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
        ImageTestClass imageUri = new ImageTestBuilder()
                .withImageString(
                        "http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png")
                .build();

        buildReport(template, imageUri, "src");
        template.save(getArtifactsDir() + "ReportingEngine.InsertImageDynamically.docx");
    }

    @Test
    public void insertHyperlinksDynamically() throws Exception {
        Document template = new Document(getMyDir() + "ReportingEngine.InsertingHyperlinks.Java.docx");
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
        Document doc = new Document(getMyDir() + "ReportingEngine.SingleColumnTableRow.Java.docx");
        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRow.docx");
    }

    @Test
    public void workWithSingleColumnTableRowGreedy() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.SingleColumnTableRowGreedy.Java.docx");
        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.SingleColumnTableRowGreedy.docx");
    }

    @Test
    public void tableRowConditionalBlocks() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.TableRowConditionalBlocks.Java.docx");

        ArrayList<ClientTestClass> clients = new ArrayList<>();
        clients.add(new ClientTestClass("John Monrou", "France", "27 RUE PASTEUR"));
        clients.add(new ClientTestClass("James White", "England", "14 Tottenham Court Road"));
        clients.add(new ClientTestClass("Kate Otts", "New Zealand", "Wellington 6004"));

        buildReport(doc, clients, "clients", new Class[]{ClientTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.TableRowConditionalBlocks.docx");
    }

    @Test
    public void ifGreedy() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.IfGreedy.Java.docx");

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
    public void withMissingMembers() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        // Add templete to the document for reporting engine
        DocumentHelper.insertBuilderText(builder,
                new String[]{"<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>"});

        buildReport(builder.getDocument(), new DataSet(), "", ReportBuildOptions.ALLOW_MISSING_MEMBERS);

        // Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
        Assert.assertEquals(ControlChar.PARAGRAPH_BREAK + ControlChar.PARAGRAPH_BREAK + ControlChar.SECTION_BREAK,
                builder.getDocument().getText());
    }

    @Test(dataProvider = "inlineErrorMessagesDataProvider")
    public void inlineErrorMessages(String templateText, String result) throws Exception {
        DocumentBuilder builder = new DocumentBuilder();
        DocumentHelper.insertBuilderText(builder, new String[]{templateText});

        buildReport(builder.getDocument(), new DataSet(), "", ReportBuildOptions.INLINE_ERROR_MESSAGES);

        Assert.assertEquals(builder.getDocument().getFirstSection().getBody().getParagraphs().get(0).getText().replaceAll("\\s+$", ""), result);
    }

    //JAVA-added data provider for test method
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
    public void setBackgroundColor() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.BackColor.Java.docx");

        ArrayList<ColorItemTestClass> colors = new ArrayList<>();
        colors.add(new ColorItemTestBuilder().withColor("Black", Color.BLACK).build());
        colors.add(new ColorItemTestBuilder().withColor("Red", new Color((255), (0), (0))).build());
        colors.add(new ColorItemTestBuilder().withColor("Empty", null).build());

        buildReport(doc, colors, "Colors", new Class[]{ColorItemTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.BackColor.docx");
    }

    @Test
    public void doNotRemoveEmptyParagraphs() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.RemoveEmptyParagraphs.Java.docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class});

        doc.save(getArtifactsDir() + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx");
    }

    @Test
    public void removeEmptyParagraphs() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.RemoveEmptyParagraphs.Java.docx");

        buildReport(doc, Common.getManagers(), "Managers", new Class[]{ManagerTestClass.class}, ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);

        doc.save(getArtifactsDir() + "ReportingEngine.RemoveEmptyParagraphs.docx");
    }

    @Test(dataProvider = "mergingTableCellsDynamicallyDataProvider")
    public void mergingTableCellsDynamically(final String value1, final String value2, final String resultDocumentName) throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.MergingTableCellsDynamically.Java.docx");

        ArrayList<ClientTestClass> clients = new ArrayList<>();
        clients.add(new ClientTestClass("John Monrou", "France", "27 RUE PASTEUR"));
        clients.add(new ClientTestClass("James White", "New Zealand", "14 Tottenham Court Road"));
        clients.add(new ClientTestClass("Kate Otts", "New Zealand", "Wellington 6004"));

        buildReport(doc, new Object[]{value1, value2, clients}, new String[]{"value1", "value2", "clients"}, new Class[]{ClientTestClass.class});

        doc.save(getArtifactsDir() + resultDocumentName + FileFormatUtil.saveFormatToExtension(SaveFormat.DOCX));
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "mergingTableCellsDynamicallyDataProvider")
    public static Object[][] mergingTableCellsDynamicallyDataProvider() {
        return new Object[][]
                {
                        {"Hello", "Hello", "ReportingEngine.MergingTableCellsDynamically.Merged"},
                        {"Hello", "Name", "ReportingEngine.MergingTableCellsDynamically.NotMerged"}
                };
    }

    @Test
    public void xmlDataStringWithoutSchema() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.DataSource.Java.docx");

        XmlDataSource dataSource = new XmlDataSource(getMyDir() + "List of people.xml");
        buildReport(doc, dataSource, "persons");

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataString.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataString.docx",
                getGoldsDir() + "ReportingEngine.DataSource Gold.docx"));
    }

    @Test
    public void xmlDataStreamWithoutSchema() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.DataSource.Java.docx");

        InputStream stream = new FileInputStream(getMyDir() + "List of people.xml");
        try {
            XmlDataSource dataSource = new XmlDataSource(stream);
            buildReport(doc, dataSource, "persons");
        } finally {
            stream.close();
        }

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataStream.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataStream.docx",
                getGoldsDir() + "ReportingEngine.DataSource Gold.docx"));
    }

    @Test
    public void xmlDataWithNestedElements() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.DataSourceWithNestedElements.Java.docx");

        XmlDataSource dataSource = new XmlDataSource(getMyDir() + "Nested elements.xml");
        buildReport(doc, dataSource, "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.XmlDataWithNestedElements.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.XmlDataWithNestedElements.docx",
                getGoldsDir() + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
    }

    @Test
    public void jsonDataString() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.DataSource.Java.docx");

        JsonDataLoadOptions options = new JsonDataLoadOptions();
        {
            options.setExactDateTimeParseFormats(Arrays.asList(new String[]{"MM/dd/yyyy", "MM.d.yy", "MM d yy"}));
        }

        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "List of people.json", options);
        buildReport(doc, dataSource, "persons");

        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataString.docx");
    }

    @Test
    public void jsonDataStream() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.DataSource.Java.docx");

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
    }

    @Test
    public void jsonDataWithNestedElements() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.DataSourceWithNestedElements.Java.docx");

        JsonDataSource dataSource = new JsonDataSource(getMyDir() + "Nested elements.json");
        buildReport(doc, dataSource, "managers");

        doc.save(getArtifactsDir() + "ReportingEngine.JsonDataWithNestedElements.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "ReportingEngine.JsonDataWithNestedElements.docx",
                getGoldsDir() + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
    }

    @Test
    public void csvDataString() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.CsvData.Java.docx");

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
    public void csvDataStream() throws Exception {
        Document doc = new Document(getMyDir() + "ReportingEngine.CsvData.Java.docx");

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