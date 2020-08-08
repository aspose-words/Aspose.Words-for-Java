package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.List;
import com.aspose.words.Shape;
import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataColumn;
import com.aspose.words.net.System.Data.DataTable;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.MessageFormat;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.Locale;
import java.util.regex.Pattern;

public class ExField extends ApiExampleBase {
    @Test
    public void updateTOC() throws Exception {
        Document doc = new Document();
        doc.updateFields();
    }

    @Test
    public void getFieldFromDocument() throws Exception {
        //ExStart
        //ExFor:FieldType
        //ExFor:FieldChar
        //ExFor:FieldChar.FieldType
        //ExFor:FieldChar.IsDirty
        //ExFor:FieldChar.IsLocked
        //ExFor:FieldChar.GetField
        //ExFor:Field.IsLocked
        //ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldDate field = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        field.getFormat().setDateTimeFormat("dddd, MMMM dd, yyyy");
        field.update();

        FieldChar fieldStart = field.getStart();
        Assert.assertEquals(FieldType.FIELD_DATE, fieldStart.getFieldType());
        Assert.assertEquals(false, fieldStart.isDirty());
        Assert.assertEquals(false, fieldStart.isLocked());

        // Retrieve the facade object which represents the field in the document
        field = (FieldDate) fieldStart.getField();

        Assert.assertEquals(false, field.isLocked());
        Assert.assertEquals(" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field.getFieldCode());

        // This updates only this field in the document
        field.update();
        //ExEnd
    }

    @Test
    public void getFieldCode() throws Exception {
        //ExStart
        //ExFor:Field.GetFieldCode
        //ExFor:Field.GetFieldCode(bool)
        //ExSummary:Shows how to get text between field start and field separator (or field end if there is no separator).
        // Open a document which contains a MERGEFIELD inside an IF field
        Document doc = new Document(getMyDir() + "Nested fields.docx");
        Assert.assertEquals(1, DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_IF)); //ExSkip

        // Get the outer IF field and print its full field code
        FieldIf fieldIf = (FieldIf) doc.getRange().getFields().get(0);
        System.out.println("Full field code including child fields:\n\t{fieldIf.GetFieldCode()}");

        // All inner nested fields are printed by default
        Assert.assertEquals(fieldIf.getFieldCode(), fieldIf.getFieldCode(true));

        // Print the field code again but this time without the inner MERGEFIELD
        System.out.println("Field code with nested fields omitted:\n\t{fieldIf.GetFieldCode(false)}");
        //ExEnd

        Assert.assertEquals(" IF  > 0 \" (surplus of ) \" \"\" ", fieldIf.getFieldCode(false));
        Assert.assertEquals(MessageFormat.format(" IF {0} MERGEFIELD NetIncome {1}{2} > 0 \" (surplus of {3} MERGEFIELD  NetIncome \\f $ {4}{5}) \" \"\" ", ControlChar.FIELD_START_CHAR, ControlChar.FIELD_SEPARATOR_CHAR, ControlChar.FIELD_END_CHAR, ControlChar.FIELD_START_CHAR, ControlChar.FIELD_SEPARATOR_CHAR, ControlChar.FIELD_END_CHAR),
                fieldIf.getFieldCode(true));
    }

    @Test
    public void fieldDisplayResult() throws Exception {
        //ExStart
        //ExFor:Field.DisplayResult
        //ExSummary:Shows how to get the text that represents the displayed field result.
        Document document = new Document(getMyDir() + "Various fields.docx");

        FieldCollection fields = document.getRange().getFields();

        Assert.assertEquals("111", fields.get(0).getDisplayResult());
        Assert.assertEquals("222", fields.get(1).getDisplayResult());
        Assert.assertEquals("Multi\rLine\rText", fields.get(2).getDisplayResult());
        Assert.assertEquals("%", fields.get(3).getDisplayResult());
        Assert.assertEquals("Macro Button Text", fields.get(4).getDisplayResult());
        Assert.assertEquals("", fields.get(5).getDisplayResult());

        // Method must be called to obtain correct value for the "FieldListNum", "FieldAutoNum",
        // "FieldAutoNumOut" and "FieldAutoNumLgl" fields
        document.updateListLabels();

        Assert.assertEquals("1)", fields.get(5).getDisplayResult());
        //ExEnd
    }

    @Test
    public void createWithFieldBuilder() throws Exception {
        //ExStart
        //ExFor:FieldBuilder.#ctor(FieldType)
        //ExFor:FieldBuilder.BuildAndInsert(Inline)
        //ExSummary:Builds and inserts a field into the document before the specified inline node.
        Document doc = new Document();

        // A convenient way of adding text content to a document is with a DocumentBuilder
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write(" Hello world! This text is one Run, which is an inline node.");

        // Fields can be constructed in a similar way with a FieldBuilder, with arguments and switches added individually
        // In this case we will construct a BARCODE field which represents a US postal code
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_BARCODE);
        fieldBuilder.addArgument("90210");
        fieldBuilder.addSwitch("\\f", "A");
        fieldBuilder.addSwitch("\\u");

        // Insert the field before any inline node
        fieldBuilder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0));
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CreateWithFieldBuilder.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.CreateWithFieldBuilder.docx");

        TestUtil.verifyField(FieldType.FIELD_BARCODE, " BARCODE 90210 \\f A \\u ", "", doc.getRange().getFields().get(0));

        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(11).getPreviousSibling(), doc.getRange().getFields().get(0).getEnd());
        Assert.assertEquals(MessageFormat.format("BARCODE 90210 \\f A \\u {0} Hello world! This text is one Run, which is an inline node.", ControlChar.FIELD_END_CHAR),
                doc.getText().trim());
    }

    @Test
    public void createRevNumFieldByDocumentBuilder() throws Exception {
        //ExStart
        //ExFor:FieldRevNum
        //ExSummary:Shows how to work with REVNUM fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to a blank document with a DocumentBuilder
        builder.write("Current revision #");

        // Insert a REVNUM field, which displays the document's current revision number property
        FieldRevNum field = (FieldRevNum) builder.insertField(FieldType.FIELD_REVISION_NUM, true);

        Assert.assertEquals(" REVNUM ", field.getFieldCode());
        Assert.assertEquals("1", field.getResult());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getRevisionNumber());

        // This property counts how many times a document has been saved in Microsoft Word, is unrelated to revision tracking,
        // can be found by right clicking the document in Windows Explorer via Properties > Details
        // This property is only manually updated by Aspose.Words
        doc.getBuiltInDocumentProperties().setRevisionNumber(doc.getBuiltInDocumentProperties().getRevisionNumber() + 1)/*Property++*/;
        Assert.assertEquals("1", field.getResult()); //ExSkip
        field.update();

        Assert.assertEquals("2", field.getResult());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        Assert.assertEquals(2, doc.getBuiltInDocumentProperties().getRevisionNumber());

        TestUtil.verifyField(FieldType.FIELD_REVISION_NUM, " REVNUM ", "2", doc.getRange().getFields().get(0));
    }

    @Test
    public void createInfoFieldWithFieldBuilder() throws Exception {
        Document doc = new Document();
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_INFO);
        fieldBuilder.buildAndInsert(run);

        doc.updateFields();
        doc = DocumentHelper.saveOpen(doc);

        FieldInfo info = (FieldInfo) doc.getRange().getFields().get(0);
        Assert.assertNotNull(info);
    }

    @Test
    public void createInfoFieldWithDocumentBuilder() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("INFO MERGEFORMAT");

        doc = DocumentHelper.saveOpen(doc);

        FieldInfo info = (FieldInfo) doc.getRange().getFields().get(0);
        Assert.assertNotNull(info);
    }

    @Test
    public void getFieldFromFieldCollection() throws Exception {
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        Field field = doc.getRange().getFields().get(0);

        // This should be the first field in the document - a TOC field
        System.out.println(field.getType());
    }

    @Test
    public void insertFieldNone() throws Exception {
        //ExStart
        //ExFor:FieldUnknown
        //ExSummary:Shows how to work with 'FieldNone' field in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field that does not denote a real field type in its field code
        Field field = builder.insertField(" NOTAREALFIELD //a");

        // Fields like that can be written and read, and are assigned a special "FieldNone" type
        Assert.assertEquals(FieldType.FIELD_NONE, field.getType());

        // We can also still work with these fields, and assign them as instances of a special "FieldUnknown" class
        FieldUnknown fieldUnknown = (FieldUnknown) field;
        Assert.assertEquals(" NOTAREALFIELD //a", fieldUnknown.getFieldCode());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.verifyField(FieldType.FIELD_NONE, " NOTAREALFIELD //a", "Error! Bookmark not defined.", doc.getRange().getFields().get(0));
    }

    @Test
    public void insertTcField() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TC field at the current document builder position
        builder.insertField("TC \"Entry Text\" \\f t");
    }

    @Test
    public void fieldLocale() throws Exception {
        //ExStart
        //ExFor:Field.LocaleId
        //ExSummary:Shows how to insert a field and work with its locale.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DATE field and print the date it will display, formatted according to your thread's current culture
        Field field = builder.insertField("DATE");
        System.out.println(MessageFormat.format("Today's date, as displayed in the \"{0}\" culture: {1}", Locale.getDefault().getDisplayName(), field.getResult()));

        Assert.assertEquals(1033, field.getLocaleId());
        Assert.assertEquals(FieldUpdateCultureSource.CURRENT_THREAD, doc.getFieldOptions().getFieldUpdateCultureSource()); //ExSkip

        // We can get the field to display a date in a different format if we change the current thread's culture
        // If we want to avoid making such an all encompassing change,
        // we can set this option to get the document's fields to get their culture from themselves
        // Then, we can change a field's LocaleId and it will display its result in any culture we choose
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        field.setLocaleId(1031);
        field.update();

        System.out.println(MessageFormat.format("Today's date, as displayed according to the \"{0}\" culture: {1}", field.getLocaleId(), field.getResult()));
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        field = doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DATE, "DATE", LocalDate.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy")), field);
        Assert.assertEquals(1031, field.getLocaleId());
    }

    @Test
    public void changeLocale() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete
        Locale currentCulture = Locale.getDefault();
        // Set to German language so dates and numbers are formatted using this culture during mail merge
        Locale.setDefault(new Locale("de", "DE"));

        // Execute mail merge
        doc.getMailMerge().execute(new String[]{"Date"}, new Object[]{new Date()});

        // Restore the original culture
        Locale.setDefault(currentCulture);

        doc.save(getArtifactsDir() + "Field.ChangeLocale.docx");
    }

    @Test
    public void removeTocFromDocument() throws Exception {
        // Open a document which contains a TOC
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        // Remove the first TOC from the document
        Field tocField = doc.getRange().getFields().get(0);
        tocField.remove();

        doc.save(getArtifactsDir() + "Field.RemoveTocFromDocument.docx");
    }

    @Test
    public void insertTcFieldsAtText() throws Exception {
        Document doc = new Document();

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new InsertTcFieldHandler("Chapter 1", "\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document
        doc.getRange().replace(Pattern.compile("The Beginning"), "", options);
    }

    private static class InsertTcFieldHandler implements IReplacingCallback {
        // Store the text and switches to be used for the TC fields
        private String mFieldText;
        private String mFieldSwitches;

        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty String or null.
        /// </summary>
        public InsertTcFieldHandler(String text, String switches) {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        public int replacing(final ReplacingArgs args) throws Exception {
            // Create a builder to insert the field
            DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
            // Move to the first node of the match
            builder.moveTo(args.getMatchNode());

            // If the user specified text to be used in the field as display text then use that, otherwise use the
            // match string as the display text
            String insertText;

            if (!(mFieldText == null || "".equals(mFieldText))) {
                insertText = mFieldText;
            } else {
                insertText = args.getMatch().group();
            }

            // Insert the TC field before this node using the specified string as the display text and user defined switches
            builder.insertField(MessageFormat.format("TC \"{0}\" {1}", insertText, mFieldSwitches));

            // We have done what we want so skip replacement
            return ReplaceAction.SKIP;
        }
    }

    @Test(enabled = false, description = "WORDSNET-16037", dataProvider = "updateDirtyFieldsDataProvider")
    public void updateDirtyFields(boolean doUpdateDirtyFields) throws Exception {
        //ExStart
        //ExFor:Field.IsDirty
        //ExFor:LoadOptions.UpdateDirtyFields
        //ExSummary:Shows how to use special property for updating field result.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Give the document's built in property "Author" a value and display it with a field
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        FieldAuthor field = (FieldAuthor) builder.insertField(FieldType.FIELD_AUTHOR, true);

        Assert.assertFalse(field.isDirty());
        Assert.assertEquals("John Doe", field.getResult());

        // Update the "Author" property
        doc.getBuiltInDocumentProperties().setAuthor("John & Jane Doe");

        // AUTHOR is one of the field types whose fields do not update according to their source values in real time,
        // and need to be updated manually beforehand every time an accurate value is required
        Assert.assertEquals("John Doe", field.getResult());

        // Since the field's value is out of date, we can mark it as "Dirty"
        field.isDirty(true);

        OutputStream docStream = new FileOutputStream(getArtifactsDir() + "Filed.UpdateDirtyFields.docx");
        try {
            doc.save(docStream, SaveFormat.DOCX);

            // Re-open the document from the stream while using a LoadOptions object to specify
            // whether to update all fields marked as "Dirty" in the process, so they can display accurate values immediately
            LoadOptions options = new LoadOptions();
            options.setUpdateDirtyFields(doUpdateDirtyFields);

            doc = new Document(String.valueOf(docStream), options);

            Assert.assertEquals("John & Jane Doe", doc.getBuiltInDocumentProperties().getAuthor());

            field = (FieldAuthor) doc.getRange().getFields().get(0);

            if (doUpdateDirtyFields) {
                Assert.assertEquals("John & Jane Doe", field.getResult());
                Assert.assertFalse(field.isDirty());
            } else {
                Assert.assertEquals("John Doe", field.getResult());
                Assert.assertTrue(field.isDirty());
            }
        } finally {
            if (docStream != null) docStream.close();
        }
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "updateDirtyFieldsDataProvider")
    public static Object[][] updateDirtyFieldsDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void insertFieldWithFieldBuilderException() throws Exception {
        Document doc = new Document();

        // Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
        argumentBuilder.addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD));
        argumentBuilder.addNode(run);
        argumentBuilder.addText("Text argument builder");

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_INCLUDE_TEXT);

        Assert.assertThrows(IllegalArgumentException.class, () -> fieldBuilder.addArgument(argumentBuilder).addArgument("=").addArgument("BestField")
                .addArgument(10).addArgument(20.0).buildAndInsert(run));
    }

    @Test
    public void barcodeGenerator() throws Exception {
        //ExStart
        //ExFor:BarcodeParameters
        //ExFor:BarcodeParameters.AddStartStopChar
        //ExFor:BarcodeParameters.BackgroundColor
        //ExFor:BarcodeParameters.BarcodeType
        //ExFor:BarcodeParameters.BarcodeValue
        //ExFor:BarcodeParameters.CaseCodeStyle
        //ExFor:BarcodeParameters.DisplayText
        //ExFor:BarcodeParameters.ErrorCorrectionLevel
        //ExFor:BarcodeParameters.FacingIdentificationMark
        //ExFor:BarcodeParameters.FixCheckDigit
        //ExFor:BarcodeParameters.ForegroundColor
        //ExFor:BarcodeParameters.IsBookmark
        //ExFor:BarcodeParameters.IsUSPostalAddress
        //ExFor:BarcodeParameters.PosCodeStyle
        //ExFor:BarcodeParameters.PostalAddress
        //ExFor:BarcodeParameters.ScalingFactor
        //ExFor:BarcodeParameters.SymbolHeight
        //ExFor:BarcodeParameters.SymbolRotation
        //ExFor:IBarcodeGenerator
        //ExFor:IBarcodeGenerator.GetBarcodeImage(BarcodeParameters)
        //ExFor:IBarcodeGenerator.GetOldBarcodeImage(BarcodeParameters)
        //ExFor:FieldOptions.BarcodeGenerator
        //ExSummary:Shows how to create barcode images using a barcode generator.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertNull(doc.getFieldOptions().getBarcodeGenerator());

        // Barcodes generated in this way will be images, and we can use a custom IBarcodeGenerator implementation to generate them
        doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());

        // Configure barcode parameters for a QR barcode
        BarcodeParameters barcodeParameters = new BarcodeParameters();
        barcodeParameters.setBarcodeType("QR");
        barcodeParameters.setBarcodeValue("ABC123");
        barcodeParameters.setBackgroundColor("0xF8BD69");
        barcodeParameters.setForegroundColor("0xB5413B");
        barcodeParameters.setErrorCorrectionLevel("3");
        barcodeParameters.setScalingFactor("250");
        barcodeParameters.setSymbolHeight("1000");
        barcodeParameters.setSymbolRotation("0");

        // Save the generated barcode image to the file system
        BufferedImage img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "Field.BarcodeGenerator.QR.jpg"));

        // Insert the image into the document
        builder.insertImage(img);

        // Configure barcode parameters for a EAN13 barcode
        barcodeParameters = new BarcodeParameters();
        barcodeParameters.setBarcodeType("EAN13");
        barcodeParameters.setBarcodeValue("501234567890");
        barcodeParameters.setDisplayText(true);
        barcodeParameters.setPosCodeStyle("CASE");
        barcodeParameters.setFixCheckDigit(true);

        img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "Field.BarcodeGenerator.EAN13.jpg"));
        builder.insertImage(img);

        // Configure barcode parameters for a CODE39 barcode
        barcodeParameters = new BarcodeParameters();
        barcodeParameters.setBarcodeType("CODE39");
        barcodeParameters.setBarcodeValue("12345ABCDE");
        barcodeParameters.setAddStartStopChar(true);

        img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "Field.BarcodeGenerator.CODE39.jpg"));
        builder.insertImage(img);

        // Configure barcode parameters for an ITF14 barcode
        barcodeParameters = new BarcodeParameters();
        barcodeParameters.setBarcodeType("ITF14");
        barcodeParameters.setBarcodeValue("09312345678907");
        barcodeParameters.setCaseCodeStyle("STD");

        img = doc.getFieldOptions().getBarcodeGenerator().getBarcodeImage(barcodeParameters);
        ImageIO.write(img, "jpg", new File(getArtifactsDir() + "Field.BarcodeGenerator.ITF14.jpg"));
        builder.insertImage(img);

        doc.save(getArtifactsDir() + "Field.BarcodeGenerator.docx");
        //ExEnd

        TestUtil.verifyImage(378, 378, getArtifactsDir() + "Field.BarcodeGenerator.QR.jpg");
        TestUtil.verifyImage(220, 78, getArtifactsDir() + "Field.BarcodeGenerator.EAN13.jpg");
        TestUtil.verifyImage(414, 65, getArtifactsDir() + "Field.BarcodeGenerator.CODE39.jpg");
        TestUtil.verifyImage(300, 65, getArtifactsDir() + "Field.BarcodeGenerator.ITF14.jpg");

        doc = new Document(getArtifactsDir() + "Field.BarcodeGenerator.docx");
        Shape barcode = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertTrue(barcode.hasImage());
    }

    @Test(enabled = false, description = "WORDSNET-13854")
    public void fieldDatabase() throws Exception {
        //ExStart
        //ExFor:FieldDatabase
        //ExFor:FieldDatabase.Connection
        //ExFor:FieldDatabase.FileName
        //ExFor:FieldDatabase.FirstRecord
        //ExFor:FieldDatabase.FormatAttributes
        //ExFor:FieldDatabase.InsertHeadings
        //ExFor:FieldDatabase.InsertOnceOnMailMerge
        //ExFor:FieldDatabase.LastRecord
        //ExFor:FieldDatabase.Query
        //ExFor:FieldDatabase.TableFormat
        //ExSummary:Shows how to extract data from a database and insert it as a field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a database field
        FieldDatabase field = (FieldDatabase) builder.insertField(FieldType.FIELD_DATABASE, true);

        // Create a simple query that extracts one table from the database
        field.setFileName(getDatabaseDir() + "Northwind.mdb");
        field.setConnection("DSN=MS Access Databases");
        field.setQuery("SELECT * FROM [Products]");

        Assert.assertEquals(" DATABASE  \\d \"{DatabaseDir.Replace('\\', '\\\\') + 'Northwind.mdb'}\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"",
                field.getFieldCode());

        // Insert another database field
        field = (FieldDatabase) builder.insertField(FieldType.FIELD_DATABASE, true);
        field.setFileName(getMyDir() + "Database\\Northwind.mdb");
        field.setConnection("DSN=MS Access Databases");

        // This query will sort all the products by their gross sales in descending order
        field.setQuery("SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
                "FROM([Products] " +
                "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
                "GROUP BY[Products].ProductName " +
                "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC");

        // You can use these variables instead of a LIMIT or TOP clause, to simplify your query
        // In this case we are taking the first 10 values of the result of our query
        field.setFirstRecord("1");
        field.setLastRecord("10");

        // The number we put here is the index of the format we want to use for our table
        // The list of table formats is in the "Table AutoFormat..." menu we find in MS Word when we create a data table field
        // Index "10" corresponds to the "Colorful 3" format
        field.setTableFormat("10");

        // This attribute decides which elements of the table format we picked above we incorporate into our table
        // The number we use is a sum of a combination of values corresponding to which elements we choose
        // 63 represents borders (1) + shading (2) + font (4) + colour (8) + autofit (16) + heading rows (32)
        field.setFormatAttributes("63");
        field.setInsertHeadings(true);
        field.setInsertOnceOnMailMerge(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.DATABASE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DATABASE.docx");

        Assert.assertEquals(2, doc.getRange().getFields().getCount());

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(77, table.getRows().getCount());
        Assert.assertEquals(10, table.getRows().get(0).getCells().getCount());

        field = (FieldDatabase) doc.getRange().getFields().get(0);

        Assert.assertEquals(" DATABASE  \\d \"{DatabaseDir.Replace('\\', '\\\\') + 'Northwind.mdb'}\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"",
                field.getFieldCode());

        TestUtil.tableMatchesQueryResult(table, getDatabaseDir() + "Northwind.mdb", field.getQuery());

        table = (Table) doc.getChild(NodeType.TABLE, 1, true);
        field = (FieldDatabase) doc.getRange().getFields().get(1);

        Assert.assertEquals(11, table.getRows().getCount());
        Assert.assertEquals(2, table.getRows().get(0).getCells().getCount());
        Assert.assertEquals("ProductName\u0007", table.getRows().get(0).getCells().get(0).getText());
        Assert.assertEquals("GrossSales\u0007", table.getRows().get(0).getCells().get(1).getText());

        Assert.assertEquals(" DATABASE  \\d \"{DatabaseDir.Replace('\\', '\\\\') + 'Northwind.mdb'}\" \\c \"DSN=MS Access Databases\" " +
                        "\\s \"SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
                        "FROM([Products] " +
                        "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
                        "GROUP BY[Products].ProductName " +
                        "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC\" \\f 1 \\t 10 \\l 10 \\b 63 \\h \\o",
                field.getFieldCode());

        table.getRows().get(0).remove();

        TestUtil.tableMatchesQueryResult(table, getDatabaseDir() + "Northwind.mdb", new StringBuffer(field.getQuery()).insert(7, " TOP 10 ").toString());
    }

    @Test
    public void updateFieldIgnoringMergeFormat() throws Exception {
        //ExStart
        //ExFor:Field.Update(bool)
        //ExFor:LoadOptions.PreserveIncludePictureField
        //ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch.
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setPreserveIncludePictureField(true);
        }
        Document doc = new Document(getMyDir() + "Field sample - INCLUDEPICTURE.docx", loadOptions);

        for (Field field : doc.getRange().getFields()) {
            if (((field.getType()) == (FieldType.FIELD_INCLUDE_PICTURE))) {
                FieldIncludePicture includePicture = (FieldIncludePicture) field;
                includePicture.setSourceFullName(getImageDir() + "Transparent background logo.png");
                includePicture.update(true);

                doc.updateFields();
                doc.save(getArtifactsDir() + "Field.UpdateFieldIgnoringMergeFormat.docx");
                //ExEnd

                Assert.assertTrue(DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_INCLUDE_PICTURE) > 0);

                doc = new Document(getArtifactsDir() + "Field.UpdateFieldIgnoringMergeFormat.docx");
                Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

                Assert.assertTrue(shape.isImage());
                Assert.assertFalse(DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_INCLUDE_PICTURE) > 0);
            }
        }
    }

    @Test
    public void fieldFormat() throws Exception {
        //ExStart
        //ExFor:Field.Format
        //ExFor:FieldFormat
        //ExFor:FieldFormat.DateTimeFormat
        //ExFor:FieldFormat.NumericFormat
        //ExFor:FieldFormat.GeneralFormats
        //ExFor:GeneralFormat
        //ExFor:GeneralFormatCollection
        //ExFor:GeneralFormatCollection.Add(GeneralFormat)
        //ExFor:GeneralFormatCollection.Count
        //ExFor:GeneralFormatCollection.Item(Int32)
        //ExFor:GeneralFormatCollection.Remove(GeneralFormat)
        //ExFor:GeneralFormatCollection.RemoveAt(Int32)
        //ExFor:GeneralFormatCollection.GetEnumerator
        //ExSummary:Shows how to format fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert field with no format
        Field field = builder.insertField("= 2 + 3");

        // We can format our field here instead of in the field code
        FieldFormat format = field.getFormat();
        format.setNumericFormat("$###.00");
        field.update();

        Assert.assertEquals("$  5.00", field.getResult());

        // Apply a date/time format
        field = builder.insertField("DATE");
        format = field.getFormat();
        format.setDateTimeFormat("dddd, MMMM dd, yyyy");
        field.update();

        System.out.println("Today's date, in {format.DateTimeFormat} format:\n\t{field.Result}");

        // Apply 2 general formats at the same time
        field = builder.insertField("= 25 + 33");
        format = field.getFormat();
        format.getGeneralFormats().add(GeneralFormat.LOWERCASE_ROMAN);
        format.getGeneralFormats().add(GeneralFormat.UPPER);
        field.update();

        int index = 0;
        Iterator<Integer> generalFormatEnumerator = format.getGeneralFormats().iterator();
        while (generalFormatEnumerator.hasNext()) {
            System.out.println(MessageFormat.format("General format index {0}: {1}", index++, generalFormatEnumerator.toString()));
        }

        Assert.assertEquals("LVIII", field.getResult());
        Assert.assertEquals(2, format.getGeneralFormats().getCount());
        Assert.assertEquals(format.getGeneralFormats().get(0), GeneralFormat.LOWERCASE_ROMAN);

        // Removing field formats
        format.getGeneralFormats().remove(GeneralFormat.LOWERCASE_ROMAN);
        format.getGeneralFormats().removeAt(0);
        Assert.assertEquals(format.getGeneralFormats().getCount(), 0);
        field.update();

        // Our field has no general formats left and is back to default form
        Assert.assertEquals(field.getResult(), "58");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals("$###.00", doc.getRange().getFields().get(0).getFormat().getNumericFormat());
        Assert.assertEquals("$  5.00", doc.getRange().getFields().get(0).getResult());

        Assert.assertEquals(doc.getRange().getFields().get(2).getFormat().getGeneralFormats().getCount(), 0);
        Assert.assertEquals("58", doc.getRange().getFields().get(2).getResult());

    }

    @Test
    public void unlinkAllFieldsInDocument() throws Exception {
        //ExStart
        //ExFor:Document.UnlinkFields
        //ExSummary:Shows how to unlink all fields in the document.
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        doc.unlinkFields();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);
        Assert.assertEquals(paraWithFields, "Fields.Docx   Элементы указателя не найдены.     1.\r");
    }

    @Test
    public void unlinkAllFieldsInRange() throws Exception {
        //ExStart
        //ExFor:Range.UnlinkFields
        //ExSummary:Shows how to unlink all fields in range.
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        Section newSection = (Section) doc.getSections().get(0).deepClone(true);
        doc.getSections().add(newSection);

        doc.getSections().get(1).getRange().unlinkFields();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        String secWithFields = DocumentHelper.getSectionText(doc, 1);

        Assert.assertTrue(secWithFields.trim().endsWith(
                "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4."));
    }

    @Test
    public void unlinkSingleField() throws Exception {
        //ExStart
        //ExFor:Field.Unlink
        //ExSummary:Shows how to unlink specific field.
        Document doc = new Document(getMyDir() + "Linked fields.docx");
        doc.getRange().getFields().get(1).unlink();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);
        Assert.assertEquals(paraWithFields, "\u0013 FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015\r");
    }

    @Test
    public void updateTocPageNumbers() throws Exception {
        Document doc = new Document(getMyDir() + "Field sample - TOC.docx");

        Node startNode = DocumentHelper.getParagraph(doc, 2);
        Node endNode = null;

        NodeCollection paragraphCollection = doc.getChildNodes(NodeType.PARAGRAPH, true);

        for (Paragraph para : (Iterable<Paragraph>) paragraphCollection) {
            // Check all runs in the paragraph for the first page breaks
            for (Run run : para.getRuns()) {
                if (run.getText().contains(ControlChar.PAGE_BREAK)) {
                    endNode = run;
                    break;
                }
            }
        }

        if (startNode != null && endNode != null) {
            removeSequence(startNode, endNode);

            startNode.remove();
            endNode.remove();
        }

        NodeCollection fStart = doc.getChildNodes(NodeType.FIELD_START, true);

        for (FieldStart field : (Iterable<FieldStart>) fStart) {
            int fType = field.getFieldType();
            if (fType == FieldType.FIELD_TOC) {
                Paragraph para = (Paragraph) field.getAncestor(NodeType.PARAGRAPH);
                para.getRange().updateFields();
                break;
            }
        }

        doc.save(getArtifactsDir() + "Field.UpdateTocPageNumbers.docx");
    }

    private static void removeSequence(Node start, Node end) {
        Node curNode = start.nextPreOrder(start.getDocument());
        while (curNode != null && !curNode.equals(end)) {
            // Move to next node
            Node nextNode = curNode.nextPreOrder(start.getDocument());

            // Check whether current contains end node
            if (curNode.isComposite()) {
                CompositeNode curComposite = (CompositeNode) curNode;
                if (!curComposite.getChildNodes(NodeType.ANY, true).contains(end) &&
                        !curComposite.getChildNodes(NodeType.ANY, true).contains(start)) {
                    nextNode = curNode.getNextSibling();
                    curNode.remove();
                }
            } else {
                curNode.remove();
            }

            curNode = nextNode;
        }
    }

    @Test
    public void dropDownItemCollection() throws Exception {
        //ExStart
        //ExFor:Fields.DropDownItemCollection
        //ExFor:Fields.DropDownItemCollection.Add(String)
        //ExFor:Fields.DropDownItemCollection.Clear
        //ExFor:Fields.DropDownItemCollection.Contains(String)
        //ExFor:Fields.DropDownItemCollection.Count
        //ExFor:Fields.DropDownItemCollection.GetEnumerator
        //ExFor:Fields.DropDownItemCollection.IndexOf(String)
        //ExFor:Fields.DropDownItemCollection.Insert(Int32, String)
        //ExFor:Fields.DropDownItemCollection.Item(Int32)
        //ExFor:Fields.DropDownItemCollection.Remove(String)
        //ExFor:Fields.DropDownItemCollection.RemoveAt(Int32)
        //ExSummary:Shows how to insert a combo box field and manipulate the elements in its item collection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to create and populate a combo box
        String[] items = {"One", "Two", "Three"};
        FormField comboBoxField = builder.insertComboBox("DropDown", items, 0);

        // Get the list of drop down items
        DropDownItemCollection dropDownItems = comboBoxField.getDropDownItems();

        Assert.assertEquals(dropDownItems.getCount(), 3);
        Assert.assertEquals(dropDownItems.get(0), "One");
        Assert.assertEquals(dropDownItems.indexOf("Two"), 1);
        Assert.assertTrue(dropDownItems.contains("Three"));

        // We can add an item to the end of the collection or insert it at a desired index
        dropDownItems.add("Four");
        dropDownItems.insert(3, "Three and a half");
        Assert.assertEquals(dropDownItems.getCount(), 5);

        // Iterate over the collection and print every element
        Iterator<String> dropDownCollectionEnumerator = dropDownItems.iterator();
        try {
            while (dropDownCollectionEnumerator.hasNext()) {
                String currentItem = dropDownCollectionEnumerator.next();
                System.out.println(currentItem);
            }
        } finally {
            if (dropDownCollectionEnumerator != null) {
                dropDownCollectionEnumerator.remove();
            }
        }

        // We can remove elements in the same way we added them
        dropDownItems.remove("Four");
        dropDownItems.removeAt(3);
        Assert.assertFalse(dropDownItems.contains("Three and a half"));
        Assert.assertFalse(dropDownItems.contains("Four"));

        doc.save(getArtifactsDir() + "Field.DropDownItemCollection.docx");

        // Empty the collection
        dropDownItems.clear();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        dropDownItems = doc.getRange().getFormFields().get(0).getDropDownItems();

        Assert.assertEquals(0, dropDownItems.getCount());

        doc = new Document(getArtifactsDir() + "Field.DropDownItemCollection.docx");
        dropDownItems = doc.getRange().getFormFields().get(0).getDropDownItems();

        Assert.assertEquals(3, dropDownItems.getCount());
        Assert.assertEquals("One", dropDownItems.get(0));
        Assert.assertEquals("Two", dropDownItems.get(1));
        Assert.assertEquals("Three", dropDownItems.get(2));
    }

    //ExStart
    //ExFor:Fields.FieldAsk
    //ExFor:Fields.FieldAsk.BookmarkName
    //ExFor:Fields.FieldAsk.DefaultResponse
    //ExFor:Fields.FieldAsk.PromptOnceOnMailMerge
    //ExFor:Fields.FieldAsk.PromptText
    //ExFor:FieldOptions.UserPromptRespondent
    //ExFor:IFieldUserPromptRespondent
    //ExFor:IFieldUserPromptRespondent.Respond(String,String)
    //ExSummary:Shows how to create an ASK field and set its properties.
    @Test
    public void fieldAsk() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Place a field where the response to our ASK field will be placed
        FieldRef fieldRef = (FieldRef) builder.insertField(FieldType.FIELD_REF, true);
        fieldRef.setBookmarkName("MyAskField");
        builder.writeln();

        // Insert the ASK field and edit its properties, making sure to reference our REF field
        FieldAsk fieldAsk = (FieldAsk) builder.insertField(FieldType.FIELD_ASK, true);
        fieldAsk.setBookmarkName("MyAskField");
        fieldAsk.setPromptText("Please provide a response for this ASK field");
        fieldAsk.setDefaultResponse("Response from within the field.");
        fieldAsk.setPromptOnceOnMailMerge(true);
        builder.writeln();

        // ASK fields apply the default response to their respective REF fields during a mail merge
        DataTable table = new DataTable("My Table");
        table.getColumns().add("Column 1");
        table.getRows().add("Row 1");
        table.getRows().add("Row 2");

        FieldMergeField fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Column 1");

        // We can modify or override the default response in our ASK fields with a custom prompt responder, which will take place during a mail merge
        doc.getFieldOptions().setUserPromptRespondent(new MyPromptRespondent());
        doc.getMailMerge().execute(table);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.ASK.docx");

        Assert.assertEquals(" REF  MyAskField", fieldRef.getFieldCode());
        Assert.assertEquals(
                " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
                fieldAsk.getFieldCode());
        testFieldAsk(table, doc); //ExSkip
    }

    /// <summary>
    /// IFieldUserPromptRespondent implementation that appends a line to the default response of an ASK field during a mail merge.
    /// </summary>
    private static class MyPromptRespondent implements IFieldUserPromptRespondent {
        public String respond(final String promptText, final String defaultResponse) {
            return "Response from MyPromptRespondent. " + defaultResponse;
        }
    }
    //ExEnd

    private void testFieldAsk(DataTable dataTable, Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        FieldRef fieldRef = (FieldRef) DocumentHelper.getField(doc.getRange().getFields(), FieldType.FIELD_REF);
        TestUtil.verifyField(FieldType.FIELD_REF,
                " REF  MyAskField", "Response from MyPromptRespondent. Response from within the field.", fieldRef);

        FieldAsk fieldAsk = (FieldAsk) DocumentHelper.getField(doc.getRange().getFields(), FieldType.FIELD_ASK);
        TestUtil.verifyField(FieldType.FIELD_ASK,
                " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
                "Response from MyPromptRespondent. Response from within the field.", fieldAsk);

        Assert.assertEquals("MyAskField", fieldAsk.getBookmarkName());
        Assert.assertEquals("Please provide a response for this ASK field", fieldAsk.getPromptText());
        Assert.assertEquals("Response from within the field.", fieldAsk.getDefaultResponse());
        Assert.assertEquals(true, fieldAsk.getPromptOnceOnMailMerge());
    }

    @Test
    public void fieldAdvance() throws Exception {
        //ExStart
        //ExFor:Fields.FieldAdvance
        //ExFor:Fields.FieldAdvance.DownOffset
        //ExFor:Fields.FieldAdvance.HorizontalPosition
        //ExFor:Fields.FieldAdvance.LeftOffset
        //ExFor:Fields.FieldAdvance.RightOffset
        //ExFor:Fields.FieldAdvance.UpOffset
        //ExFor:Fields.FieldAdvance.VerticalPosition
        //ExSummary:Shows how to insert an advance field and edit its properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("This text is in its normal place.");
        // Create an advance field using document builder
        FieldAdvance field = (FieldAdvance) builder.insertField(FieldType.FIELD_ADVANCE, true);

        builder.write("This text is moved up and to the right.");

        Assert.assertEquals(FieldType.FIELD_ADVANCE, field.getType()); //ExSkip
        Assert.assertEquals(" ADVANCE ", field.getFieldCode()); //ExSkip
        // The second text that the builder added will now be moved
        field.setRightOffset("5");
        field.setUpOffset("5");

        Assert.assertEquals(field.getFieldCode(), " ADVANCE  \\r 5 \\u 5");
        // If we want to move text in the other direction, and try do that by using negative values for the above field members, we will get an error in our document
        // Instead, we need to specify a positive value for the opposite respective field directional variable
        field = (FieldAdvance) builder.insertField(FieldType.FIELD_ADVANCE, true);
        field.setDownOffset("5");
        field.setLeftOffset("100");

        Assert.assertEquals(field.getFieldCode(), " ADVANCE  \\d 5 \\l 100");
        // We are still on one paragraph
        Assert.assertEquals(doc.getFirstSection().getBody().getParagraphs().getCount(), 1);
        // Since we're setting horizontal and vertical positions next, we need to end the paragraph so the previous line does not get moved with the next one
        builder.writeln("This text is moved down and to the left, overlapping the previous text.");
        // This time we can also use negative values
        field = (FieldAdvance) builder.insertField(FieldType.FIELD_ADVANCE, true);
        field.setHorizontalPosition("-100");
        field.setVerticalPosition("200");

        Assert.assertEquals(field.getFieldCode(), " ADVANCE  \\x -100 \\y 200");

        builder.write("This text is in a custom position.");

        doc.save(getArtifactsDir() + "Field.ADVANCE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.ADVANCE.docx");

        field = (FieldAdvance) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_ADVANCE, " ADVANCE  \\r 5 \\u 5", "", field);
        Assert.assertEquals("5", field.getRightOffset());
        Assert.assertEquals("5", field.getUpOffset());

        field = (FieldAdvance) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_ADVANCE, " ADVANCE  \\d 5 \\l 100", "", field);
        Assert.assertEquals("5", field.getDownOffset());
        Assert.assertEquals("100", field.getLeftOffset());

        field = (FieldAdvance) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_ADVANCE, " ADVANCE  \\x -100 \\y 200", "", field);
        Assert.assertEquals("-100", field.getHorizontalPosition());
        Assert.assertEquals("200", field.getVerticalPosition());
    }


    @Test
    public void fieldAddressBlock() throws Exception {
        //ExStart
        //ExFor:Fields.FieldAddressBlock.ExcludedCountryOrRegionName
        //ExFor:Fields.FieldAddressBlock.FormatAddressOnCountryOrRegion
        //ExFor:Fields.FieldAddressBlock.IncludeCountryOrRegionName
        //ExFor:Fields.FieldAddressBlock.LanguageId
        //ExFor:Fields.FieldAddressBlock.NameAndAddressFormat
        //ExSummary:Shows how to build a field address block.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a field address block
        FieldAddressBlock field = (FieldAddressBlock) builder.insertField(FieldType.FIELD_ADDRESS_BLOCK, true);

        // Initially our field is an empty address block field with null attributes
        Assert.assertEquals(field.getFieldCode(), " ADDRESSBLOCK ");

        // Setting this to "2" will cause all countries/regions to be included, unless it is the one specified in the ExcludedCountryOrRegionName attribute
        field.setIncludeCountryOrRegionName("2");
        field.setFormatAddressOnCountryOrRegion(true);
        field.setExcludedCountryOrRegionName("United States");

        // Specify our own name and address format
        field.setNameAndAddressFormat("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");

        // By default, the language ID will be set to that of the first character of the document
        // In this case we will specify it to be English
        field.setLanguageId("1033");

        // Our field code has changed according to the attribute values that we set
        Assert.assertEquals(
                " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
                field.getFieldCode());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        field = (FieldAddressBlock) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_ADDRESS_BLOCK,
                " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
                "«AddressBlock»", field);
        Assert.assertEquals("2", field.getIncludeCountryOrRegionName());
        Assert.assertEquals(true, field.getFormatAddressOnCountryOrRegion());
        Assert.assertEquals("United States", field.getExcludedCountryOrRegionName());
        Assert.assertEquals("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>",
                field.getNameAndAddressFormat());
        Assert.assertEquals("1033", field.getLanguageId());
    }

    //ExStart
    //ExFor:FieldCollection
    //ExFor:FieldCollection.Clear
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.GetEnumerator
    //ExFor:FieldCollection.Item(Int32)
    //ExFor:FieldCollection.Remove(Field)
    //ExFor:FieldCollection.Remove(FieldStart)
    //ExFor:FieldCollection.RemoveAt(Int32)
    //ExFor:FieldStart
    //ExFor:FieldStart.Accept(DocumentVisitor)
    //ExFor:FieldSeparator
    //ExFor:FieldSeparator.Accept(DocumentVisitor)
    //ExFor:FieldEnd
    //ExFor:FieldEnd.Accept(DocumentVisitor)
    //ExFor:FieldEnd.HasSeparator
    //ExFor:Field.End
    //ExFor:Field.Remove()
    //ExFor:Field.Separator
    //ExFor:Field.Start
    //ExSummary:Shows how to work with a document's field collection.
    @Test //ExSkip
    public void fieldCollection() throws Exception {
        // Create a new document and insert some fields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder.insertField(" TIME ");
        builder.insertField(" REVNUM ");
        builder.insertField(" AUTHOR  \"John Doe\" ");
        builder.insertField(" SUBJECT \"My Subject\" ");
        builder.insertField(" QUOTE \"Hello world!\" ");
        doc.updateFields();

        // Get the collection that contains all the fields in a document
        FieldCollection fields = doc.getRange().getFields();
        Assert.assertEquals(fields.getCount(), 6);

        // Iterate over the field collection and print contents and type of every field using a custom visitor implementation
        FieldVisitor fieldVisitor = new FieldVisitor();

        Iterator<Field> fieldEnumerator = fields.iterator();

        while (fieldEnumerator.hasNext()) {
            if (fieldEnumerator.next() != null) {
                Field currentField = fieldEnumerator.next();

                currentField.getStart().accept(fieldVisitor);
                if (currentField.getSeparator() != null) {
                    currentField.getSeparator().accept(fieldVisitor);
                }
                currentField.getEnd().accept(fieldVisitor);
            } else {
                System.out.println("There are no fields in the document.");
            }
        }

        System.out.println(fieldVisitor.getText());

        // Get a field to remove itself
        fields.get(0).remove();
        Assert.assertEquals(fields.getCount(), 5);

        // Remove a field by reference
        Field lastField = fields.get(3);
        fields.remove(lastField);
        Assert.assertEquals(fields.getCount(), 4);

        // Remove a field by index
        fields.removeAt(2);
        Assert.assertEquals(fields.getCount(), 3);

        // Remove all fields from the document
        fields.clear();
        Assert.assertEquals(fields.getCount(), 0);
    }

    /// <summary>
    /// Document visitor implementation that prints field info.
    /// </summary>
    public static class FieldVisitor extends DocumentVisitor {
        public FieldVisitor() {
            mBuilder = new StringBuilder();
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        public int visitFieldStart(final FieldStart fieldStart) {
            mBuilder.append("Found field: " + fieldStart.getFieldType() + "\r\n");
            mBuilder.append("\tField code: " + fieldStart.getField().getFieldCode() + "\r\n");
            mBuilder.append("\tDisplayed as: " + fieldStart.getField().getResult() + "\r\n");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        public int visitFieldSeparator(final FieldSeparator fieldSeparator) {
            mBuilder.append("\tFound separator: " + fieldSeparator.getText() + "\r\n");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        public int visitFieldEnd(final FieldEnd fieldEnd) {
            mBuilder.append("End of field: " + fieldEnd.getFieldType() + "\r\n");

            return VisitorAction.CONTINUE;
        }

        private StringBuilder mBuilder;
    }
    //ExEnd

    @Test
    public void fieldCompare() throws Exception {
        //ExStart
        //ExFor:FieldCompare
        //ExFor:FieldCompare.ComparisonOperator
        //ExFor:FieldCompare.LeftExpression
        //ExFor:FieldCompare.RightExpression
        //ExSummary:Shows how to insert a field that compares expressions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a compare field using a document builder
        FieldCompare field = (FieldCompare) builder.insertField(FieldType.FIELD_COMPARE, true);

        // Construct a comparison statement
        field.setLeftExpression("3");
        field.setComparisonOperator("<");
        field.setRightExpression("2");

        // The compare field will print a "0" or "1" depending on the truth of its statement
        // The result of this statement is false, so a "0" will be show up in the document
        Assert.assertEquals(field.getFieldCode(), " COMPARE  3 < 2");

        builder.writeln();

        // Here a "1" will show up, because the statement is true
        field = (FieldCompare) builder.insertField(FieldType.FIELD_COMPARE, true);
        field.setLeftExpression("5");
        field.setComparisonOperator("=");
        field.setRightExpression("2 + 3");

        Assert.assertEquals(field.getFieldCode(), " COMPARE  5 = \"2 + 3\"");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.COMPARE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.COMPARE.docx");

        field = (FieldCompare) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_COMPARE, " COMPARE  3 < 2", "0", field);
        Assert.assertEquals("3", field.getLeftExpression());
        Assert.assertEquals("<", field.getComparisonOperator());
        Assert.assertEquals("2", field.getRightExpression());

        field = (FieldCompare) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_COMPARE, " COMPARE  5 = \"2 + 3\"", "1", field);
        Assert.assertEquals("5", field.getLeftExpression());
        Assert.assertEquals("=", field.getComparisonOperator());
        Assert.assertEquals("\"2 + 3\"", field.getRightExpression());
    }

    @Test
    public void fieldIf() throws Exception {
        //ExStart
        //ExFor:FieldIf
        //ExFor:FieldIf.ComparisonOperator
        //ExFor:FieldIf.EvaluateCondition
        //ExFor:FieldIf.FalseText
        //ExFor:FieldIf.LeftExpression
        //ExFor:FieldIf.RightExpression
        //ExFor:FieldIf.TrueText
        //ExFor:FieldIfComparisonResult
        //ExSummary:Shows how to insert an if field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Statement 1: ");

        // Use document builder to insert an if field
        FieldIf field = (FieldIf) builder.insertField(FieldType.FIELD_IF, true);

        // The if field will output either the TrueText or FalseText string into the document, depending on the truth of the statement
        // In this case, "0 = 1" is incorrect, so the output will be "False"
        field.setLeftExpression("0");
        field.setComparisonOperator("=");
        field.setRightExpression("1");
        field.setTrueText("True");
        field.setFalseText("False");

        Assert.assertEquals(" IF  0 = 1 True False", field.getFieldCode());
        Assert.assertEquals(FieldIfComparisonResult.FALSE, field.evaluateCondition());

        // This time, the statement is correct, so the output will be "True"
        builder.write("\nStatement 2: ");
        field = (FieldIf) builder.insertField(FieldType.FIELD_IF, true);
        field.setLeftExpression("5");
        field.setComparisonOperator("=");
        field.setRightExpression("2 + 3");
        field.setTrueText("True");
        field.setFalseText("False");

        Assert.assertEquals(" IF  5 = \"2 + 3\" True False", field.getFieldCode());
        Assert.assertEquals(FieldIfComparisonResult.TRUE, field.evaluateCondition());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.IF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.IF.docx");
        field = (FieldIf) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_IF, " IF  0 = 1 True False", "False", field);
        Assert.assertEquals("0", field.getLeftExpression());
        Assert.assertEquals("=", field.getComparisonOperator());
        Assert.assertEquals("1", field.getRightExpression());
        Assert.assertEquals("True", field.getTrueText());
        Assert.assertEquals("False", field.getFalseText());

        field = (FieldIf) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_IF, " IF  5 = \"2 + 3\" True False", "True", field);
        Assert.assertEquals("5", field.getLeftExpression());
        Assert.assertEquals("=", field.getComparisonOperator());
        Assert.assertEquals("\"2 + 3\"", field.getRightExpression());
        Assert.assertEquals("True", field.getTrueText());
        Assert.assertEquals("False", field.getFalseText());
    }

    @Test
    public void fieldAutoNum() throws Exception {
        //ExStart
        //ExFor:FieldAutoNum
        //ExFor:FieldAutoNum.SeparatorCharacter
        //ExSummary:Shows how to number paragraphs using autonum fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The two fields we insert here will be automatically numbered 1 and 2
        FieldAutoNum field = (FieldAutoNum) builder.insertField(FieldType.FIELD_AUTO_NUM, true);
        builder.writeln("\tParagraph 1.");

        Assert.assertEquals(" AUTONUM ", field.getFieldCode());

        field = (FieldAutoNum) builder.insertField(FieldType.FIELD_AUTO_NUM, true);
        builder.writeln("\tParagraph 2.");

        // Leaving the FieldAutoNum.SeparatorCharacter field null will set the separator character to '.' by default
        Assert.assertNull(field.getSeparatorCharacter());

        // The first character of the string entered here will be used as the separator character
        field.setSeparatorCharacter(":");

        Assert.assertEquals(" AUTONUM  \\s :", field.getFieldCode());

        doc.save(getArtifactsDir() + "Field.AUTONUM.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTONUM.docx");

        TestUtil.verifyField(FieldType.FIELD_AUTO_NUM, " AUTONUM ", "", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_AUTO_NUM, " AUTONUM  \\s :", "", doc.getRange().getFields().get(1));
    }

    //ExStart
    //ExFor:FieldAutoNumLgl
    //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
    //ExFor:FieldAutoNumLgl.SeparatorCharacter
    //ExSummary:Shows how to organize a document using AUTONUMLGL fields.
    @Test //ExSkip
    public void fieldAutoNumLgl() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a filler paragraph string
        String loremIpsum =
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                        "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

        // In this case our autonum legal field will number our first paragraph as "1."
        insertNumberedClause(builder, "\tHeading 1", loremIpsum, StyleIdentifier.HEADING_1);

        // Our heading style number will be 1 again, so this field will keep counting headings at a heading level of 1
        insertNumberedClause(builder, "\tHeading 2", loremIpsum, StyleIdentifier.HEADING_1);

        // Our heading style is 2, setting the paragraph numbering depth to 2, setting this field's value to "2.1."
        insertNumberedClause(builder, "\tHeading 3", loremIpsum, StyleIdentifier.HEADING_2);

        // Our heading style is 3, so we are going deeper again to "2.1.1."
        insertNumberedClause(builder, "\tHeading 4", loremIpsum, StyleIdentifier.HEADING_3);

        // Our heading style is 2, and the next field number at that level is "2.2."
        insertNumberedClause(builder, "\tHeading 5", loremIpsum, StyleIdentifier.HEADING_2);

        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_AUTO_NUM_LEGAL) {
                // By default the separator will appear as "." in the document but here it is null
                Assert.assertNull(((FieldAutoNumLgl) field).getSeparatorCharacter());

                // Change the separator character and remove trailing separators
                ((FieldAutoNumLgl) field).setSeparatorCharacter(":");
                ((FieldAutoNumLgl) field).setRemoveTrailingPeriod(true);
                Assert.assertEquals(field.getFieldCode(), " AUTONUMLGL  \\s : \\e");
            }
        }

        doc.save(getArtifactsDir() + "Field.AUTONUMLGL.docx");
        testFieldAutoNumLgl(doc); //ExSkip
    }

    /// <summary>
    /// Get a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    private void insertNumberedClause(final DocumentBuilder builder, final String heading, final String contents, final int headingStyle) throws Exception {
        // This legal field will automatically number our clauses, taking heading style level into account
        builder.insertField(FieldType.FIELD_AUTO_NUM_LEGAL, true);
        builder.getCurrentParagraph().getParagraphFormat().setStyleIdentifier(headingStyle);
        builder.writeln(heading);

        // This text will belong to the auto num legal field above it
        // It will collapse when the arrow next to the corresponding autonum legal field is clicked in MS Word
        builder.getCurrentParagraph().getParagraphFormat().setStyleIdentifier(StyleIdentifier.BODY_TEXT);
        builder.writeln(contents);
    }
    //ExEnd

    private void testFieldAutoNumLgl(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_AUTO_NUM_LEGAL) {

                FieldAutoNumLgl fieldAutoNumLgl = (FieldAutoNumLgl) field;
                TestUtil.verifyField(FieldType.FIELD_AUTO_NUM_LEGAL, " AUTONUMLGL  \\s : \\e", "", fieldAutoNumLgl);

                Assert.assertEquals(":", fieldAutoNumLgl.getSeparatorCharacter());
                Assert.assertTrue(fieldAutoNumLgl.getRemoveTrailingPeriod());
            }
        }
    }

    @Test
    public void fieldAutoNumOut() throws Exception {
        //ExStart
        //ExFor:FieldAutoNumOut
        //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The two fields that we insert here will be numbered 1 and 2
        builder.insertField(FieldType.FIELD_AUTO_NUM_OUTLINE, true);
        builder.writeln("\tParagraph 1.");
        builder.insertField(FieldType.FIELD_AUTO_NUM_OUTLINE, true);
        builder.writeln("\tParagraph 2.");

        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_AUTO_NUM_OUTLINE) {
                Assert.assertEquals(field.getFieldCode(), " AUTONUMOUT ");
            }
        }

        doc.save(getArtifactsDir() + "Field.AUTONUMOUT.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTONUMOUT.docx");

        for (Field field : doc.getRange().getFields())
            TestUtil.verifyField(FieldType.FIELD_AUTO_NUM_OUTLINE, " AUTONUMOUT ", "", field);
    }

    @Test
    public void fieldAutoText() throws Exception {
        //ExStart
        //ExFor:Fields.FieldAutoText
        //ExFor:FieldAutoText.EntryName
        //ExFor:FieldOptions.BuiltInTemplatesPaths
        //ExFor:FieldGlossary
        //ExFor:FieldGlossary.EntryName
        //ExSummary:Shows how to insert a building block into a document and display it with AUTOTEXT and GLOSSARY fields.
        Document doc = new Document();

        // Create a glossary document and add an AutoText building block
        doc.setGlossaryDocument(new GlossaryDocument());
        BuildingBlock buildingBlock = new BuildingBlock(doc.getGlossaryDocument());
        buildingBlock.setName("MyBlock");
        buildingBlock.setGallery(BuildingBlockGallery.AUTO_TEXT);
        buildingBlock.setCategory("General");
        buildingBlock.setDescription("MyBlock description");
        buildingBlock.setBehavior(BuildingBlockBehavior.PARAGRAPH);
        doc.getGlossaryDocument().appendChild(buildingBlock);

        // Create a source and add it as text content to our building block
        Document buildingBlockSource = new Document();
        DocumentBuilder buildingBlockSourceBuilder = new DocumentBuilder(buildingBlockSource);
        buildingBlockSourceBuilder.writeln("Hello World!");

        Node buildingBlockContent = doc.getGlossaryDocument().importNode(buildingBlockSource.getFirstSection(), true);
        buildingBlock.appendChild(buildingBlockContent);

        // Create an advance field using document builder
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldAutoText fieldAutoText = (FieldAutoText) builder.insertField(FieldType.FIELD_AUTO_TEXT, true);

        // Refer to our building block by name
        fieldAutoText.setEntryName("MyBlock");

        Assert.assertEquals(fieldAutoText.getFieldCode(), " AUTOTEXT  MyBlock");

        // Put additional templates here
        doc.getFieldOptions().setBuiltInTemplatesPaths(new String[]{getMyDir() + "Busniess brochure.dotx"});

        // We can also display our building block with a GLOSSARY field
        FieldGlossary fieldGlossary = (FieldGlossary) builder.insertField(FieldType.FIELD_GLOSSARY, true);
        fieldGlossary.setEntryName("MyBlock");

        Assert.assertEquals(fieldGlossary.getFieldCode(), " GLOSSARY  MyBlock");

        // The text content of our building block will be visible in the output
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.AUTOTEXT.dotx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTOTEXT.dotx");

        Assert.assertEquals(doc.getFieldOptions().getBuiltInTemplatesPaths(), new String[0]);

        fieldAutoText = (FieldAutoText) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_AUTO_TEXT, " AUTOTEXT  MyBlock", "Hello World!\r", fieldAutoText);
        Assert.assertEquals("MyBlock", fieldAutoText.getEntryName());

        fieldGlossary = (FieldGlossary) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_GLOSSARY, " GLOSSARY  MyBlock", "Hello World!\r", fieldGlossary);
        Assert.assertEquals("MyBlock", fieldGlossary.getEntryName());
    }

    //ExStart
    //ExFor:Fields.FieldAutoTextList
    //ExFor:Fields.FieldAutoTextList.EntryName
    //ExFor:Fields.FieldAutoTextList.ListStyle
    //ExFor:Fields.FieldAutoTextList.ScreenTip
    //ExSummary:Shows how to use an AutoTextList field to select from a list of AutoText entries.
    @Test //ExSkip
    public void fieldAutoTextList() throws Exception {
        Document doc = new Document();

        // Create a glossary document and populate it with auto text entries that our auto text list will let us select from
        doc.setGlossaryDocument(new GlossaryDocument());
        appendAutoTextEntry(doc.getGlossaryDocument(), "AutoText 1", "Contents of AutoText 1");
        appendAutoTextEntry(doc.getGlossaryDocument(), "AutoText 2", "Contents of AutoText 2");
        appendAutoTextEntry(doc.getGlossaryDocument(), "AutoText 3", "Contents of AutoText 3");

        // Insert an auto text list using a document builder and change its properties
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldAutoTextList field = (FieldAutoTextList) builder.insertField(FieldType.FIELD_AUTO_TEXT_LIST, true);
        // This is the text that will be visible in the document
        field.setEntryName("Right click here to pick an AutoText block");
        field.setListStyle("Heading 1");
        field.setScreenTip("Hover tip text for AutoTextList goes here");

        Assert.assertEquals(" AUTOTEXTLIST  \"Right click here to pick an AutoText block\" " +
                "\\s \"Heading 1\" " +
                "\\t \"Hover tip text for AutoTextList goes here\"", field.getFieldCode());

        doc.save(getArtifactsDir() + "Field.AUTOTEXTLIST.dotx");
        testFieldAutoTextList(doc); //ExSkip
    }

    /// <summary>
    /// Create an AutoText entry and add it to a glossary document.
    /// </summary>
    private static void appendAutoTextEntry(final GlossaryDocument glossaryDoc, final String name, final String contents) {
        // Create building block and set it up as an auto text entry
        BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc);
        buildingBlock.setName(name);
        buildingBlock.setGallery(BuildingBlockGallery.AUTO_TEXT);
        buildingBlock.setCategory("General");
        buildingBlock.setBehavior(BuildingBlockBehavior.PARAGRAPH);

        // Add content to the building block
        Section section = new Section(glossaryDoc);
        section.appendChild(new Body(glossaryDoc));
        section.getBody().appendParagraph(contents);
        buildingBlock.appendChild(section);

        // Add auto text entry to glossary document
        glossaryDoc.appendChild(buildingBlock);
    }
    //ExEnd

    private void testFieldAutoTextList(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(3, doc.getGlossaryDocument().getCount());
        Assert.assertEquals("AutoText 1", doc.getGlossaryDocument().getBuildingBlocks().get(0).getName());
        Assert.assertEquals("Contents of AutoText 1", doc.getGlossaryDocument().getBuildingBlocks().get(0).getText().trim());
        Assert.assertEquals("AutoText 2", doc.getGlossaryDocument().getBuildingBlocks().get(1).getName());
        Assert.assertEquals("Contents of AutoText 2", doc.getGlossaryDocument().getBuildingBlocks().get(1).getText().trim());
        Assert.assertEquals("AutoText 3", doc.getGlossaryDocument().getBuildingBlocks().get(2).getName());
        Assert.assertEquals("Contents of AutoText 3", doc.getGlossaryDocument().getBuildingBlocks().get(2).getText().trim());

        FieldAutoTextList field = (FieldAutoTextList) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_AUTO_TEXT_LIST,
                " AUTOTEXTLIST  \"Right click here to pick an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"",
                "", field);
        Assert.assertEquals("Right click here to pick an AutoText block", field.getEntryName());
        Assert.assertEquals("Heading 1", field.getListStyle());
        Assert.assertEquals("Hover tip text for AutoTextList goes here", field.getScreenTip());
    }

    @Test
    public void fieldGreetingLine() throws Exception {
        //ExStart
        //ExFor:FieldGreetingLine
        //ExFor:FieldGreetingLine.AlternateText
        //ExFor:FieldGreetingLine.GetFieldNames
        //ExFor:FieldGreetingLine.LanguageId
        //ExFor:FieldGreetingLine.NameFormat
        //ExSummary:Shows how to insert a GREETINGLINE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a custom greeting field with document builder along with some content
        FieldGreetingLine field = (FieldGreetingLine) builder.insertField(FieldType.FIELD_GREETING_LINE, true);
        builder.writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!");

        // This array contains strings that correspond to column names in the data table that we will mail merge into our document
        Assert.assertEquals(0, field.getFieldNames().length);

        // To populate that array, we need to specify a format for our greeting line
        field.setNameFormat("<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> ");

        // In this case, our greeting line's field names array now has "Courtesy Title" and "Last Name"
        Assert.assertEquals(2, field.getFieldNames().length);

        // This string will cover any cases where the data in the data table is incorrect by substituting the malformed name with a string
        field.setAlternateText("Sir or Madam");

        // We can set the language ID here too
        field.setLanguageId("1033");

        Assert.assertEquals(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033",
                field.getFieldCode());

        // Create a source table for our mail merge that has columns that our greeting line will look for
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");
        // This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text
        table.getRows().add("", "No", "Name");

        doc.getMailMerge().execute(table);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.GREETINGLINE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.GREETINGLINE.docx");

        Assert.assertEquals(doc.getRange().getFields().getCount(), 0);
        Assert.assertEquals("Dear Mr. Doe,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                        "\fDear Mrs. Cardholder,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                        "\fDear Sir or Madam,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!",
                doc.getText().trim());
    }

    @Test
    public void fieldListNum() throws Exception {
        //ExStart
        //ExFor:FieldListNum
        //ExFor:FieldListNum.HasListName
        //ExFor:FieldListNum.ListLevel
        //ExFor:FieldListNum.ListName
        //ExFor:FieldListNum.StartingNumber
        //ExSummary:Shows how to number paragraphs with LISTNUM fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a list num field using a document builder
        FieldListNum field = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // Lists start counting at 1 by default, but we can change this number at any time
        // In this case, we'll do a zero-based count
        field.setStartingNumber("0");
        builder.writeln("Paragraph 1");

        Assert.assertEquals(" LISTNUM  \\s 0", field.getFieldCode());

        // Placing several list num fields in one paragraph increases the list level instead of the current number,
        // in this case resulting in "1)a)i)", list level 3
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.writeln("Paragraph 2");

        // The list level resets with new paragraphs, so to keep counting at a desired list level, we need to set the ListLevel property accordingly
        field = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);
        field.setListLevel("3");
        builder.writeln("Paragraph 3");

        Assert.assertEquals(" LISTNUM  \\l 3", field.getFieldCode());

        field = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // Setting this property to this particular value will emulate the AUTONUMOUT field
        field.setListName("OutlineDefault");

        Assert.assertTrue(field.hasListName());
        Assert.assertEquals(" LISTNUM  OutlineDefault", field.getFieldCode());

        // Start counting from 1
        field.setStartingNumber("1");
        builder.writeln("Paragraph 4");

        // Our fields keep track of the count automatically, but the ListName needs to be set with each new field
        field = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);
        field.setListName("OutlineDefault");
        builder.writeln("Paragraph 5");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LISTNUM.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.LISTNUM.docx");

        Assert.assertEquals(7, doc.getRange().getFields().getCount());

        field = (FieldListNum) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM  \\s 0", "", field);
        Assert.assertEquals("0", field.getStartingNumber());
        Assert.assertNull(field.getListLevel());
        Assert.assertFalse(field.hasListName());
        Assert.assertNull(field.getListName());

        for (int i = 1; i < 4; i++) {
            field = (FieldListNum) doc.getRange().getFields().get(i);

            TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM ", "", field);
            Assert.assertNull(field.getStartingNumber());
            Assert.assertNull(field.getListLevel());
            Assert.assertFalse(field.hasListName());
            Assert.assertNull(field.getListName());
        }

        field = (FieldListNum) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM  \\l 3", "", field);
        Assert.assertNull(field.getStartingNumber());
        Assert.assertEquals("3", field.getListLevel());
        Assert.assertFalse(field.hasListName());
        Assert.assertNull(field.getListName());

        field = (FieldListNum) doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM  OutlineDefault \\s 1", "", field);
        Assert.assertEquals("1", field.getStartingNumber());
        Assert.assertNull(field.getListLevel());
        Assert.assertTrue(field.hasListName());
        Assert.assertEquals("OutlineDefault", field.getListName());
    }

    @Test
    public void mergeField() throws Exception {
        //ExStart
        //ExFor:FieldMergeField
        //ExFor:FieldMergeField.FieldName
        //ExFor:FieldMergeField.FieldNameNoPrefix
        //ExFor:FieldMergeField.IsMapped
        //ExFor:FieldMergeField.IsVerticalFormatting
        //ExFor:FieldMergeField.TextAfter
        //ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create data source for our merge fields
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");

        // Insert a merge field that corresponds to one of our columns and put text before and after it
        FieldMergeField fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Courtesy Title");
        fieldMergeField.isMapped(true);
        fieldMergeField.isVerticalFormatting(false);
        fieldMergeField.setTextBefore("Dear ");
        fieldMergeField.setTextAfter(" ");

        Assert.assertEquals(" MERGEFIELD  \"Courtesy Title\" \\m \\b \"Dear \" \\f \" \"", fieldMergeField.getFieldCode());

        // Insert another merge field for another column
        // We don't need to use every column to perform a mail merge
        fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Last Name");
        fieldMergeField.setTextAfter(":");

        doc.updateFields();
        doc.getMailMerge().execute(table);
        doc.save(getArtifactsDir() + "Field.MERGEFIELD.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEFIELD.docx");

        Assert.assertEquals(doc.getRange().getFields().getCount(), 0);
        Assert.assertEquals("Dear Mr. Doe:\fDear Mrs. Cardholder:", doc.getText().trim());
    }

    //ExStart
    //ExFor:FieldToc
    //ExFor:FieldToc.BookmarkName
    //ExFor:FieldToc.CustomStyles
    //ExFor:FieldToc.EntrySeparator
    //ExFor:FieldToc.HeadingLevelRange
    //ExFor:FieldToc.HideInWebLayout
    //ExFor:FieldToc.InsertHyperlinks
    //ExFor:FieldToc.PageNumberOmittingLevelRange
    //ExFor:FieldToc.PreserveLineBreaks
    //ExFor:FieldToc.PreserveTabs
    //ExFor:FieldToc.UpdatePageNumbers
    //ExFor:FieldToc.UseParagraphOutlineLevel
    //ExFor:FieldOptions.CustomTocStyleSeparator
    //ExSummary:Shows how to insert a TOC and populate it with entries based on heading styles.
    @Test //ExSkip
    public void fieldToc() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The table of contents we will insert will accept entries that are only within the scope of this bookmark
        builder.startBookmark("MyBookmark");

        // Insert a list num field using a document builder
        FieldToc field = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);

        // Limit possible TOC entries to only those within the bookmark we name here
        field.setBookmarkName("MyBookmark");

        // Normally paragraphs with a "Heading n" style will be the only ones that will be added to a TOC as entries
        // We can set this attribute to include other styles, such as "Quote" and "Intense Quote" in this case
        field.setCustomStyles("Quote; 6; Intense Quote; 7");

        // Styles are normally separated by a comma (",") but we can use this property to set a custom delimiter
        doc.getFieldOptions().setCustomTocStyleSeparator(";");

        // Filter out any headings that are outside this range
        field.setHeadingLevelRange("1-3");

        // Headings in this range won't display their page number in their TOC entry
        field.setPageNumberOmittingLevelRange("2-5");

        field.setEntrySeparator("-");
        field.setInsertHyperlinks(true);
        field.setHideInWebLayout(false);
        field.setPreserveLineBreaks(true);
        field.setPreserveTabs(true);
        field.setUseParagraphOutlineLevel(false);

        insertNewPageWithHeading(builder, "First entry", "Heading 1");
        builder.writeln("Paragraph text.");
        insertNewPageWithHeading(builder, "Second entry", "Heading 1");
        insertNewPageWithHeading(builder, "Third entry", "Quote");
        insertNewPageWithHeading(builder, "Fourth entry", "Intense Quote");

        // These two headings will have the page numbers omitted because they are within the "2-5" range
        insertNewPageWithHeading(builder, "Fifth entry", "Heading 2");
        insertNewPageWithHeading(builder, "Sixth entry", "Heading 3");

        // This entry will be omitted because "Heading 4" is outside of the "1-3" range we set earlier
        insertNewPageWithHeading(builder, "Seventh entry", "Heading 4");

        builder.endBookmark("MyBookmark");
        builder.writeln("Paragraph text.");

        // This entry will be omitted because it is outside the bookmark specified by the TOC
        insertNewPageWithHeading(builder, "Eighth entry", "Heading 1");

        Assert.assertEquals(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.getFieldCode());

        field.updatePageNumbers();
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOC.docx");

        testFieldToc(doc); //ExSkip
    }

    /// <summary>
    /// Start a new page and insert a paragraph of a specified style.
    /// </summary>
    @Test(enabled = false)
    public void insertNewPageWithHeading(final DocumentBuilder builder, final String captionText, final String styleName) {
        builder.insertBreak(BreakType.PAGE_BREAK);
        String originalStyle = builder.getParagraphFormat().getStyleName();
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get(styleName));
        builder.writeln(captionText);
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get(originalStyle));
    }
    //ExEnd

    private void testFieldToc(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);
        FieldToc field = (FieldToc) doc.getRange().getFields().get(0);

        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertEquals("Quote; 6; Intense Quote; 7", field.getCustomStyles());
        Assert.assertEquals("-", field.getEntrySeparator());
        Assert.assertEquals("1-3", field.getHeadingLevelRange());
        Assert.assertEquals("2-5", field.getPageNumberOmittingLevelRange());
        Assert.assertFalse(field.getHideInWebLayout());
        Assert.assertTrue(field.getInsertHyperlinks());
        Assert.assertTrue(field.getPreserveLineBreaks());
        Assert.assertTrue(field.getPreserveTabs());
        Assert.assertTrue(field.updatePageNumbers());
        Assert.assertFalse(field.getUseParagraphOutlineLevel());
        Assert.assertEquals(" TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w", field.getFieldCode());
        Assert.assertEquals("\u0013 HYPERLINK \\l \"_Toc256000001\" \u0014First entry-\u0013 PAGEREF _Toc256000001 \\h \u00142\u0015\u0015\r" +
                "\u0013 HYPERLINK \\l \"_Toc256000002\" \u0014Second entry-\u0013 PAGEREF _Toc256000002 \\h \u00143\u0015\u0015\r" +
                "\u0013 HYPERLINK \\l \"_Toc256000003\" \u0014Third entry-\u0013 PAGEREF _Toc256000003 \\h \u00144\u0015\u0015\r" +
                "\u0013 HYPERLINK \\l \"_Toc256000004\" \u0014Fourth entry-\u0013 PAGEREF _Toc256000004 \\h \u00145\u0015\u0015\r" +
                "\u0013 HYPERLINK \\l \"_Toc256000005\" \u0014Fifth entry\u0015\r" +
                "\u0013 HYPERLINK \\l \"_Toc256000006\" \u0014Sixth entry\u0015\r", field.getResult());
    }

    //ExStart
    //ExFor:FieldToc.EntryIdentifier
    //ExFor:FieldToc.EntryLevelRange
    //ExFor:FieldTC
    //ExFor:FieldTC.OmitPageNumber
    //ExFor:FieldTC.Text
    //ExFor:FieldTC.TypeIdentifier
    //ExFor:FieldTC.EntryLevel
    //ExSummary:Shows how to insert a TOC field and filter which TC fields end up as entries.
    @Test //ExSkip
    public void fieldTocEntryIdentifier() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");

        // Insert a list num field using a document builder
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);
        fieldToc.setEntryIdentifier("A");
        fieldToc.setEntryLevelRange("1-3");

        Assert.assertEquals(" TOC  \\f A \\l 1-3", fieldToc.getFieldCode());

        // These two entries will appear in the table
        builder.insertBreak(BreakType.PAGE_BREAK);
        insertTocEntry(builder, "TC field 1", "A", "1");
        insertTocEntry(builder, "TC field 2", "A", "2");

        Assert.assertEquals(" TC  \"TC field 1\" \\n \\f A \\l 1", doc.getRange().getFields().get(1).getFieldCode());

        // These two entries will be omitted because of an incorrect type identifier
        insertTocEntry(builder, "TC field 3", "B", "1");

        // ...and an out-of-range entry level
        insertTocEntry(builder, "TC field 4", "A", "5");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TC.docx");
        testFieldTocEntryIdentifier(doc); //ExSkip
    }

    /// <summary>
    /// Insert a table of contents entry via a document builder.
    /// </summary>
    @Test(enabled = false)
    public void insertTocEntry(final DocumentBuilder builder, final String text, final String typeIdentifier, final String entryLevel) throws Exception {
        FieldTC fieldTc = (FieldTC) builder.insertField(FieldType.FIELD_TOC_ENTRY, true);
        fieldTc.setOmitPageNumber(true);
        fieldTc.setText(text);
        fieldTc.setTypeIdentifier(typeIdentifier);
        fieldTc.setEntryLevel(entryLevel);
    }
    //ExEnd

    private void testFieldTocEntryIdentifier(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);
        FieldToc fieldToc = (FieldToc) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_TOC, " TOC  \\f A \\l 1-3", "TC field 1\rTC field 2\r", fieldToc);
        Assert.assertEquals("A", fieldToc.getEntryIdentifier());
        Assert.assertEquals("1-3", fieldToc.getEntryLevelRange());

        FieldTC fieldTc = (FieldTC) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 1\" \\n \\f A \\l 1", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 1", fieldTc.getText());
        Assert.assertEquals("A", fieldTc.getTypeIdentifier());
        Assert.assertEquals("1", fieldTc.getEntryLevel());

        fieldTc = (FieldTC) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 2\" \\n \\f A \\l 2", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 2", fieldTc.getText());
        Assert.assertEquals("A", fieldTc.getTypeIdentifier());
        Assert.assertEquals("2", fieldTc.getEntryLevel());

        fieldTc = (FieldTC) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 3\" \\n \\f B \\l 1", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 3", fieldTc.getText());
        Assert.assertEquals("B", fieldTc.getTypeIdentifier());
        Assert.assertEquals("1", fieldTc.getEntryLevel());

        fieldTc = (FieldTC) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 4\" \\n \\f A \\l 5", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 4", fieldTc.getText());
        Assert.assertEquals("A", fieldTc.getTypeIdentifier());
        Assert.assertEquals("5", fieldTc.getEntryLevel());
    }

    @Test
    public void tocSeqPrefix() throws Exception {
        //ExStart
        //ExFor:FieldToc
        //ExFor:FieldToc.TableOfFiguresLabel
        //ExFor:FieldToc.PrefixedSequenceIdentifier
        //ExFor:FieldToc.SequenceSeparator
        //ExFor:FieldSeq
        //ExFor:FieldSeq.SequenceIdentifier
        //ExSummary:Shows how to populate a TOC field with entries using SEQ fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field that creates a table of contents entry for each paragraph
        // that contains a SEQ field with a sequence identifier of "MySequence" with the number of the page which contains that field
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);
        fieldToc.setTableOfFiguresLabel("MySequence");

        // This identifier is for a parallel SEQ sequence,
        // the number that it is at will be displayed in front of the page number of the paragraph with the other sequence,
        // separated by a sequence separator character also defined below
        fieldToc.setPrefixedSequenceIdentifier("PrefixSequence");
        fieldToc.setSequenceSeparator(">");

        Assert.assertEquals(" TOC  \\c MySequence \\s PrefixSequence \\d >", fieldToc.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert a SEQ field to increment the sequence counter of "PrefixSequence" to 1
        // Since this paragraph doesn't contain a SEQ field of the "MySequence" sequence,
        // this will not appear as an entry in the TOC
        FieldSeq fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("PrefixSequence");
        builder.insertParagraph();

        Assert.assertEquals(" SEQ  PrefixSequence", fieldSeq.getFieldCode());

        // Insert two SEQ fields, one for each of the sequences we defined above
        // The "MySequence" SEQ appears on page 2 and the "PrefixSequence" is at number 1 in this paragraph,
        // which means that our TOC will display this as an entry with the contents on the left and "1>2" on the right
        builder.write("First TOC entry, MySequence #");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOC.SEQ.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.TOC.SEQ.docx");

        Assert.assertEquals(5, doc.getRange().getFields().getCount());

        fieldToc = (FieldToc) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_TOC, " TOC  \\c MySequence \\s PrefixSequence \\d >",
                "First TOC entry, MySequence #1\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r", fieldToc);
        Assert.assertEquals("MySequence", fieldToc.getTableOfFiguresLabel());
        Assert.assertEquals("PrefixSequence", fieldToc.getPrefixedSequenceIdentifier());
        Assert.assertEquals(">", fieldToc.getSequenceSeparator());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ PrefixSequence _Toc256000000 \\* ARABIC ", "1", fieldSeq);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());

        // Byproduct field created by Aspose.Words
        FieldPageRef fieldPageRef = (FieldPageRef) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF _Toc256000000 \\h ", "2", fieldPageRef);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());
        Assert.assertEquals("_Toc256000000", fieldPageRef.getBookmarkName());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  PrefixSequence", "1", fieldSeq);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "1", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
    }

    @Test
    public void tocSeqNumbering() throws Exception {
        //ExStart
        //ExFor:FieldSeq
        //ExFor:FieldSeq.InsertNextNumber
        //ExFor:FieldSeq.ResetHeadingLevel
        //ExFor:FieldSeq.ResetNumber
        //ExFor:FieldSeq.SequenceIdentifier
        //ExSummary:Shows how to reset numbering of a SEQ field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the current number of the sequence to 100
        builder.write("#");
        FieldSeq fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setResetNumber("100");

        Assert.assertEquals(" SEQ  MySequence \\r 100", fieldSeq.getFieldCode());

        builder.write(", #");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");

        // Insert a heading
        builder.insertBreak(BreakType.PARAGRAPH_BREAK);
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("This level 1 heading will reset MySequence to 1");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));

        // Reset the sequence back to 1 when we encounter a heading of a specified level, which in this case is "1", same as the heading above
        builder.write("\n#");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setResetHeadingLevel("1");

        Assert.assertEquals(" SEQ  MySequence \\s 1", fieldSeq.getFieldCode());

        // Move to the next number
        builder.write(", #");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setInsertNextNumber(true);

        Assert.assertEquals(" SEQ  MySequence \\n", fieldSeq.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SEQ.ResetNumbering.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SEQ.ResetNumbering.docx");

        Assert.assertEquals(4, doc.getRange().getFields().getCount());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\r 100", "100", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "101", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\s 1", "1", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\n", "2", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
    }

    @Test(enabled = false, description = "WORDSNET-18083")
    public void tocSeqBookmark() throws Exception {
        //ExStart
        //ExFor:FieldSeq
        //ExFor:FieldSeq.BookmarkName
        //ExSummary:Shows how to combine table of contents and sequence fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // This TOC takes in all SEQ fields with "MySequence" inside "TOCBookmark"
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);
        fieldToc.setTableOfFiguresLabel("MySequence");
        fieldToc.setBookmarkName("TOCBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);

        Assert.assertEquals(" TOC  \\c MySequence \\b TOCBookmark", fieldToc.getFieldCode());

        builder.write("MySequence #");
        FieldSeq fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        builder.writeln(", won't show up in the TOC because it is outside of the bookmark.");

        builder.startBookmark("TOCBookmark");

        builder.write("MySequence #");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        builder.writeln(", will show up in the TOC next to the entry for the above caption.");

        builder.write("MySequence #");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("OtherSequence");
        builder.writeln(", won't show up in the TOC because it's from a different sequence identifier.");

        // The contents of the bookmark we reference here will not appear at the SEQ field, but will appear in the corresponding TOC entry
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setBookmarkName("SEQBookmark");
        Assert.assertEquals(" SEQ  MySequence SEQBookmark", fieldSeq.getFieldCode());

        // Add bookmark to reference
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark("SEQBookmark");
        builder.write("MySequence #");
        fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        builder.writeln(", text from inside SEQBookmark.");
        builder.endBookmark("SEQBookmark");

        builder.endBookmark("TOCBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SEQ.Bookmark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SEQ.Bookmark.docx");

        Assert.assertEquals(8, doc.getRange().getFields().getCount());

        fieldToc = (FieldToc) doc.getRange().getFields().get(0);
        String[] pageRefIds = (String[]) Arrays.stream(fieldToc.getResult().split(" ")).filter(s -> s.endsWith("_Toc")).toArray();

        Assert.assertEquals(FieldType.FIELD_TOC, fieldToc.getType());
        Assert.assertEquals("MySequence", fieldToc.getTableOfFiguresLabel());
        TestUtil.verifyField(FieldType.FIELD_TOC, " TOC  \\c MySequence \\b TOCBookmark",
                "MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF {pageRefIds[0]} \\h \u00142\u0015\r" +
                        "3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF {pageRefIds[1]} \\h \u00142\u0015\r", fieldToc);

        FieldPageRef fieldPageRef = (FieldPageRef) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF {pageRefIds[0]} \\h ", "2", fieldPageRef);
        Assert.assertEquals(pageRefIds[0], fieldPageRef.getBookmarkName());

        fieldPageRef = (FieldPageRef) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF {pageRefIds[1]} \\h ", "2", fieldPageRef);
        Assert.assertEquals(pageRefIds[1], fieldPageRef.getBookmarkName());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "1", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "2", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  OtherSequence", "1", fieldSeq);
        Assert.assertEquals("OtherSequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(6);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence SEQBookmark", "3", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
        Assert.assertEquals("SEQBookmark", fieldSeq.getBookmarkName());

        fieldSeq = (FieldSeq) doc.getRange().getFields().get(7);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "3", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
    }

    @Test(enabled = false, description = "WORDSNET-13854")
    public void fieldCitation() throws Exception {
        //ExStart
        //ExFor:FieldCitation
        //ExFor:FieldCitation.AnotherSourceTag
        //ExFor:FieldCitation.FormatLanguageId
        //ExFor:FieldCitation.PageNumber
        //ExFor:FieldCitation.Prefix
        //ExFor:FieldCitation.SourceTag
        //ExFor:FieldCitation.Suffix
        //ExFor:FieldCitation.SuppressAuthor
        //ExFor:FieldCitation.SuppressTitle
        //ExFor:FieldCitation.SuppressYear
        //ExFor:FieldCitation.VolumeNumber
        //ExFor:FieldBibliography
        //ExFor:FieldBibliography.FormatLanguageId
        //ExSummary:Shows how to work with CITATION and BIBLIOGRAPHY fields.
        // Open a document that has bibliographical sources
        Document doc = new Document(getMyDir() + "Bibliography.docx");
        Assert.assertEquals(2, doc.getRange().getFields().getCount()); //ExSkip

        // Add text that we can cite
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Text to be cited with one source.");

        // Create a citation field using the document builder
        FieldCitation fieldCitation = (FieldCitation) builder.insertField(FieldType.FIELD_CITATION, true);

        // A simple citation can have just the page number and author's name
        fieldCitation.setSourceTag("Book1"); // We refer to sources using their tag names
        fieldCitation.setPageNumber("85");
        fieldCitation.setSuppressAuthor(false);
        fieldCitation.setSuppressTitle(true);
        fieldCitation.setSuppressYear(true);

        Assert.assertEquals(" CITATION  Book1 \\p 85 \\t \\y", fieldCitation.getFieldCode());

        // We can make a more detailed citation and make it cite 2 sources
        builder.insertParagraph();
        builder.write("Text to be cited with two sources.");
        fieldCitation = (FieldCitation) builder.insertField(FieldType.FIELD_CITATION, true);
        fieldCitation.setSourceTag("Book1");
        fieldCitation.setAnotherSourceTag("Book2");
        fieldCitation.setFormatLanguageId("en-US");
        fieldCitation.setPageNumber("19");
        fieldCitation.setPrefix("Prefix ");
        fieldCitation.setSuffix(" Suffix");
        fieldCitation.setSuppressAuthor(false);
        fieldCitation.setSuppressTitle(false);
        fieldCitation.setSuppressYear(false);
        fieldCitation.setVolumeNumber("VII");

        Assert.assertEquals(" CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII", fieldCitation.getFieldCode());

        // Insert a new page which will contain our bibliography
        builder.insertBreak(BreakType.PAGE_BREAK);

        // All our sources can be displayed using a BIBLIOGRAPHY field
        FieldBibliography fieldBibliography = (FieldBibliography) builder.insertField(FieldType.FIELD_BIBLIOGRAPHY, true);
        fieldBibliography.setFormatLanguageId("1124");

        Assert.assertEquals(" BIBLIOGRAPHY  \\l 1124", fieldBibliography.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CITATION.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.CITATION.docx");

        Assert.assertEquals(5, doc.getRange().getFields().getCount());

        fieldCitation = (FieldCitation) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_CITATION, " CITATION  Book1 \\p 85 \\t \\y", " (Doe, p. 85)", fieldCitation);
        Assert.assertEquals("Book1", fieldCitation.getSourceTag());
        Assert.assertEquals("85", fieldCitation.getPageNumber());
        Assert.assertFalse(fieldCitation.getSuppressAuthor());
        Assert.assertTrue(fieldCitation.getSuppressTitle());
        Assert.assertTrue(fieldCitation.getSuppressYear());

        fieldCitation = (FieldCitation) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_CITATION,
                " CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII",
                " (Doe, 2018; Prefix Cardholder, 2018, VII:19 Suffix)", fieldCitation);
        Assert.assertEquals("Book1", fieldCitation.getSourceTag());
        Assert.assertEquals("Book2", fieldCitation.getAnotherSourceTag());
        Assert.assertEquals("en-US", fieldCitation.getFormatLanguageId());
        Assert.assertEquals("Prefix ", fieldCitation.getPrefix());
        Assert.assertEquals(" Suffix", fieldCitation.getSuffix());
        Assert.assertEquals("19", fieldCitation.getPageNumber());
        Assert.assertFalse(fieldCitation.getSuppressAuthor());
        Assert.assertFalse(fieldCitation.getSuppressTitle());
        Assert.assertFalse(fieldCitation.getSuppressYear());
        Assert.assertEquals("VII", fieldCitation.getVolumeNumber());

        fieldBibliography = (FieldBibliography) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_BIBLIOGRAPHY, " BIBLIOGRAPHY  \\l 1124",
                "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);
        Assert.assertEquals("1124", fieldBibliography.getFormatLanguageId());

        fieldCitation = (FieldCitation) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_CITATION, " CITATION Book1 \\l 1033 ", "(Doe, 2018)", fieldCitation);
        Assert.assertEquals("Book1", fieldCitation.getSourceTag());
        Assert.assertEquals("1033", fieldCitation.getFormatLanguageId());

        fieldBibliography = (FieldBibliography) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_BIBLIOGRAPHY, " BIBLIOGRAPHY ",
                "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);
    }

    @Test
    public void fieldData() throws Exception {
        //ExStart
        //ExFor:FieldData
        //ExSummary:Shows how to insert a data field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a data field
        FieldData field = (FieldData) builder.insertField(FieldType.FIELD_DATA, true);
        Assert.assertEquals(field.getFieldCode(), " DATA ");
        //ExEnd

        TestUtil.verifyField(FieldType.FIELD_DATA, " DATA ", "", DocumentHelper.saveOpen(doc).getRange().getFields().get(0));
    }

    @Test
    public void fieldInclude() throws Exception {
        //ExStart
        //ExFor:FieldInclude
        //ExFor:FieldInclude.BookmarkName
        //ExFor:FieldInclude.LockFields
        //ExFor:FieldInclude.SourceFullName
        //ExFor:FieldInclude.TextConverter
        //ExSummary:Shows how to create an INCLUDE field and set its properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an INCLUDE field with document builder and import a portion of the document defined by a bookmark
        FieldInclude field = (FieldInclude) builder.insertField(FieldType.FIELD_INCLUDE, true);
        field.setSourceFullName(getMyDir() + "Bookmarks.docx");
        field.setBookmarkName("MyBookmark1");
        field.setLockFields(false);
        field.setTextConverter("Microsoft Word");

        Assert.assertTrue(field.getFieldCode().matches(" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\""));

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INCLUDE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INCLUDE.docx");
        field = (FieldInclude) doc.getRange().getFields().get(0);

        Assert.assertEquals(FieldType.FIELD_INCLUDE, field.getType());
        Assert.assertEquals("First bookmark.", field.getResult());
        Assert.assertTrue(field.getFieldCode().matches(" INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\""));

        Assert.assertEquals(getMyDir() + "Bookmarks.docx", field.getSourceFullName());
        Assert.assertEquals("MyBookmark1", field.getBookmarkName());
        Assert.assertFalse(field.getLockFields());
        Assert.assertEquals("Microsoft Word", field.getTextConverter());
    }

    @Test
    public void fieldIncludePicture() throws Exception {
        //ExStart
        //ExFor:FieldIncludePicture
        //ExFor:FieldIncludePicture.GraphicFilter
        //ExFor:FieldIncludePicture.IsLinked
        //ExFor:FieldIncludePicture.ResizeHorizontally
        //ExFor:FieldIncludePicture.ResizeVertically
        //ExFor:FieldIncludePicture.SourceFullName
        //ExFor:FieldImport
        //ExFor:FieldImport.GraphicFilter
        //ExFor:FieldImport.IsLinked
        //ExFor:FieldImport.SourceFullName
        //ExSummary:Shows how to insert images using IMPORT and INCLUDEPICTURE fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldIncludePicture fieldIncludePicture = (FieldIncludePicture) builder.insertField(FieldType.FIELD_INCLUDE_PICTURE, true);
        fieldIncludePicture.setSourceFullName(getImageDir() + "Transparent background logo.png");

        Assert.assertTrue(fieldIncludePicture.getFieldCode().matches(" INCLUDEPICTURE  .*"));

        // Here we apply the PNG32.FLT filter
        fieldIncludePicture.setGraphicFilter("PNG32");
        fieldIncludePicture.isLinked(true);
        fieldIncludePicture.setResizeHorizontally(true);
        fieldIncludePicture.setResizeVertically(true);

        // We can do the same thing with an IMPORT field
        FieldImport fieldImport = (FieldImport) builder.insertField(FieldType.FIELD_IMPORT, true);
        fieldImport.setSourceFullName(getImageDir() + "Transparent background logo.png");
        fieldImport.setGraphicFilter("PNG32");
        fieldImport.isLinked(true);

        Assert.assertTrue(fieldImport.getFieldCode().matches(" IMPORT  .* \\\\c PNG32 \\\\d"));

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INCLUDEPICTURE.docx");
        //ExEnd

        Assert.assertEquals(getImageDir() + "Transparent background logo.png", fieldIncludePicture.getSourceFullName());
        Assert.assertEquals("PNG32", fieldIncludePicture.getGraphicFilter());
        Assert.assertTrue(fieldIncludePicture.isLinked());
        Assert.assertTrue(fieldIncludePicture.getResizeHorizontally());
        Assert.assertTrue(fieldIncludePicture.getResizeVertically());

        Assert.assertEquals(getImageDir() + "Transparent background logo.png", fieldImport.getSourceFullName());
        Assert.assertEquals("PNG32", fieldImport.getGraphicFilter());
        Assert.assertTrue(fieldImport.isLinked());

        doc = new Document(getArtifactsDir() + "Field.INCLUDEPICTURE.docx");

        // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading
        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape image = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertTrue(image.isImage());
        Assert.assertNull(image.getImageData().getImageBytes());
        Assert.assertEquals(getImageDir() + "Transparent background logo.png", image.getImageData().getSourceFullName());

        image = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertTrue(image.isImage());
        Assert.assertNull(image.getImageData().getImageBytes());
        Assert.assertEquals(getImageDir() + "Transparent background logo.png", image.getImageData().getSourceFullName());
    }

    @Test(enabled = false, description = "WORDSNET-17545")
    public void fieldHyperlink() throws Exception {
        //ExStart
        //ExFor:FieldHyperlink
        //ExFor:FieldHyperlink.Address
        //ExFor:FieldHyperlink.IsImageMap
        //ExFor:FieldHyperlink.OpenInNewWindow
        //ExFor:FieldHyperlink.ScreenTip
        //ExFor:FieldHyperlink.SubAddress
        //ExFor:FieldHyperlink.Target
        //ExSummary:Shows how to insert HYPERLINK fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a hyperlink with a document builder
        FieldHyperlink field = (FieldHyperlink) builder.insertField(FieldType.FIELD_HYPERLINK, true);

        // When link is clicked, open a document and place the cursor on the bookmarked location
        field.setAddress(getMyDir() + "Bookmarks.docx");
        field.setSubAddress("MyBookmark3");
        field.setScreenTip("Open " + field.getAddress() + " on bookmark " + field.getSubAddress() + " in a new window");

        builder.writeln();

        // Open html file at a specific frame
        field = (FieldHyperlink) builder.insertField(FieldType.FIELD_HYPERLINK, true);
        field.setAddress(getMyDir() + "Iframes.html");
        field.setScreenTip("Open " + field.getAddress());
        field.setTarget("iframe_3");
        field.setOpenInNewWindow(true);
        field.isImageMap(false);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.HYPERLINK.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.HYPERLINK.docx");
        field = (FieldHyperlink) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_HYPERLINK,
                " HYPERLINK \"" + getMyDir().replace("\\", "\\\\") + "Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + getMyDir() + "Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
                getMyDir() + "Bookmarks.docx - MyBookmark3", field);
        Assert.assertEquals(getMyDir() + "Bookmarks.docx", field.getAddress());
        Assert.assertEquals("MyBookmark3", field.getSubAddress());
        Assert.assertEquals("Open " + field.getAddress().replace("\\", "") + " on bookmark " + field.getSubAddress() + " in a new window", field.getScreenTip());

        field = (FieldHyperlink) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_HYPERLINK, " HYPERLINK \"file:///" + getMyDir().replace("\\", "\\\\").replace(" ", "%20") + "Iframes.html\" \\t \"iframe_3\" \\o \"Open " + getMyDir().replace("\\", "\\\\") + "Iframes.html\" ",
                getMyDir() + "Iframes.html", field);
        Assert.assertEquals("file:///" + getMyDir().replace(" ", "%20") + "Iframes.html", field.getAddress());
        Assert.assertEquals("Open " + getMyDir() + "Iframes.html", field.getScreenTip());
        Assert.assertEquals("iframe_3", field.getTarget());
        Assert.assertFalse(field.getOpenInNewWindow());
        Assert.assertFalse(field.isImageMap());
    }

    //ExStart
    //ExFor:MergeFieldImageDimension
    //ExFor:MergeFieldImageDimension.#ctor
    //ExFor:MergeFieldImageDimension.#ctor(Double)
    //ExFor:MergeFieldImageDimension.#ctor(Double,MergeFieldImageDimensionUnit)
    //ExFor:MergeFieldImageDimension.Unit
    //ExFor:MergeFieldImageDimension.Value
    //ExFor:MergeFieldImageDimensionUnit
    //ExFor:ImageFieldMergingArgs
    //ExFor:ImageFieldMergingArgs.ImageFileName
    //ExFor:ImageFieldMergingArgs.ImageWidth
    //ExFor:ImageFieldMergingArgs.ImageHeight
    //ExSummary:Shows how to set the dimensions of merged images.
    @Test //ExSkip
    public void mergeFieldImageDimension() throws Exception {
        Document doc = new Document();

        // Insert a merge field where images will be placed during the mail merge
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldMergeField field = (FieldMergeField) builder.insertField("MERGEFIELD Image:ImageColumn");

        Assert.assertEquals("Image:ImageColumn", field.getFieldName());

        // Create a data table for the mail merge
        // The name of the column that contains our image filenames needs to match the name of our merge field
        DataTable dataTable = new DataTable("Images");
        dataTable.getColumns().add(new DataColumn("ImageColumn"));
        dataTable.getRows().add(getImageDir() + "Logo.jpg");
        dataTable.getRows().add(getImageDir() + "Transparent background logo.png");
        dataTable.getRows().add(getImageDir() + "Enhanced Windows MetaFile.emf");

        doc.getMailMerge().setFieldMergingCallback(new MergedImageResizer(200.0, 200.0, MergeFieldImageDimensionUnit.POINT));
        doc.getMailMerge().execute(dataTable);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.MERGEFIELD.ImageDimension.docx");
        testMergeFieldImageDimension(doc); //ExSkip
    }

    /// <summary>
    /// Sets the size of all mail merged images to one defined width and height.
    /// </summary>
    private static class MergedImageResizer implements IFieldMergingCallback {
        public MergedImageResizer(final double imageWidth, final double imageHeight, final int unit) {
            mImageWidth = imageWidth;
            mImageHeight = imageHeight;
            mUnit = unit;
        }

        public void fieldMerging(final FieldMergingArgs args) {
            throw new UnsupportedOperationException();
        }

        public void imageFieldMerging(final ImageFieldMergingArgs args) {
            args.setImageFileName(args.getFieldValue().toString());
            args.setImageWidth(new MergeFieldImageDimension(mImageWidth, mUnit));
            args.setImageHeight(new MergeFieldImageDimension(mImageHeight, mUnit));

            Assert.assertEquals(mImageWidth, args.getImageWidth().getValue());
            Assert.assertEquals(mUnit, args.getImageWidth().getUnit());
            Assert.assertEquals(mImageHeight, args.getImageHeight().getValue());
            Assert.assertEquals(mUnit, args.getImageHeight().getUnit());
        }

        private double mImageWidth;
        private double mImageHeight;
        private int mUnit;
    }
    //ExEnd

    private void testMergeFieldImageDimension(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals(3, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, shape);
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyImageInShape(534, 534, ImageType.EMF, shape);
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());
    }

    private void testMergeFieldImages(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(300.0d, shape.getWidth());
        Assert.assertEquals(300.0d, shape.getHeight());

        shape = (Shape) doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, shape);
        Assert.assertEquals(300.0d, shape.getWidth());
        Assert.assertEquals(300.0d, shape.getHeight());
    }

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldIndexFilter() throws Exception {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.BookmarkName
        //ExFor:FieldIndex.EntryType
        //ExFor:FieldXE
        //ExFor:FieldXE.EntryType
        //ExFor:FieldXE.Text
        //ExSummary:Shows how to omit entries while populating an INDEX field with entries from XE fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Set these attributes so that an XE field shows up in the INDEX field's result
        // only if it is within the bounds of a bookmark named "MainBookmark", and is of type "A"
        index.setBookmarkName("MainBookmark");
        index.setEntryType("A");

        Assert.assertEquals(" INDEX  \\b MainBookmark \\f A", index.getFieldCode());

        // Our index will take up the first page
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Start the bookmark that will contain all eligible XE entries
        builder.startBookmark("MainBookmark");

        // This entry will be picked up by the INDEX field because it is inside the bookmark
        // and its type matches the INDEX field's type
        // Note that even though the type is a string, it is defined by only the first character
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 1");
        indexEntry.setEntryType("A");

        Assert.assertEquals(" XE  \"Index entry 1\" \\f A", indexEntry.getFieldCode());

        // Insert an XE field that will not appear in the INDEX field because it is of the wrong type
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 2");
        indexEntry.setEntryType("B");

        // End the bookmark and insert an XE field afterwards
        // It is of the same type as the INDEX field, but will not appear since it is outside of the bookmark
        // Note that the INDEX field itself does not have to be within its bookmark
        builder.endBookmark("MainBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 3");
        indexEntry.setEntryType("A");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Filtering.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Filtering.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\b MainBookmark \\f A", "Index entry 1, 2\r", index);
        Assert.assertEquals("MainBookmark", index.getBookmarkName());
        Assert.assertEquals("A", index.getEntryType());

        indexEntry = (FieldXE) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 1\" \\f A", "", indexEntry);
        Assert.assertEquals("Index entry 1", indexEntry.getText());
        Assert.assertEquals("A", indexEntry.getEntryType());

        indexEntry = (FieldXE) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 2\" \\f B", "", indexEntry);
        Assert.assertEquals("Index entry 2", indexEntry.getText());
        Assert.assertEquals("B", indexEntry.getEntryType());

        indexEntry = (FieldXE) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 3\" \\f A", "", indexEntry);
        Assert.assertEquals("Index entry 3", indexEntry.getText());
        Assert.assertEquals("A", indexEntry.getEntryType());
    }

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldIndexFormatting() throws Exception {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.Heading
        //ExFor:FieldIndex.NumberOfColumns
        //ExFor:FieldIndex.LanguageId
        //ExFor:FieldIndex.LetterRange
        //ExFor:FieldXE
        //ExFor:FieldXE.IsBold
        //ExFor:FieldXE.IsItalic
        //ExFor:FieldXE.Text
        //ExSummary:Shows how to modify an INDEX field's appearance while populating it with XE field entries.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);
        index.setLanguageId("1033");

        // Setting this attribute's value to "A" will group all the entries by their first letter
        // and place that letter in uppercase above each group
        index.setHeading("A");

        // Set the table created by the INDEX field to span over 2 columns
        index.setNumberOfColumns("2");

        // Set any entries with starting letters outside the "a-c" character range to be omitted
        index.setLetterRange("a-c");

        Assert.assertEquals(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.getFieldCode());

        // These next two XE fields will show up under the "A" heading,
        // with their respective text stylings also applied to their page numbers 
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Apple");
        indexEntry.isItalic(true);

        Assert.assertEquals(" XE  Apple \\i", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Apricot");
        indexEntry.isBold(true);

        Assert.assertEquals(" XE  Apricot \\b", indexEntry.getFieldCode());

        // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Banana");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Cherry");

        // All INDEX field entries are sorted alphabetically, so this entry will show up under "A" with the other two
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Avocado");

        // This entry will be excluded because, starting with the letter "D", it is outside the "a-c" range
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Durian");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Formatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Formatting.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        Assert.assertEquals("1033", index.getLanguageId());
        Assert.assertEquals("A", index.getHeading());
        Assert.assertEquals("2", index.getNumberOfColumns());
        Assert.assertEquals("a-c", index.getLetterRange());
        Assert.assertEquals(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.getFieldCode());
        Assert.assertEquals("\fA\r" +
                "Apple, 2\r" +
                "Apricot, 3\r" +
                "Avocado, 6\r" +
                "B\r" +
                "Banana, 4\r" +
                "C\r" +
                "Cherry, 5\r\f", index.getResult());

        indexEntry = (FieldXE) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Apple \\i", "", indexEntry);
        Assert.assertEquals("Apple", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertTrue(indexEntry.isItalic());

        indexEntry = (FieldXE) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Apricot \\b", "", indexEntry);
        Assert.assertEquals("Apricot", indexEntry.getText());
        Assert.assertTrue(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Banana", "", indexEntry);
        Assert.assertEquals("Banana", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Cherry", "", indexEntry);
        Assert.assertEquals("Cherry", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE) doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Avocado", "", indexEntry);
        Assert.assertEquals("Avocado", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE) doc.getRange().getFields().get(6);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Durian", "", indexEntry);
        Assert.assertEquals("Durian", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());
    }

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldIndexSequence() throws Exception {
        //ExStart
        //ExFor:FieldIndex.HasSequenceName
        //ExFor:FieldIndex.SequenceName
        //ExFor:FieldIndex.SequenceSeparator
        //ExSummary:Shows how to split a document into sections by combining INDEX and SEQ fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Set these two attributes to get the INDEX field's table of contents
        // to place the number that the "MySeq" sequence is at in each XE entry's location before the entry's page number,
        // separated by a custom character
        // Note that PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters
        index.setSequenceName("MySequence");
        index.setPageNumberSeparator("\tMySequence at ");
        index.setSequenceSeparator(" on page ");
        Assert.assertTrue(index.hasSequenceName());

        Assert.assertEquals(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.getFieldCode());

        // Insert a SEQ field which moves the "MySequence" sequence to 1
        // This field is treated as normal document text and will not show up on an INDEX field's table of contents
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldSeq sequenceField = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        sequenceField.setSequenceIdentifier("MySequence");

        Assert.assertEquals(" SEQ  MySequence", sequenceField.getFieldCode());

        // Insert a XE field which will show up in the INDEX field
        // Since "MySequence" is at 1 and this XE field is on page 2, along with with the custom separators we defined above,
        // this field's INDEX entry will say "MySequence at 1 on page 2"
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Cat");

        Assert.assertEquals(" XE  Cat", indexEntry.getFieldCode());

        // Insert a page break and advance "MySequence" by 2
        builder.insertBreak(BreakType.PAGE_BREAK);
        sequenceField = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        sequenceField.setSequenceIdentifier("MySequence");
        sequenceField = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        sequenceField.setSequenceIdentifier("MySequence");

        // Insert a XE field with the same text as the one above, which will thus be appended to the same entry in the INDEX field
        // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Cat");

        // Insert an XE field which makes a new entry with MySequence at 3 on page 4
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Dog");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Sequence.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Sequence.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        Assert.assertEquals("MySequence", index.getSequenceName());
        Assert.assertEquals("\tMySequence at ", index.getPageNumberSeparator());
        Assert.assertEquals(" on page ", index.getSequenceSeparator());
        Assert.assertTrue(index.hasSequenceName());
        Assert.assertEquals(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.getFieldCode());
        Assert.assertEquals("Cat\tMySequence at 1 on page 2, 3 on page 3\r" +
                "Dog\tMySequence at 3 on page 4\r", index.getResult());

        Assert.assertEquals(3, DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_SEQUENCE));
    }

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldIndexPageNumberSeparator() throws Exception {
        //ExStart
        //ExFor:FieldIndex.HasPageNumberSeparator
        //ExFor:FieldIndex.PageNumberSeparator
        //ExFor:FieldIndex.PageNumberListSeparator
        //ExSummary:Shows how to edit the page number separator in an INDEX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display a table with the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Set a page number separator and a page number separator
        // The page number separator will go between the INDEX entry's name and first page a corresponsing XE field appears,
        // while the page number list separator will appear between page numbers if there are multiple in the same INDEX field entry
        index.setPageNumberSeparator(", on page(s) ");
        index.setPageNumberListSeparator(" & ");

        Assert.assertEquals(" INDEX  \\e \", on page(s) \" \\l \" & \"", index.getFieldCode());
        Assert.assertTrue(index.hasPageNumberSeparator());

        // Insert 3 XE entries with the same name on three different pages so they all end up in one INDEX field table entry,
        // where both our separators will be applied, resulting in a value of "First entry, on page(s) 2 & 3 & 4"
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("First entry");

        Assert.assertEquals(" XE  \"First entry\"", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("First entry");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("First entry");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.PageNumberList.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.PageNumberList.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\e \", on page(s) \" \\l \" & \"", "First entry, on page(s) 2 & 3 & 4\r", index);
        Assert.assertEquals(", on page(s) ", index.getPageNumberSeparator());
        Assert.assertEquals(" & ", index.getPageNumberListSeparator());
        Assert.assertTrue(index.hasPageNumberSeparator());
    }

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldIndexPageRangeBookmark() throws Exception {
        //ExStart
        //ExFor:FieldIndex.PageRangeSeparator
        //ExFor:FieldXE.HasPageRangeBookmarkName
        //ExFor:FieldXE.PageRangeBookmarkName
        //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        index.setPageNumberSeparator(", on page(s) ");
        index.setPageRangeSeparator(" to ");

        Assert.assertEquals(" INDEX  \\e \", on page(s) \" \\g \" to \"", index.getFieldCode());

        // Insert an XE field on page 2
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("My entry");

        // If we use this attribute to refer to a bookmark,
        // this XE field's page number will be substituted by the page range that the referenced bookmark spans 
        indexEntry.setPageRangeBookmarkName("MyBookmark");

        Assert.assertEquals(" XE  \"My entry\" \\r MyBookmark", indexEntry.getFieldCode());
        Assert.assertTrue(indexEntry.hasPageRangeBookmarkName());

        // Insert a bookmark that starts on page 3 and ends on page 5
        // Since the XE field references this bookmark,
        // its location page number will show up in the INDEX field's table as "3 to 5" instead of "2",
        // which is its actual page
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark("MyBookmark");
        builder.write("Start of MyBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("End of MyBookmark");
        builder.endBookmark("MyBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.PageRangeBookmark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.PageRangeBookmark.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\e \", on page(s) \" \\g \" to \"", "My entry, on page(s) 3 to 5\r", index);
        Assert.assertEquals(", on page(s) ", index.getPageNumberSeparator());
        Assert.assertEquals(" to ", index.getPageRangeSeparator());

        indexEntry = (FieldXE) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"My entry\" \\r MyBookmark", "", indexEntry);
        Assert.assertEquals("My entry", indexEntry.getText());
        Assert.assertEquals("MyBookmark", indexEntry.getPageRangeBookmarkName());
        Assert.assertTrue(indexEntry.hasPageRangeBookmarkName());
    }

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldIndexCrossReferenceSeparator() throws Exception {
        //ExStart
        //ExFor:FieldIndex.CrossReferenceSeparator
        //ExFor:FieldXE.PageNumberReplacement
        //ExSummary:Shows how to define cross references in an INDEX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Define a custom separator that is applied if an XE field contains a page number replacement
        index.setCrossReferenceSeparator(", see: ");

        Assert.assertEquals(" INDEX  \\k \", see: \"", index.getFieldCode());

        // Insert an XE field on page 2
        // That page number, together with the field's Text attribute, will show up as a table of contents entry in the INDEX field,
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Apple");

        Assert.assertEquals(" XE  Apple", indexEntry.getFieldCode());

        // Insert another XE field on page 3, and set a value for "PageNumberReplacement"
        // In the INDEX field's table, this field will display the value of that attribute after the field's CrossReferenceSeparator instead of the page number
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Banana");
        indexEntry.setPageNumberReplacement("Tropical fruit");

        Assert.assertEquals(" XE  Banana \\t \"Tropical fruit\"", indexEntry.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.CrossReferenceSeparator.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.CrossReferenceSeparator.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " INDEX  \\k \", see: \"",
                "Apple, 2\r" +
                        "Banana, see: Tropical fruit\r", index);
        Assert.assertEquals(", see: ", index.getCrossReferenceSeparator());

        indexEntry = (FieldXE) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Apple", "", indexEntry);
        Assert.assertEquals("Apple", indexEntry.getText());
        Assert.assertNull(indexEntry.getPageNumberReplacement());

        indexEntry = (FieldXE) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Banana \\t \"Tropical fruit\"", "", indexEntry);
        Assert.assertEquals("Banana", indexEntry.getText());
        Assert.assertEquals("Tropical fruit", indexEntry.getPageNumberReplacement());
    }

    @Test(enabled = false, description = "WORDSNET-17524", dataProvider = "fieldIndexSubheadingDataProvider")
    public void fieldIndexSubheading(boolean doRunSubentriesOnTheSameLine) throws Exception {
        //ExStart
        //ExFor:FieldIndex.RunSubentriesOnSameLine
        //ExSummary:Shows how to work with subentries in an INDEX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Normally, every XE field that's a subheading of any level is displayed on a unique line entry
        // in the INDEX field's table of contents
        // We can reduce the length of our INDEX table by putting all subheading entries along with their page locations on one line
        index.setRunSubentriesOnSameLine(doRunSubentriesOnTheSameLine);
        index.setPageNumberSeparator(", see page ");
        index.setHeading("A");

        if (doRunSubentriesOnTheSameLine)
            Assert.assertEquals(" INDEX  \\r \\e \", see page \" \\h A", index.getFieldCode());
        else
            Assert.assertEquals(" INDEX  \\e \", see page \" \\h A", index.getFieldCode());

        // An XE field's "Text" attribute is the same thing as the "Heading" that will appear in the INDEX field's table of contents
        // This attribute can also contain one or multiple subheadings, separated by a colon (:),
        // which will be grouped under their parent headings/subheadings in the INDEX field
        // If index.RunSubentriesOnSameLine is false, "Heading 1" will take up one line as a heading,
        // followed by a two-line indented list of "Subheading 1" and "Subheading 2" with their respective page numbers
        // Otherwise, the two subheadings and their page numbers will be on tha same line as their heading
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Heading 1:Subheading 1");

        Assert.assertEquals(" XE  \"Heading 1:Subheading 1\"", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Heading 1:Subheading 2");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Subheading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Subheading.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        if (doRunSubentriesOnTheSameLine) {
            TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\r \\e \", see page \" \\h A",
                    "H\r" +
                            "Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index);
            Assert.assertTrue(index.getRunSubentriesOnSameLine());
        } else {
            TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\e \", see page \" \\h A",
                    "H\r" +
                            "Heading 1\r" +
                            "Subheading 1, see page 2\r" +
                            "Subheading 2, see page 3\r", index);
            Assert.assertFalse(index.getRunSubentriesOnSameLine());
        }

        indexEntry = (FieldXE) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Heading 1:Subheading 1\"", "", indexEntry);
        Assert.assertEquals("Heading 1:Subheading 1", indexEntry.getText());

        indexEntry = (FieldXE) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Heading 1:Subheading 2\"", "", indexEntry);
        Assert.assertEquals("Heading 1:Subheading 2", indexEntry.getText());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "fieldIndexSubheadingDataProvider")
    public static Object[][] fieldIndexSubheadingDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(enabled = false, description = "WORDSNET-17524", dataProvider = "fieldIndexYomiDataProvider")
    public void fieldIndexYomi(boolean doSortEntriesUsingYomi) throws Exception {
        //ExStart
        //ExFor:FieldIndex.UseYomi
        //ExFor:FieldXE.Yomi
        //ExSummary:Shows how to sort INDEX field entries phonetically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display the page locations of XE fields in the document body
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Set the INDEX table to sort entries phonetically using Hiragana
        index.setUseYomi(doSortEntriesUsingYomi);

        if (doSortEntriesUsingYomi)
            Assert.assertEquals(" INDEX  \\y", index.getFieldCode());
        else
            Assert.assertEquals(" INDEX ", index.getFieldCode());

        // Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents,
        // sorted in lexicographic order on their "Text" attribute
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("愛子");

        // The "Text" attrubute may contain a word's spelling in Kanji, whose pronounciation may be ambiguous,
        // while a "Yomi" version of the word will be spelled exactly how it is pronounced using Hiragana
        // If our INDEX field is set to use Yomi, then we can sort phonetically using the "Yomi" attribute values instead of the "Text" attribute
        indexEntry.setYomi("あ");

        Assert.assertEquals(" XE  愛子 \\y あ", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("明美");
        indexEntry.setYomi("あ");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("恵美");
        indexEntry.setYomi("え");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("愛美");
        indexEntry.setYomi("え");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Yomi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Yomi.docx");
        index = (FieldIndex) doc.getRange().getFields().get(0);

        if (doSortEntriesUsingYomi) {
            Assert.assertTrue(index.getUseYomi());
            Assert.assertEquals(" INDEX  \\y", index.getFieldCode());
            Assert.assertEquals("愛子, 2\r" +
                    "明美, 3\r" +
                    "恵美, 4\r" +
                    "愛美, 5\r", index.getResult());
        } else {
            Assert.assertFalse(index.getUseYomi());
            Assert.assertEquals(" INDEX ", index.getFieldCode());
            Assert.assertEquals("恵美, 4\r" +
                    "愛子, 2\r" +
                    "愛美, 5\r" +
                    "明美, 3\r", index.getResult());
        }

        indexEntry = (FieldXE) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  愛子 \\y あ", "", indexEntry);
        Assert.assertEquals("愛子", indexEntry.getText());
        Assert.assertEquals("あ", indexEntry.getYomi());

        indexEntry = (FieldXE) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  明美 \\y あ", "", indexEntry);
        Assert.assertEquals("明美", indexEntry.getText());
        Assert.assertEquals("あ", indexEntry.getYomi());

        indexEntry = (FieldXE) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  恵美 \\y え", "", indexEntry);
        Assert.assertEquals("恵美", indexEntry.getText());
        Assert.assertEquals("え", indexEntry.getYomi());

        indexEntry = (FieldXE) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  愛美 \\y え", "", indexEntry);
        Assert.assertEquals("愛美", indexEntry.getText());
        Assert.assertEquals("え", indexEntry.getYomi());
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "fieldIndexYomiDataProvider")
    public static Object[][] fieldIndexYomiDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void fieldBarcode() throws Exception {
        //ExStart
        //ExFor:FieldBarcode
        //ExFor:FieldBarcode.FacingIdentificationMark
        //ExFor:FieldBarcode.IsBookmark
        //ExFor:FieldBarcode.IsUSPostalAddress
        //ExFor:FieldBarcode.PostalAddress
        //ExSummary:Shows how to insert a BARCODE field and set its properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a bookmark with a US postal code in it
        builder.startBookmark("BarcodeBookmark");
        builder.writeln("96801");
        builder.endBookmark("BarcodeBookmark");

        builder.writeln();

        // Reference a US postal code directly
        FieldBarcode field = (FieldBarcode) builder.insertField(FieldType.FIELD_BARCODE, true);
        field.setFacingIdentificationMark("C");
        field.setPostalAddress("96801");
        field.isUSPostalAddress(true);

        Assert.assertEquals(" BARCODE  96801 \\f C \\u", field.getFieldCode());

        builder.writeln();

        // Reference a US postal code from a bookmark
        field = (FieldBarcode) builder.insertField(FieldType.FIELD_BARCODE, true);
        field.setPostalAddress("BarcodeBookmark");
        field.isBookmark(true);

        Assert.assertEquals(" BARCODE  BarcodeBookmark \\b", field.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.BARCODE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.BARCODE.docx");

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        field = (FieldBarcode) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_BARCODE, " BARCODE  96801 \\f C \\u", "", field);
        Assert.assertEquals("C", field.getFacingIdentificationMark());
        Assert.assertEquals("96801", field.getPostalAddress());
        Assert.assertTrue(field.isUSPostalAddress());

        field = (FieldBarcode) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_BARCODE, " BARCODE  BarcodeBookmark \\b", "", field);
        Assert.assertEquals("BarcodeBookmark", field.getPostalAddress());
        Assert.assertTrue(field.isBookmark());
    }

    @Test
    public void fieldDisplayBarcode() throws Exception {
        //ExStart
        //ExFor:FieldDisplayBarcode
        //ExFor:FieldDisplayBarcode.AddStartStopChar
        //ExFor:FieldDisplayBarcode.BackgroundColor
        //ExFor:FieldDisplayBarcode.BarcodeType
        //ExFor:FieldDisplayBarcode.BarcodeValue
        //ExFor:FieldDisplayBarcode.CaseCodeStyle
        //ExFor:FieldDisplayBarcode.DisplayText
        //ExFor:FieldDisplayBarcode.ErrorCorrectionLevel
        //ExFor:FieldDisplayBarcode.FixCheckDigit
        //ExFor:FieldDisplayBarcode.ForegroundColor
        //ExFor:FieldDisplayBarcode.PosCodeStyle
        //ExFor:FieldDisplayBarcode.ScalingFactor
        //ExFor:FieldDisplayBarcode.SymbolHeight
        //ExFor:FieldDisplayBarcode.SymbolRotation
        //ExSummary:Shows how to insert a DISPLAYBARCODE field and set its properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldDisplayBarcode field = (FieldDisplayBarcode) builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);

        // Insert a QR code
        field.setBarcodeType("QR");
        field.setBarcodeValue("ABC123");
        field.setBackgroundColor("0xF8BD69");
        field.setForegroundColor("0xB5413B");
        field.setErrorCorrectionLevel("3");
        field.setScalingFactor("250");
        field.setSymbolHeight("1000");
        field.setSymbolRotation("0");

        Assert.assertEquals(field.getFieldCode(), " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0");
        builder.writeln();

        // insert a EAN13 barcode
        field = (FieldDisplayBarcode) builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);
        field.setBarcodeType("EAN13");
        field.setBarcodeValue("501234567890");
        field.setDisplayText(true);
        field.setPosCodeStyle("CASE");
        field.setFixCheckDigit(true);

        Assert.assertEquals(field.getFieldCode(), " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x");
        builder.writeln();

        // insert a CODE39 barcode
        field = (FieldDisplayBarcode) builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);
        field.setBarcodeType("CODE39");
        field.setBarcodeValue("12345ABCDE");
        field.setAddStartStopChar(true);

        Assert.assertEquals(field.getFieldCode(), " DISPLAYBARCODE  12345ABCDE CODE39 \\d");
        builder.writeln();

        // insert a ITF14 barcode
        field = (FieldDisplayBarcode) builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);
        field.setBarcodeType("ITF14");
        field.setBarcodeValue("09312345678907");
        field.setCaseCodeStyle("STD");

        Assert.assertEquals(field.getFieldCode(), " DISPLAYBARCODE  09312345678907 ITF14 \\c STD");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.DISPLAYBARCODE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DISPLAYBARCODE.docx");

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        field = (FieldDisplayBarcode) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", "", field);
        Assert.assertEquals("QR", field.getBarcodeType());
        Assert.assertEquals("ABC123", field.getBarcodeValue());
        Assert.assertEquals("0xF8BD69", field.getBackgroundColor());
        Assert.assertEquals("0xB5413B", field.getForegroundColor());
        Assert.assertEquals("3", field.getErrorCorrectionLevel());
        Assert.assertEquals("250", field.getScalingFactor());
        Assert.assertEquals("1000", field.getSymbolHeight());
        Assert.assertEquals("0", field.getSymbolRotation());

        field = (FieldDisplayBarcode) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", "", field);
        Assert.assertEquals("EAN13", field.getBarcodeType());
        Assert.assertEquals("501234567890", field.getBarcodeValue());
        Assert.assertTrue(field.getDisplayText());
        Assert.assertEquals("CASE", field.getPosCodeStyle());
        Assert.assertTrue(field.getFixCheckDigit());

        field = (FieldDisplayBarcode) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  12345ABCDE CODE39 \\d", "", field);
        Assert.assertEquals("CODE39", field.getBarcodeType());
        Assert.assertEquals("12345ABCDE", field.getBarcodeValue());
        Assert.assertTrue(field.getAddStartStopChar());

        field = (FieldDisplayBarcode) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  09312345678907 ITF14 \\c STD", "", field);
        Assert.assertEquals("ITF14", field.getBarcodeType());
        Assert.assertEquals("09312345678907", field.getBarcodeValue());
        Assert.assertEquals("STD", field.getCaseCodeStyle());
    }


    @Test
    public void fieldMergeBarcode_QR() throws Exception {
        //ExStart
        //ExFor:FieldDisplayBarcode
        //ExFor:FieldMergeBarcode
        //ExFor:FieldMergeBarcode.BackgroundColor
        //ExFor:FieldMergeBarcode.BarcodeType
        //ExFor:FieldMergeBarcode.BarcodeValue
        //ExFor:FieldMergeBarcode.ErrorCorrectionLevel
        //ExFor:FieldMergeBarcode.ForegroundColor
        //ExFor:FieldMergeBarcode.ScalingFactor
        //ExFor:FieldMergeBarcode.SymbolHeight
        //ExFor:FieldMergeBarcode.SymbolRotation
        //ExSummary:Shows how to perform a mail merge on QR barcodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field,
        // which functions similar to a MERGEFIELD by creating a barcode from the merged data source's values
        // This field will convert all rows in a merge data source's "MyQRCode" column into QR barcodes
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("QR");
        field.setBarcodeValue("MyQRCode");

        // Edit its appearance such as colors and scale
        field.setBackgroundColor("0xF8BD69");
        field.setForegroundColor("0xB5413B");
        field.setErrorCorrectionLevel("3");
        field.setScalingFactor("250");
        field.setSymbolHeight("1000");
        field.setSymbolRotation("0");

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyQRCode QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0",
                field.getFieldCode());
        builder.writeln();

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
        // When we execute the mail merge,
        // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyQRCode");
        table.getRows().add(new String[]{"ABC123"});
        table.getRows().add(new String[]{"DEF456"});

        doc.getMailMerge().execute(table);

        // Every row in the "MyQRCode" column has created a DISPLAYBARCODE field, which shows a barcode with the merged value
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
                doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
                doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.QR.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.QR.docx");

        Assert.assertEquals(0, DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_MERGE_BARCODE));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE,
                "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", "", barcode);
        Assert.assertEquals("ABC123", barcode.getBarcodeValue());
        Assert.assertEquals("QR", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE,
                "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", "", barcode);
        Assert.assertEquals("DEF456", barcode.getBarcodeValue());
        Assert.assertEquals("QR", barcode.getBarcodeType());
    }

    @Test
    public void fieldMergeBarcode_EAN13() throws Exception {
        //ExStart
        //ExFor:FieldMergeBarcode
        //ExFor:FieldMergeBarcode.BarcodeType
        //ExFor:FieldMergeBarcode.BarcodeValue
        //ExFor:FieldMergeBarcode.DisplayText
        //ExFor:FieldMergeBarcode.FixCheckDigit
        //ExFor:FieldMergeBarcode.PosCodeStyle
        //ExSummary:Shows how to perform a mail merge on EAN13 barcodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field,
        // which functions similar to a MERGEFIELD by creating a barcode from the merged data source's values
        // This field will convert all rows in a merge data source's "MyEAN13Barcode" column into EAN13 barcodes
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("EAN13");
        field.setBarcodeValue("MyEAN13Barcode");

        // Edit its appearance to display barcode data under the lines
        field.setDisplayText(true);
        field.setPosCodeStyle("CASE");
        field.setFixCheckDigit(true);

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x", field.getFieldCode());
        builder.writeln();

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
        // When we execute the mail merge,
        // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyEAN13Barcode");
        table.getRows().add(new String[]{"501234567890"});
        table.getRows().add(new String[]{"123456789012"});

        doc.getMailMerge().execute(table);

        // Every row in the "MyEAN13Barcode" column has created a DISPLAYBARCODE field,
        // which shows a barcode with the merged value
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x",
                doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x",
                doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.EAN13.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.EAN13.docx");

        Assert.assertEquals(0, DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_MERGE_BARCODE));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x", "", barcode);
        Assert.assertEquals("501234567890", barcode.getBarcodeValue());
        Assert.assertEquals("EAN13", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x", "", barcode);
        Assert.assertEquals("123456789012", barcode.getBarcodeValue());
        Assert.assertEquals("EAN13", barcode.getBarcodeType());
    }

    @Test
    public void fieldMergeBarcode_CODE39() throws Exception {
        //ExStart
        //ExFor:FieldMergeBarcode
        //ExFor:FieldMergeBarcode.AddStartStopChar
        //ExFor:FieldMergeBarcode.BarcodeType
        //ExSummary:Shows how to perform a mail merge on CODE39 barcodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field,
        // which functions similar to a MERGEFIELD by creating a barcode from the merged data source's values
        // This field will convert all rows in a merge data source's "MyCODE39Barcode" column into CODE39 barcodes
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("CODE39");
        field.setBarcodeValue("MyCODE39Barcode");

        // Edit its appearance to display start/stop characters
        field.setAddStartStopChar(true);

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyCODE39Barcode CODE39 \\d", field.getFieldCode());
        builder.writeln();

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
        // When we execute the mail merge,
        // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyCODE39Barcode");
        table.getRows().add(new String[]{"12345ABCDE"});
        table.getRows().add(new String[]{"67890FGHIJ"});

        doc.getMailMerge().execute(table);

        // Every row in the "MyCODE39Barcode" column has created a DISPLAYBARCODE field,
        // which shows a barcode with the merged value
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d",
                doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d",
                doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.CODE39.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.CODE39.docx");

        Assert.assertEquals(0, DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_MERGE_BARCODE));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d", "", barcode);
        Assert.assertEquals("12345ABCDE", barcode.getBarcodeValue());
        Assert.assertEquals("CODE39", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d", "", barcode);
        Assert.assertEquals("67890FGHIJ", barcode.getBarcodeValue());
        Assert.assertEquals("CODE39", barcode.getBarcodeType());
    }

    @Test
    public void fieldMergeBarcode_ITF14() throws Exception {
        //ExStart
        //ExFor:FieldMergeBarcode
        //ExFor:FieldMergeBarcode.BarcodeType
        //ExFor:FieldMergeBarcode.CaseCodeStyle
        //ExSummary:Shows how to perform a mail merge on ITF14 barcodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field,
        // which functions similar to a MERGEFIELD by creating a barcode from the merged data source's values
        // This field will convert all rows in a merge data source's "MyITF14Barcode" column into ITF14 barcodes
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("ITF14");
        field.setBarcodeValue("MyITF14Barcode");
        field.setCaseCodeStyle("STD");

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyITF14Barcode ITF14 \\c STD", field.getFieldCode());

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue
        // When we execute the mail merge,
        // a barcode of a type we specified in the MERGEBARCODE field will be created with each row's value
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyITF14Barcode");
        table.getRows().add(new String[]{"09312345678907"});
        table.getRows().add(new String[]{"1234567891234"});

        doc.getMailMerge().execute(table);

        // Every row in the "MyITF14Barcode" column has created a DISPLAYBARCODE field,
        // which shows a barcode with the merged value
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD",
                doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD",
                doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.ITF14.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.ITF14.docx");

        Assert.assertEquals(0, DocumentHelper.getFieldsCount(doc.getRange().getFields(), FieldType.FIELD_MERGE_BARCODE));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD", "", barcode);
        Assert.assertEquals("09312345678907", barcode.getBarcodeValue());
        Assert.assertEquals("ITF14", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD", "", barcode);
        Assert.assertEquals("1234567891234", barcode.getBarcodeValue());
        Assert.assertEquals("ITF14", barcode.getBarcodeType());
    }

    //ExStart
    //ExFor:FieldLink
    //ExFor:FieldLink.AutoUpdate
    //ExFor:FieldLink.FormatUpdateType
    //ExFor:FieldLink.InsertAsBitmap
    //ExFor:FieldLink.InsertAsHtml
    //ExFor:FieldLink.InsertAsPicture
    //ExFor:FieldLink.InsertAsRtf
    //ExFor:FieldLink.InsertAsText
    //ExFor:FieldLink.InsertAsUnicode
    //ExFor:FieldLink.IsLinked
    //ExFor:FieldLink.ProgId
    //ExFor:FieldLink.SourceFullName
    //ExFor:FieldLink.SourceItem
    //ExFor:FieldDde
    //ExFor:FieldDde.AutoUpdate
    //ExFor:FieldDde.InsertAsBitmap
    //ExFor:FieldDde.InsertAsHtml
    //ExFor:FieldDde.InsertAsPicture
    //ExFor:FieldDde.InsertAsRtf
    //ExFor:FieldDde.InsertAsText
    //ExFor:FieldDde.InsertAsUnicode
    //ExFor:FieldDde.IsLinked
    //ExFor:FieldDde.ProgId
    //ExFor:FieldDde.SourceFullName
    //ExFor:FieldDde.SourceItem
    //ExFor:FieldDdeAuto
    //ExFor:FieldDdeAuto.InsertAsBitmap
    //ExFor:FieldDdeAuto.InsertAsHtml
    //ExFor:FieldDdeAuto.InsertAsPicture
    //ExFor:FieldDdeAuto.InsertAsRtf
    //ExFor:FieldDdeAuto.InsertAsText
    //ExFor:FieldDdeAuto.InsertAsUnicode
    //ExFor:FieldDdeAuto.IsLinked
    //ExFor:FieldDdeAuto.ProgId
    //ExFor:FieldDdeAuto.SourceFullName
    //ExFor:FieldDdeAuto.SourceItem
    //ExSummary:Shows how to insert linked objects as LINK, DDE and DDEAUTO fields and present them within the document in different ways.
    @Test(enabled = false, description = "WORDSNET-16226", dataProvider = "fieldLinkedObjectsAsTextDataProvider")
    //ExSkip
    public void fieldLinkedObjectsAsText(/*InsertLinkedObjectAs*/int insertLinkedObjectAs) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert fields containing text from another document and present them as text (see InsertLinkedObjectAs enum)
        builder.writeln("FieldLink:\n");
        insertFieldLink(builder, insertLinkedObjectAs, "Word.Document.8", getMyDir() + "Document.docx", null, true);

        builder.writeln("FieldDde:\n");
        insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true, true);

        builder.writeln("FieldDdeAuto:\n");
        insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LINK.DDE.DDEAUTO.docx");
    }

    @DataProvider(name = "fieldLinkedObjectsAsTextDataProvider")
    public static Object[][] fieldLinkedObjectsAsTextDataProvider() {
        return new Object[][]
                {
                        {InsertLinkedObjectAs.TEXT},
                        {InsertLinkedObjectAs.UNICODE},
                        {InsertLinkedObjectAs.HTML},
                        {InsertLinkedObjectAs.RTF},
                };
    }

    @Test(enabled = false, description = "WORDSNET-16226", dataProvider = "fieldLinkedObjectsAsImageDataProvider")
    //ExSkip
    public void fieldLinkedObjectsAsImage(/*InsertLinkedObjectAs*/int insertLinkedObjectAs) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert one cell from a spreadsheet as an image (see InsertLinkedObjectAs enum)
        builder.writeln("FieldLink:\n");
        insertFieldLink(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "MySpreadsheet.xlsx",
                "Sheet1!R2C2", true);

        builder.writeln("FieldDde:\n");
        insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true, true);

        builder.writeln("FieldDdeAuto:\n");
        insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
                "Sheet1!R1C1", true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LINK.DDE.DDEAUTO.AsImage.docx");
    }

    @DataProvider(name = "fieldLinkedObjectsAsImageDataProvider")
    public static Object[][] fieldLinkedObjectsAsImageDataProvider() {
        return new Object[][]
                {
                        {InsertLinkedObjectAs.PICTURE},
                        {InsertLinkedObjectAs.BITMAP},
                };
    }

    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters.
    /// </summary>
    private void insertFieldLink(final DocumentBuilder builder, final int insertLinkedObjectAs,
                                 final String progId, final String sourceFullName, final String sourceItem,
                                 final boolean shouldAutoUpdate) throws Exception {
        FieldLink field = (FieldLink) builder.insertField(FieldType.FIELD_LINK, true);

        switch (insertLinkedObjectAs) {
            case InsertLinkedObjectAs.TEXT:
                field.setInsertAsText(true);
                break;
            case InsertLinkedObjectAs.UNICODE:
                field.setInsertAsUnicode(true);
                break;
            case InsertLinkedObjectAs.HTML:
                field.setInsertAsHtml(true);
                break;
            case InsertLinkedObjectAs.RTF:
                field.setInsertAsRtf(true);
                break;
            case InsertLinkedObjectAs.PICTURE:
                field.setInsertAsPicture(true);
                break;
            case InsertLinkedObjectAs.BITMAP:
                field.setInsertAsBitmap(true);
                break;
        }

        field.setAutoUpdate(shouldAutoUpdate);
        field.setProgId(progId);
        field.setSourceFullName(sourceFullName);
        field.setSourceItem(sourceItem);

        builder.writeln("\n");
    }

    /// <summary>
    /// Use a document builder to insert a DDE field and set its properties according to parameters.
    /// </summary>
    private void insertFieldDde(final DocumentBuilder builder, final int insertLinkedObjectAs, final String progId,
                                final String sourceFullName, final String sourceItem, final boolean isLinked,
                                final boolean shouldAutoUpdate) throws Exception {
        FieldDde field = (FieldDde) builder.insertField(FieldType.FIELD_DDE, true);

        switch (insertLinkedObjectAs) {
            case InsertLinkedObjectAs.TEXT:
                field.setInsertAsText(true);
                break;
            case InsertLinkedObjectAs.UNICODE:
                field.setInsertAsUnicode(true);
                break;
            case InsertLinkedObjectAs.HTML:
                field.setInsertAsHtml(true);
                break;
            case InsertLinkedObjectAs.RTF:
                field.setInsertAsRtf(true);
                break;
            case InsertLinkedObjectAs.PICTURE:
                field.setInsertAsPicture(true);
                break;
            case InsertLinkedObjectAs.BITMAP:
                field.setInsertAsBitmap(true);
                break;
        }

        field.setAutoUpdate(shouldAutoUpdate);
        field.setProgId(progId);
        field.setSourceFullName(sourceFullName);
        field.setSourceItem(sourceItem);
        field.isLinked(isLinked);

        builder.writeln("\n");
    }

    /// <summary>
    /// Use a document builder to insert a DDEAUTO field and set its properties according to parameters.
    /// </summary>
    private void insertFieldDdeAuto(final DocumentBuilder builder, final int insertLinkedObjectAs,
                                    final String progId, final String sourceFullName, final String sourceItem,
                                    final boolean isLinked) throws Exception {
        FieldDdeAuto field = (FieldDdeAuto) builder.insertField(FieldType.FIELD_DDE_AUTO, true);

        switch (insertLinkedObjectAs) {
            case InsertLinkedObjectAs.TEXT:
                field.setInsertAsText(true);
                break;
            case InsertLinkedObjectAs.UNICODE:
                field.setInsertAsUnicode(true);
                break;
            case InsertLinkedObjectAs.HTML:
                field.setInsertAsHtml(true);
                break;
            case InsertLinkedObjectAs.RTF:
                field.setInsertAsRtf(true);
                break;
            case InsertLinkedObjectAs.PICTURE:
                field.setInsertAsPicture(true);
                break;
            case InsertLinkedObjectAs.BITMAP:
                field.setInsertAsBitmap(true);
                break;
        }

        field.setProgId(progId);
        field.setSourceFullName(sourceFullName);
        field.setSourceItem(sourceItem);
        field.isLinked(isLinked);
    }

    public final class InsertLinkedObjectAs {
        private InsertLinkedObjectAs() {
        }

        // LinkedObjectAsText
        public static final int TEXT = 0;
        public static final int UNICODE = 1;
        public static final int HTML = 2;
        public static final int RTF = 3;
        // LinkedObjectAsImage
        public static final int PICTURE = 4;
        public static final int BITMAP = 5;
    }
    //ExEnd

    @Test
    public void fieldUserAddress() throws Exception {
        //ExStart
        //ExFor:FieldUserAddress
        //ExFor:FieldUserAddress.UserAddress
        //ExSummary:Shows how to use the USERADDRESS field.
        Document doc = new Document();

        // Create a user information object and set it as the data source for our field
        UserInformation userInformation = new UserInformation();
        userInformation.setAddress("123 Main Street");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Display the current user's address with a USERADDRESS field
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldUserAddress fieldUserAddress = (FieldUserAddress) builder.insertField(FieldType.FIELD_USER_ADDRESS, true);
        Assert.assertEquals(userInformation.getAddress(), fieldUserAddress.getResult());

        Assert.assertEquals(" USERADDRESS ", fieldUserAddress.getFieldCode());
        Assert.assertEquals("123 Main Street", fieldUserAddress.getResult());

        // We can set this attribute to get our field to display a different value
        fieldUserAddress.setUserAddress("456 North Road");
        fieldUserAddress.update();

        Assert.assertEquals(" USERADDRESS  \"456 North Road\"", fieldUserAddress.getFieldCode());
        Assert.assertEquals("456 North Road", fieldUserAddress.getResult());

        // This does not change the value in the user information object
        Assert.assertEquals("123 Main Street", doc.getFieldOptions().getCurrentUser().getAddress());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USERADDRESS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.USERADDRESS.docx");

        fieldUserAddress = (FieldUserAddress) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_USER_ADDRESS, " USERADDRESS  \"456 North Road\"", "456 North Road", fieldUserAddress);
        Assert.assertEquals("456 North Road", fieldUserAddress.getUserAddress());
    }

    @Test
    public void fieldUserInitials() throws Exception {
        //ExStart
        //ExFor:FieldUserInitials
        //ExFor:FieldUserInitials.UserInitials
        //ExSummary:Shows how to use the USERINITIALS field.
        Document doc = new Document();

        // Create a user information object and set it as the data source for our field
        UserInformation userInformation = new UserInformation();
        userInformation.setInitials("J. D.");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Display the current user's Initials with a USERINITIALS field
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldUserInitials fieldUserInitials = (FieldUserInitials) builder.insertField(FieldType.FIELD_USER_INITIALS, true);
        Assert.assertEquals(userInformation.getInitials(), fieldUserInitials.getResult());

        Assert.assertEquals(" USERINITIALS ", fieldUserInitials.getFieldCode());
        Assert.assertEquals("J. D.", fieldUserInitials.getResult());

        // We can set this attribute to get our field to display a different value
        fieldUserInitials.setUserInitials("J. C.");
        fieldUserInitials.update();

        Assert.assertEquals(" USERINITIALS  \"J. C.\"", fieldUserInitials.getFieldCode());
        Assert.assertEquals("J. C.", fieldUserInitials.getResult());

        // This does not change the value in the user information object
        Assert.assertEquals("J. D.", doc.getFieldOptions().getCurrentUser().getInitials());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USERINITIALS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.USERINITIALS.docx");

        fieldUserInitials = (FieldUserInitials) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_USER_INITIALS, " USERINITIALS  \"J. C.\"", "J. C.", fieldUserInitials);
        Assert.assertEquals("J. C.", fieldUserInitials.getUserInitials());
    }

    @Test
    public void fieldUserName() throws Exception {
        //ExStart
        //ExFor:FieldUserName
        //ExFor:FieldUserName.UserName
        //ExSummary:Shows how to use the USERNAME field.
        Document doc = new Document();

        // Create a user information object and set it as the data source for our field
        UserInformation userInformation = new UserInformation();
        userInformation.setName("John Doe");
        doc.getFieldOptions().setCurrentUser(userInformation);

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Display the current user's Name with a USERNAME field
        FieldUserName fieldUserName = (FieldUserName) builder.insertField(FieldType.FIELD_USER_NAME, true);
        Assert.assertEquals(userInformation.getName(), fieldUserName.getResult());

        Assert.assertEquals(" USERNAME ", fieldUserName.getFieldCode());
        Assert.assertEquals("John Doe", fieldUserName.getResult());

        // We can set this attribute to get our field to display a different value
        fieldUserName.setUserName("Jane Doe");
        fieldUserName.update();

        Assert.assertEquals(" USERNAME  \"Jane Doe\"", fieldUserName.getFieldCode());
        Assert.assertEquals("Jane Doe", fieldUserName.getResult());

        // This does not change the value in the user information object
        Assert.assertEquals("John Doe", doc.getFieldOptions().getCurrentUser().getName());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USERNAME.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.USERNAME.docx");

        fieldUserName = (FieldUserName) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_USER_NAME, " USERNAME  \"Jane Doe\"", "Jane Doe", fieldUserName);
        Assert.assertEquals("Jane Doe", fieldUserName.getUserName());
    }

    @Test(enabled = false, description = "WORDSNET-17657")
    public void fieldStyleRefParagraphNumbers() throws Exception {
        //ExStart
        //ExFor:FieldStyleRef
        //ExFor:FieldStyleRef.InsertParagraphNumber
        //ExFor:FieldStyleRef.InsertParagraphNumberInFullContext
        //ExFor:FieldStyleRef.InsertParagraphNumberInRelativeContext
        //ExFor:FieldStyleRef.InsertRelativePosition
        //ExFor:FieldStyleRef.SearchFromBottom
        //ExFor:FieldStyleRef.StyleName
        //ExFor:FieldStyleRef.SuppressNonDelimiters
        //ExSummary:Shows how to use STYLEREF fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on one of the Microsoft Word list templates
        List list = doc.getLists().add(com.aspose.words.ListTemplate.NUMBER_DEFAULT);

        // This generated list will look like "1.a )"
        // The space before the bracket is a non-delimiter character that can be suppressed
        list.getListLevels().get(0).setNumberFormat("\u0000.");
        list.getListLevels().get(1).setNumberFormat("\u0001 )");

        // Add text and apply paragraph styles that will be referenced by STYLEREF fields
        builder.getListFormat().setList(list);
        builder.getListFormat().listIndent();
        builder.getParagraphFormat().setStyle(doc.getStyles().get("List Paragraph"));
        builder.writeln("Item 1");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Quote"));
        builder.writeln("Item 2");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("List Paragraph"));
        builder.writeln("Item 3");
        builder.getListFormat().removeNumbers();
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));

        // Place a STYLEREF field in the header and have it display the first "List Paragraph"-styled text in the document
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        FieldStyleRef field = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("List Paragraph");

        // Place a STYLEREF field in the footer and have it display the last text
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        field = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("List Paragraph");
        field.setSearchFromBottom(true);

        builder.moveToDocumentEnd();

        // We can also use STYLEREF fields to reference the list numbers of lists
        builder.write("\nParagraph number: ");
        field = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumber(true);

        builder.write("\nParagraph number, relative context: ");
        field = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumberInRelativeContext(true);

        builder.write("\nParagraph number, full context: ");
        field = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumberInFullContext(true);

        builder.write("\nParagraph number, full context, non-delimiter chars suppressed: ");
        field = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumberInFullContext(true);
        field.setSuppressNonDelimiters(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.STYLEREF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.STYLEREF.docx");

        field = (FieldStyleRef) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  \"List Paragraph\"", "Item 1", field);
        Assert.assertEquals("List Paragraph", field.getStyleName());

        field = (FieldStyleRef) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  \"List Paragraph\" \\l", "Item 3", field);
        Assert.assertEquals("List Paragraph", field.getStyleName());
        Assert.assertTrue(field.getSearchFromBottom());

        field = (FieldStyleRef) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\n", "b )", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumber());

        field = (FieldStyleRef) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\r", "b )", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumberInRelativeContext());

        field = (FieldStyleRef) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\w", "1.b )", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumberInFullContext());

        field = (FieldStyleRef) doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\w \\t", "1.b)", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumberInFullContext());
        Assert.assertTrue(field.getSuppressNonDelimiters());
    }

    @Test
    public void fieldDate() throws Exception {
        //ExStart
        //ExFor:FieldDate
        //ExFor:FieldDate.UseLunarCalendar
        //ExFor:FieldDate.UseSakaEraCalendar
        //ExFor:FieldDate.UseUmAlQuraCalendar
        //ExFor:FieldDate.UseLastFormat
        //ExSummary:Shows how to insert DATE fields with different kinds of calendars.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // One way of putting dates into our documents is inserting DATE fields with document builder
        FieldDate field = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);

        // Set the field's date to the current date of the Islamic Lunar Calendar
        field.setUseLunarCalendar(true);
        Assert.assertEquals(" DATE  \\h", field.getFieldCode());
        builder.writeln();

        // Insert a date field with the current date of the Umm al-Qura calendar
        field = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseUmAlQuraCalendar(true);
        Assert.assertEquals(" DATE  \\u", field.getFieldCode());
        builder.writeln();

        // Insert a date field with the current date of the Indian national calendar
        field = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseSakaEraCalendar(true);
        Assert.assertEquals(" DATE  \\s", field.getFieldCode());
        builder.writeln();

        // Insert a date field with the current date of the calendar used in the (Insert > Date and Time) dialog box
        field = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseLastFormat(true);
        Assert.assertEquals(" DATE  \\l", field.getFieldCode());
        builder.writeln();

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.DATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DATE.docx");

        field = (FieldDate) doc.getRange().getFields().get(0);

        Assert.assertEquals(FieldType.FIELD_DATE, field.getType());
        Assert.assertTrue(field.getUseLunarCalendar());
        Assert.assertEquals(" DATE  \\h", field.getFieldCode());
        Assert.assertTrue(doc.getRange().getFields().get(0).getResult().matches("\\d{1,2}[/]\\d{1,2}[/]\\d{4}"));

        field = (FieldDate) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\u", LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")), field);
        Assert.assertTrue(field.getUseUmAlQuraCalendar());

        field = (FieldDate) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\s", LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")), field);
        Assert.assertTrue(field.getUseSakaEraCalendar());

        field = (FieldDate) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\l", LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")), field);
        Assert.assertTrue(field.getUseLastFormat());
    }

    @Test(enabled = false, description = "WORDSNET-17669")
    public void fieldCreateDate() throws Exception {
        //ExStart
        //ExFor:FieldCreateDate
        //ExFor:FieldCreateDate.UseLunarCalendar
        //ExFor:FieldCreateDate.UseSakaEraCalendar
        //ExFor:FieldCreateDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to insert CREATEDATE fields to display document creation dates.
        // Open an existing document and move a document builder to the end
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln(" Date this document was created:");

        // Insert a CREATEDATE field and display, using the Lunar Calendar, the date the document was created
        builder.write("According to the Lunar Calendar - ");
        FieldCreateDate field = (FieldCreateDate) builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        field.setUseLunarCalendar(true);

        Assert.assertEquals(" CREATEDATE  \\h", field.getFieldCode());

        // Display the date using the Umm al-Qura Calendar
        builder.write("\nAccording to the Umm al-Qura Calendar - ");
        field = (FieldCreateDate) builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        field.setUseUmAlQuraCalendar(true);

        Assert.assertEquals(" CREATEDATE  \\u", field.getFieldCode());

        // Display the date using the Indian National Calendar
        builder.write("\nAccording to the Indian National Calendar - ");
        field = (FieldCreateDate) builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        field.setUseSakaEraCalendar(true);

        Assert.assertEquals(" CREATEDATE  \\s", field.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CREATEDATE.docx");
        //ExEnd
    }

    @Test(enabled = false, description = "WORDSNET-17669")
    public void fieldSaveDate() throws Exception {
        //ExStart
        //ExFor:FieldSaveDate
        //ExFor:FieldSaveDate.UseLunarCalendar
        //ExFor:FieldSaveDate.UseSakaEraCalendar
        //ExFor:FieldSaveDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to insert SAVEDATE fields the date and time a document was last saved.
        // Open an existing document and move a document builder to the end
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln(" Date this document was last saved:");

        // Insert a SAVEDATE field and display, using the Lunar Calendar, the date the document was last saved
        builder.write("According to the Lunar Calendar - ");
        FieldSaveDate field = (FieldSaveDate) builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        field.setUseLunarCalendar(true);

        Assert.assertEquals(" SAVEDATE  \\h", field.getFieldCode());

        // Display the date using the Umm al-Qura Calendar
        builder.write("\nAccording to the Umm al-Qura calendar - ");
        field = (FieldSaveDate) builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        field.setUseUmAlQuraCalendar(true);

        Assert.assertEquals(" SAVEDATE  \\u", field.getFieldCode());

        // Display the date using the Indian National Calendar
        builder.write("\nAccording to the Indian National calendar - ");
        field = (FieldSaveDate) builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        field.setUseSakaEraCalendar(true);

        Assert.assertEquals(" SAVEDATE  \\s", field.getFieldCode());

        // While the date/time of the most recent save operation is tracked automatically by Microsoft Word,
        // we will need to update the value manually if we wish to do the same thing when calling the Save() method
        doc.getBuiltInDocumentProperties().setLastSavedTime(new Date());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SAVEDATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SAVEDATE.docx");

        System.out.println(doc.getBuiltInDocumentProperties().getLastSavedTime());

        field = (FieldSaveDate) doc.getRange().getFields().get(0);

        Assert.assertEquals(FieldType.FIELD_SAVE_DATE, field.getType());
        Assert.assertTrue(field.getUseLunarCalendar());
        Assert.assertEquals(" SAVEDATE  \\h", field.getFieldCode());

        Assert.assertTrue(field.getResult().matches("\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M"));

        field = (FieldSaveDate) doc.getRange().getFields().get(1);

        Assert.assertEquals(FieldType.FIELD_SAVE_DATE, field.getType());
        Assert.assertTrue(field.getUseUmAlQuraCalendar());
        Assert.assertEquals(" SAVEDATE  \\u", field.getFieldCode());
        Assert.assertTrue(field.getResult().matches("\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M"));
    }

    @Test
    public void fieldBuilder() throws Exception {
        //ExStart
        //ExFor:FieldBuilder
        //ExFor:FieldBuilder.AddArgument(Int32)
        //ExFor:FieldBuilder.AddArgument(FieldArgumentBuilder)
        //ExFor:FieldBuilder.AddArgument(String)
        //ExFor:FieldBuilder.AddArgument(Double)
        //ExFor:FieldBuilder.AddArgument(FieldBuilder)
        //ExFor:FieldBuilder.AddSwitch(String)
        //ExFor:FieldBuilder.AddSwitch(String, Double)
        //ExFor:FieldBuilder.AddSwitch(String, Int32)
        //ExFor:FieldBuilder.AddSwitch(String, String)
        //ExFor:FieldBuilder.BuildAndInsert(Paragraph)
        //ExFor:FieldArgumentBuilder
        //ExFor:FieldArgumentBuilder.AddField(FieldBuilder)
        //ExFor:FieldArgumentBuilder.AddText(String)
        //ExFor:FieldArgumentBuilder.AddNode(Inline)
        //ExSummary:Shows how to insert fields using a field builder.
        Document doc = new Document();

        // Use a field builder to add a SYMBOL field which displays the "F with hook" symbol
        FieldBuilder builder = new FieldBuilder(FieldType.FIELD_SYMBOL);
        builder.addArgument(402);
        builder.addSwitch("\\f", "Arial");
        builder.addSwitch("\\s", 25);
        builder.addSwitch("\\u");
        Field field = builder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph());

        Assert.assertEquals(field.getFieldCode(), " SYMBOL 402 \\f Arial \\s 25 \\u ");

        // Use a field builder to create a formula field that will be used by another field builder
        FieldBuilder innerFormulaBuilder = new FieldBuilder(FieldType.FIELD_FORMULA);
        innerFormulaBuilder.addArgument(100);
        innerFormulaBuilder.addArgument("+");
        innerFormulaBuilder.addArgument(74);

        // Add a field builder as an argument to another field builder
        // The result of our formula field will be used as an ANSI value representing the "enclosed R" symbol,
        // to be displayed by this SYMBOL field
        builder = new FieldBuilder(FieldType.FIELD_SYMBOL);
        builder.addArgument(innerFormulaBuilder);
        field = builder.buildAndInsert(doc.getFirstSection().getBody().appendParagraph(""));

        Assert.assertEquals(field.getFieldCode(), " SYMBOL \u0013 = 100 + 74 \u0014\u0015 ");

        // Now we will use our builder to construct a more complex field with nested fields
        // For our IF field, we will first create two formula fields to serve as expressions
        // Their results will be tested for equality to decide what value an IF field displays
        FieldBuilder leftExpression = new FieldBuilder(FieldType.FIELD_FORMULA);
        leftExpression.addArgument(2);
        leftExpression.addArgument("+");
        leftExpression.addArgument(3);

        FieldBuilder rightExpression = new FieldBuilder(FieldType.FIELD_FORMULA);
        rightExpression.addArgument(2.5);
        rightExpression.addArgument("*");
        rightExpression.addArgument(5.2);

        // Next, we will create two field arguments using field argument builders
        // These will serve as the two possible outputs of our IF field and they will also use our two expressions
        FieldArgumentBuilder trueOutput = new FieldArgumentBuilder();
        trueOutput.addText("True, both expressions amount to ");
        trueOutput.addField(leftExpression);

        FieldArgumentBuilder falseOutput = new FieldArgumentBuilder();
        falseOutput.addNode(new Run(doc, "False, "));
        falseOutput.addField(leftExpression);
        falseOutput.addNode(new Run(doc, " does not equal "));
        falseOutput.addField(rightExpression);

        // Finally, we will use a field builder to create an IF field which takes two field builders as expressions,
        // and two field argument builders as the two potential outputs
        builder = new FieldBuilder(FieldType.FIELD_IF);
        builder.addArgument(leftExpression);
        builder.addArgument("=");
        builder.addArgument(rightExpression);
        builder.addArgument(trueOutput);
        builder.addArgument(falseOutput);

        builder.buildAndInsert(doc.getFirstSection().getBody().appendParagraph(""));

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SYMBOL.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SYMBOL.docx");

        FieldSymbol fieldSymbol = (FieldSymbol) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL 402 \\f Arial \\s 25 \\u ", "", fieldSymbol);
        Assert.assertEquals("ƒ", fieldSymbol.getDisplayResult());

        fieldSymbol = (FieldSymbol) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", "", fieldSymbol);
        Assert.assertEquals("®", fieldSymbol.getDisplayResult());

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 100 + 74 ", "174", doc.getRange().getFields().get(2));

        TestUtil.verifyField(FieldType.FIELD_IF,
                " IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 " +
                        "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                        "\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
                "False, 5 does not equal 13", doc.getRange().getFields().get(3));

        Document finalDoc = doc;
        Assert.assertThrows(AssertionError.class, () -> TestUtil.fieldsAreNested(finalDoc.getRange().getFields().get(2), finalDoc.getRange().getFields().get(3)));

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2 + 3 ", "5", doc.getRange().getFields().get(4));
        TestUtil.fieldsAreNested(doc.getRange().getFields().get(4), doc.getRange().getFields().get(3));

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2.5 * 5.2 ", "13", doc.getRange().getFields().get(5));
        TestUtil.fieldsAreNested(doc.getRange().getFields().get(5), doc.getRange().getFields().get(3));

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2 + 3 ", "", doc.getRange().getFields().get(6));
        TestUtil.fieldsAreNested(doc.getRange().getFields().get(6), doc.getRange().getFields().get(3));

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2 + 3 ", "5", doc.getRange().getFields().get(7));
        TestUtil.fieldsAreNested(doc.getRange().getFields().get(7), doc.getRange().getFields().get(3));

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2.5 * 5.2 ", "13", doc.getRange().getFields().get(8));
        TestUtil.fieldsAreNested(doc.getRange().getFields().get(8), doc.getRange().getFields().get(3));
    }

    @Test
    public void fieldAuthor() throws Exception {
        //ExStart
        //ExFor:FieldAuthor
        //ExFor:FieldAuthor.AuthorName
        //ExFor:FieldOptions.DefaultDocumentAuthor
        //ExSummary:Shows how to display a document creator's name with an AUTHOR field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we open an existing document, the document's author's full name will be displayed by the field
        // If we create a document programmatically, we need to set this attribute to the author's name so our field has something to display
        doc.getFieldOptions().setDefaultDocumentAuthor("Joe Bloggs");

        builder.write("This document was created by ");
        FieldAuthor field = (FieldAuthor) builder.insertField(FieldType.FIELD_AUTHOR, true);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " AUTHOR ");
        Assert.assertEquals(field.getResult(), "Joe Bloggs");

        // If this property has a value, it will supersede the one we set above
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " AUTHOR ");
        Assert.assertEquals(field.getResult(), "John Doe");

        // Our field can also override the document's built in author name like this
        field.setAuthorName("Jane Doe");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " AUTHOR  \"Jane Doe\"");
        Assert.assertEquals(field.getResult(), "Jane Doe");

        // The author name in the built in properties was changed by the field, but the default document author stays the same
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getAuthor(), "Jane Doe");
        Assert.assertEquals(doc.getFieldOptions().getDefaultDocumentAuthor(), "Joe Bloggs");

        doc.save(getArtifactsDir() + "Field.AUTHOR.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTHOR.docx");

        Assert.assertNull(doc.getFieldOptions().getDefaultDocumentAuthor());
        Assert.assertEquals("Jane Doe", doc.getBuiltInDocumentProperties().getAuthor());

        field = (FieldAuthor) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_AUTHOR, " AUTHOR  \"Jane Doe\"", "Jane Doe", field);
        Assert.assertEquals("Jane Doe", field.getAuthorName());
    }

    @Test
    public void fieldDocVariable() throws Exception {
        //ExStart
        //ExFor:FieldDocProperty
        //ExFor:FieldDocVariable
        //ExFor:FieldDocVariable.VariableName
        //ExSummary:Shows how to use fields to display document properties and variables.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the value of a document property
        doc.getBuiltInDocumentProperties().setCategory("My category");

        // Display the value of that property with a DOCPROPERTY field
        FieldDocProperty fieldDocProperty = (FieldDocProperty) builder.insertField(" DOCPROPERTY Category ");
        fieldDocProperty.update();

        Assert.assertEquals(fieldDocProperty.getFieldCode(), " DOCPROPERTY Category ");
        Assert.assertEquals(fieldDocProperty.getResult(), "My category");

        builder.writeln();

        // While the set of a document's properties is fixed, we can add, name and define our own values in the variables collection
        Assert.assertEquals(doc.getVariables().getCount(), 0);
        doc.getVariables().add("My variable", "My variable's value");

        // We can access a variable using its name and display it with a DOCVARIABLE field
        FieldDocVariable fieldDocVariable = (FieldDocVariable) builder.insertField(FieldType.FIELD_DOC_VARIABLE, true);
        fieldDocVariable.setVariableName("My Variable");
        fieldDocVariable.update();

        Assert.assertEquals(" DOCVARIABLE  \"My Variable\"", fieldDocVariable.getFieldCode());
        Assert.assertEquals("My variable's value", fieldDocVariable.getResult());

        doc.save(getArtifactsDir() + "Field.DOCPROPERTY.DOCVARIABLE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DOCPROPERTY.DOCVARIABLE.docx");

        Assert.assertEquals("My category", doc.getBuiltInDocumentProperties().getCategory());

        fieldDocProperty = (FieldDocProperty) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DOC_PROPERTY, " DOCPROPERTY Category ", "My category", fieldDocProperty);

        fieldDocVariable = (FieldDocVariable) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DOC_VARIABLE, " DOCVARIABLE  \"My Variable\"", "My variable's value", fieldDocVariable);
        Assert.assertEquals("My Variable", fieldDocVariable.getVariableName());
    }

    @Test
    public void fieldSubject() throws Exception {
        //ExStart
        //ExFor:FieldSubject
        //ExFor:FieldSubject.Text
        //ExSummary:Shows how to use the SUBJECT field.
        Document doc = new Document();

        // Set a value for the document's subject property
        doc.getBuiltInDocumentProperties().setSubject("My subject");

        // We can display this value with a SUBJECT field
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldSubject field = (FieldSubject) builder.insertField(FieldType.FIELD_SUBJECT, true);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " SUBJECT ");
        Assert.assertEquals(field.getResult(), "My subject");

        // We can also set the field's Text attribute to override the current value of the Subject property
        field.setText("My new subject");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " SUBJECT  \"My new subject\"");
        Assert.assertEquals(field.getResult(), "My new subject");

        // As well as displaying a new value in our field, we also changed the value of the document property
        Assert.assertEquals("My new subject", doc.getBuiltInDocumentProperties().getSubject());

        doc.save(getArtifactsDir() + "Field.SUBJECT.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SUBJECT.docx");

        Assert.assertEquals("My new subject", doc.getBuiltInDocumentProperties().getSubject());

        field = (FieldSubject) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SUBJECT, " SUBJECT  \"My new subject\"", "My new subject", field);
        Assert.assertEquals("My new subject", field.getText());
    }

    @Test
    public void fieldComments() throws Exception {
        //ExStart
        //ExFor:FieldComments
        //ExFor:FieldComments.Text
        //ExSummary:Shows how to use the COMMENTS field to display a document's comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // This property is where the COMMENTS field will source its content from
        doc.getBuiltInDocumentProperties().setComments("My comment.");

        // Insert a COMMENTS field with a document builder
        FieldComments field = (FieldComments) builder.insertField(FieldType.FIELD_COMMENTS, true);
        field.update();

        Assert.assertEquals(" COMMENTS ", field.getFieldCode());
        Assert.assertEquals("My comment.", field.getResult());

        // We can override the comment from the document's built in properties and display any text we put here instead
        field.setText("My overriding comment.");
        field.update();

        Assert.assertEquals(" COMMENTS  \"My overriding comment.\"", field.getFieldCode());
        Assert.assertEquals("My overriding comment.", field.getResult());

        doc.save(getArtifactsDir() + "Field.COMMENTS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.COMMENTS.docx");

        Assert.assertEquals("My overriding comment.", doc.getBuiltInDocumentProperties().getComments());

        field = (FieldComments) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_COMMENTS, " COMMENTS  \"My overriding comment.\"", "My overriding comment.", field);
        Assert.assertEquals("My overriding comment.", field.getText());
    }

    @Test
    public void fieldFileSize() throws Exception {
        //ExStart
        //ExFor:FieldFileSize
        //ExFor:FieldFileSize.IsInKilobytes
        //ExFor:FieldFileSize.IsInMegabytes
        //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
        // Open a document and verify its file size
        Document doc = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals(10590, doc.getBuiltInDocumentProperties().getBytes());

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.insertParagraph();

        // By default, file size is displayed in bytes
        FieldFileSize field = (FieldFileSize) builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.update();

        Assert.assertEquals(" FILESIZE ", field.getFieldCode());
        Assert.assertEquals("10590", field.getResult());

        // Set the field to display size in kilobytes
        builder.insertParagraph();
        field = (FieldFileSize) builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.isInKilobytes(true);
        field.update();

        Assert.assertEquals(" FILESIZE  \\k", field.getFieldCode());
        Assert.assertEquals("11", field.getResult());

        // Set the field to display size in megabytes
        builder.insertParagraph();
        field = (FieldFileSize) builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.isInMegabytes(true);
        field.update();

        Assert.assertEquals(" FILESIZE  \\m", field.getFieldCode());
        Assert.assertEquals("0", field.getResult());

        // To update the values of these fields while editing in Microsoft Word,
        // the changes first have to be saved, then the fields manually updated
        doc.save(getArtifactsDir() + "Field.FILESIZE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.FILESIZE.docx");

        Assert.assertEquals(8900, doc.getBuiltInDocumentProperties().getBytes());

        field = (FieldFileSize) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_FILE_SIZE, " FILESIZE ", "10590", field);

        // These fields will need to be updated to produce an accurate result
        doc.updateFields();

        Assert.assertEquals("8900", field.getResult());

        field = (FieldFileSize) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_FILE_SIZE, " FILESIZE  \\k", "9", field);
        Assert.assertTrue(field.isInKilobytes());

        field = (FieldFileSize) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_FILE_SIZE, " FILESIZE  \\m", "0", field);
        Assert.assertTrue(field.isInMegabytes());
    }

    @Test
    public void fieldGoToButton() throws Exception {
        //ExStart
        //ExFor:FieldGoToButton
        //ExFor:FieldGoToButton.DisplayText
        //ExFor:FieldGoToButton.Location
        //ExSummary:Shows to insert a GOTOBUTTON field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a GOTOBUTTON which will take us to a bookmark referenced by "MyBookmark"
        FieldGoToButton field = (FieldGoToButton) builder.insertField(FieldType.FIELD_GO_TO_BUTTON, true);
        field.setDisplayText("My Button");
        field.setLocation("MyBookmark");

        Assert.assertEquals(field.getFieldCode(), " GOTOBUTTON  MyBookmark My Button");

        // Add an arrival destination for our button
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark(field.getLocation());
        builder.writeln("Bookmark text contents.");
        builder.endBookmark(field.getLocation());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.GOTOBUTTON.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.GOTOBUTTON.docx");
        field = (FieldGoToButton) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_GO_TO_BUTTON, " GOTOBUTTON  MyBookmark My Button", "", field);
        Assert.assertEquals("My Button", field.getDisplayText());
        Assert.assertEquals("MyBookmark", field.getLocation());
    }

    @Test
    //ExStart
    //ExFor:FieldFillIn
    //ExFor:FieldFillIn.DefaultResponse
    //ExFor:FieldFillIn.PromptOnceOnMailMerge
    //ExFor:FieldFillIn.PromptText
    //ExSummary:Shows how to use the FILLIN field to prompt the user for a response.
    public void fieldFillIn() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a FILLIN field with a document builder
        FieldFillIn field = (FieldFillIn) builder.insertField(FieldType.FIELD_FILL_IN, true);
        field.setPromptText("Please enter a response:");
        field.setDefaultResponse("A default response");

        // Set this to prompt the user for a response when a mail merge is performed
        field.setPromptOnceOnMailMerge(true);

        Assert.assertEquals(field.getFieldCode(), " FILLIN  \"Please enter a response:\" \\d \"A default response\" \\o");

        // Perform a simple mail merge
        FieldMergeField mergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        mergeField.setFieldName("MergeField");

        doc.getFieldOptions().setUserPromptRespondent(new PromptRespondent());
        doc.getMailMerge().execute(new String[]{"MergeField"}, new Object[]{""});

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FILLIN.docx");
        testFieldFillIn(new Document(getArtifactsDir() + "Field.FILLIN.docx")); //ExSKip
    }

    /// <summary>
    /// IFieldUserPromptRespondent implementation that appends a line to the default response of an FILLIN field during a mail merge.
    /// </summary>
    private static class PromptRespondent implements IFieldUserPromptRespondent {
        public String respond(final String promptText, final String defaultResponse) {
            return "Response modified by PromptRespondent. " + defaultResponse;
        }
    }
    //ExEnd

    private void testFieldFillIn(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(1, doc.getRange().getFields().getCount());

        FieldFillIn field = (FieldFillIn) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_FILL_IN, " FILLIN  \"Please enter a response:\" \\d \"A default response\" \\o",
                "Response modified by PromptRespondent. A default response", field);
        Assert.assertEquals("Please enter a response:", field.getPromptText());
        Assert.assertEquals("A default response", field.getDefaultResponse());
        Assert.assertTrue(field.getPromptOnceOnMailMerge());
    }

    @Test
    public void fieldInfo() throws Exception {
        //ExStart
        //ExFor:FieldInfo
        //ExFor:FieldInfo.InfoType
        //ExFor:FieldInfo.NewValue
        //ExSummary:Shows how to work with INFO fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the value of a document property
        doc.getBuiltInDocumentProperties().setComments("My comment");

        // We can access a property using its name and display it with an INFO field
        // In this case it will be the Comments property
        FieldInfo field = (FieldInfo) builder.insertField(FieldType.FIELD_INFO, true);
        field.setInfoType("Comments");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " INFO  Comments");
        Assert.assertEquals(field.getResult(), "My comment");

        builder.writeln();

        // We can override the value of a document property by setting an INFO field's optional new value
        field = (FieldInfo) builder.insertField(FieldType.FIELD_INFO, true);
        field.setInfoType("Comments");
        field.setNewValue("New comment");
        field.update();

        // Our field's new value has been applied to the corresponding property
        Assert.assertEquals(field.getFieldCode(), " INFO  Comments \"New comment\"");
        Assert.assertEquals(field.getResult(), "New comment");
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getComments(), "New comment");

        doc.save(getArtifactsDir() + "Field.INFO.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INFO.docx");

        Assert.assertEquals("New comment", doc.getBuiltInDocumentProperties().getComments());

        field = (FieldInfo) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INFO, " INFO  Comments", "My comment", field);
        Assert.assertEquals("Comments", field.getInfoType());

        field = (FieldInfo) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INFO, " INFO  Comments \"New comment\"", "New comment", field);
        Assert.assertEquals("Comments", field.getInfoType());
        Assert.assertEquals("New comment", field.getNewValue());
    }

    @Test
    public void fieldMacroButton() throws Exception {
        //ExStart
        //ExFor:Document.HasMacros
        //ExFor:FieldMacroButton
        //ExFor:FieldMacroButton.DisplayText
        //ExFor:FieldMacroButton.MacroName
        //ExSummary:Shows how to use MACROBUTTON fields that enable us to run macros by clicking.
        // Open a document that contains macros
        Document doc = new Document(getMyDir() + "Macro.docm");
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertTrue(doc.hasMacros());

        // Insert a MACROBUTTON field and reference by name a macro that exists within the input document
        FieldMacroButton field = (FieldMacroButton) builder.insertField(FieldType.FIELD_MACRO_BUTTON, true);
        field.setMacroName("MyMacro");
        field.setDisplayText("Double click to run macro: " + field.getMacroName());

        Assert.assertEquals(" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field.getFieldCode());

        // Reference "ViewZoom200", a macro that was shipped with Microsoft Word, found under "Word commands"
        // If our document has a macro of the same name as one from another source, the field will select ours to run
        builder.insertParagraph();
        field = (FieldMacroButton) builder.insertField(FieldType.FIELD_MACRO_BUTTON, true);
        field.setMacroName("ViewZoom200");
        field.setDisplayText("Run " + field.getMacroName());

        Assert.assertEquals(field.getFieldCode(), " MACROBUTTON  ViewZoom200 Run ViewZoom200");

        // Save the document as a macro-enabled document type
        doc.save(getArtifactsDir() + "Field.MACROBUTTON.docm");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MACROBUTTON.docm");

        field = (FieldMacroButton) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_MACRO_BUTTON, " MACROBUTTON  MyMacro Double click to run macro: MyMacro", "", field);
        Assert.assertEquals("MyMacro", field.getMacroName());
        Assert.assertEquals("Double click to run macro: MyMacro", field.getDisplayText());

        field = (FieldMacroButton) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_MACRO_BUTTON, " MACROBUTTON  ViewZoom200 Run ViewZoom200", "", field);
        Assert.assertEquals("ViewZoom200", field.getMacroName());
        Assert.assertEquals("Run ViewZoom200", field.getDisplayText());
    }

    @Test
    public void fieldKeywords() throws Exception {
        //ExStart
        //ExFor:FieldKeywords
        //ExFor:FieldKeywords.Text
        //ExSummary:Shows to insert a KEYWORDS field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some keywords, also referred to as "tags" in File Explorer
        doc.getBuiltInDocumentProperties().setKeywords("Keyword1, Keyword2");

        // Add a KEYWORDS field which will display our keywords
        FieldKeywords field = (FieldKeywords) builder.insertField(FieldType.FIELD_KEYWORD, true);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " KEYWORDS ");
        Assert.assertEquals(field.getResult(), "Keyword1, Keyword2");

        // We can set the Text property of our field to display a different value to the one within the document's properties
        field.setText("OverridingKeyword");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " KEYWORDS  OverridingKeyword");
        Assert.assertEquals(field.getResult(), "OverridingKeyword");

        // Setting a KEYWORDS field's Text property also updates the document's keywords to our new value
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getKeywords(), "OverridingKeyword");

        doc.save(getArtifactsDir() + "Field.KEYWORDS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.KEYWORDS.docx");

        Assert.assertEquals("OverridingKeyword", doc.getBuiltInDocumentProperties().getKeywords());

        field = (FieldKeywords) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_KEYWORD, " KEYWORDS  OverridingKeyword", "OverridingKeyword", field);
        Assert.assertEquals("OverridingKeyword", field.getText());
    }

    @Test
    public void fieldNum() throws Exception {
        //ExStart
        //ExFor:FieldPage
        //ExFor:FieldNumChars
        //ExFor:FieldNumPages
        //ExFor:FieldNumWords
        //ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
        // Open a document to which we want to add character/word/page counts
        Document doc = new Document(getMyDir() + "Paragraphs.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the document builder to the footer, where we will store our fields
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Insert character and word counts
        FieldNumChars fieldNumChars = (FieldNumChars) builder.insertField(FieldType.FIELD_NUM_CHARS, true);
        builder.writeln(" characters");
        FieldNumWords fieldNumWords = (FieldNumWords) builder.insertField(FieldType.FIELD_NUM_WORDS, true);
        builder.writeln(" words");

        // Insert a "Page x of y" page count
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Page ");
        FieldPage fieldPage = (FieldPage) builder.insertField(FieldType.FIELD_PAGE, true);
        builder.write(" of ");
        FieldNumPages fieldNumPages = (FieldNumPages) builder.insertField(FieldType.FIELD_NUM_PAGES, true);

        Assert.assertEquals(fieldNumChars.getFieldCode(), " NUMCHARS ");
        Assert.assertEquals(fieldNumWords.getFieldCode(), " NUMWORDS ");
        Assert.assertEquals(fieldNumPages.getFieldCode(), " NUMPAGES ");
        Assert.assertEquals(fieldPage.getFieldCode(), " PAGE ");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.NUMCHARS.NUMWORDS.NUMPAGES.PAGE.docx");

        TestUtil.verifyField(FieldType.FIELD_NUM_CHARS, " NUMCHARS ", "6009", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_NUM_WORDS, " NUMWORDS ", "1054", doc.getRange().getFields().get(1));
        TestUtil.verifyField(FieldType.FIELD_PAGE, " PAGE ", "6", doc.getRange().getFields().get(2));
        TestUtil.verifyField(FieldType.FIELD_NUM_PAGES, " NUMPAGES ", "6", doc.getRange().getFields().get(3));
    }

    @Test
    public void fieldPrint() throws Exception {
        //ExStart
        //ExFor:FieldPrint
        //ExFor:FieldPrint.PostScriptGroup
        //ExFor:FieldPrint.PrinterInstructions
        //ExSummary:Shows to insert a PRINT field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("My paragraph");

        // The PRINT field can send instructions to the printer that we use to print our document
        FieldPrint field = (FieldPrint) builder.insertField(FieldType.FIELD_PRINT, true);

        // Set the area for the printer to perform instructions over
        // In this case it will be the paragraph that contains our PRINT field
        field.setPostScriptGroup("para");

        // When our document is printed using a printer that supports PostScript,
        // this command will turn the entire area that we specified in field.PostScriptGroup white
        field.setPrinterInstructions("erasepage");

        Assert.assertEquals(" PRINT  erasepage \\p para", field.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.PRINT.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.PRINT.docx");

        field = (FieldPrint) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_PRINT, " PRINT  erasepage \\p para", "", field);
        Assert.assertEquals("para", field.getPostScriptGroup());
        Assert.assertEquals("erasepage", field.getPrinterInstructions());
    }

    @Test
    public void fieldPrintDate() throws Exception {
        //ExStart
        //ExFor:FieldPrintDate
        //ExFor:FieldPrintDate.UseLunarCalendar
        //ExFor:FieldPrintDate.UseSakaEraCalendar
        //ExFor:FieldPrintDate.UseUmAlQuraCalendar
        //ExSummary:Shows read PRINTDATE fields.
        Document doc = new Document(getMyDir() + "Field sample - PRINTDATE.docx");

        // A PRINTDATE field will display "0/0/0000" by default
        // When a document is printed by a printer or printed as a PDF (but not exported as PDF),
        // these fields will display the date/time of that print operation
        FieldPrintDate field = (FieldPrintDate) doc.getRange().getFields().get(0);

        Assert.assertEquals("3/25/2020 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE ", field.getFieldCode());

        // These fields can also display the date using other various international calendars
        field = (FieldPrintDate) doc.getRange().getFields().get(1);

        Assert.assertTrue(field.getUseLunarCalendar());
        Assert.assertEquals("8/1/1441 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE  \\h", field.getFieldCode());

        field = (FieldPrintDate) doc.getRange().getFields().get(2);

        Assert.assertTrue(field.getUseUmAlQuraCalendar());
        Assert.assertEquals("8/1/1441 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE  \\u", field.getFieldCode());

        field = (FieldPrintDate) doc.getRange().getFields().get(3);

        Assert.assertTrue(field.getUseSakaEraCalendar());
        Assert.assertEquals("1/5/1942 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE  \\s", field.getFieldCode());
        //ExEnd
    }

    @Test
    public void fieldQuote() throws Exception {
        //ExStart
        //ExFor:FieldQuote
        //ExFor:FieldQuote.Text
        //ExFor:Document.UpdateFields
        //ExSummary:Shows to use the QUOTE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QUOTE field, which will display content from the Text attribute
        FieldQuote field = (FieldQuote) builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("\"Quoted text\"");

        Assert.assertEquals(" QUOTE  \"\\\"Quoted text\\\"\"", field.getFieldCode());

        // Insert a QUOTE field with a nested DATE field
        // DATE fields normally update their value to the current date every time the document is opened
        // Nesting the DATE field inside the QUOTE field like this will freeze its value to the date when we created the document
        builder.write("\nDocument creation date: ");
        field = (FieldQuote) builder.insertField(FieldType.FIELD_QUOTE, true);
        builder.moveTo(field.getSeparator());
        builder.insertField(FieldType.FIELD_DATE, true);

        Assert.assertEquals(" QUOTE \u0013 DATE \u0014" + LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")) + "\u0015", field.getFieldCode());

        // Some field types don't display the correct result until they are manually updated
        Assert.assertEquals("", doc.getRange().getFields().get(0).getResult());

        doc.updateFields();

        Assert.assertEquals("\"Quoted text\"", doc.getRange().getFields().get(0).getResult());

        doc.save(getArtifactsDir() + "Field.QUOTE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.QUOTE.docx");

        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE  \"\\\"Quoted text\\\"\"", "\"Quoted text\"", doc.getRange().getFields().get(0));

        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE \u0013 DATE \u0014" + LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")) + "\u0015",
                LocalDate.now().format(DateTimeFormatter.ofPattern("M/d/YYYY")), doc.getRange().getFields().get(1));

    }

    //ExStart
    //ExFor:FieldNext
    //ExFor:FieldNextIf
    //ExFor:FieldNextIf.ComparisonOperator
    //ExFor:FieldNextIf.LeftExpression
    //ExFor:FieldNextIf.RightExpression
    //ExSummary:Shows how to use NEXT/NEXTIF fields to merge more than one row into one page during a mail merge.
    @Test //ExSkip
    public void fieldNext() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a data source for our mail merge with 3 rows,
        // This would normally amount to 3 pages in the output of a mail merge
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");
        table.getRows().add("Mr.", "Joe", "Bloggs");

        // Insert a set of merge fields
        insertMergeFields(builder, "First row: ");

        // If we have multiple merge fields with the same FieldName,
        // they will receive data from the same row of the data source and will display the same value after the merge
        // A NEXT field tells the mail merge instantly to move down one row,
        // so any upcoming merge fields will have data deposited from the next row
        // Make sure not to skip with a NEXT/NEXTIF field while on the last row
        FieldNext fieldNext = (FieldNext) builder.insertField(FieldType.FIELD_NEXT, true);

        Assert.assertEquals(" NEXT ", fieldNext.getFieldCode());

        // These merge fields are the same as the ones as above but will take values from the second row
        insertMergeFields(builder, "Second row: ");

        // A NEXTIF field has the same function as a NEXT field,
        // but it skips to the next row only if a condition expressed by the following 3 attributes is fulfilled
        FieldNextIf fieldNextIf = (FieldNextIf) builder.insertField(FieldType.FIELD_NEXT_IF, true);
        fieldNextIf.setLeftExpression("5");
        fieldNextIf.setRightExpression("2 + 3");
        fieldNextIf.setComparisonOperator("=");

        // If the comparison asserted by the above field is correct,
        // the following 3 merge fields will take data from the third row
        // Otherwise, these fields will take data from row 2 again
        insertMergeFields(builder, "Third row: ");

        // Our data source has 3 rows and we skipped rows twice, so our output will have one page
        // with data from all 3 rows
        doc.getMailMerge().execute(table);

        Assert.assertEquals(" NEXTIF  5 = \"2 + 3\"", fieldNextIf.getFieldCode());

        doc.save(getArtifactsDir() + "Field.NEXT.NEXTIF.docx");
        testFieldNext(doc); //ExSKip
    }

    /// <summary>
    /// Uses a document builder to insert merge fields for a data table that has "Courtesy Title", "First Name" and "Last Name" columns.
    /// </summary>
    @Test(enabled = false)
    public void insertMergeFields(final DocumentBuilder builder, final String firstFieldTextBefore) throws Exception {
        insertMergeField(builder, "Courtesy Title", firstFieldTextBefore, " ");
        insertMergeField(builder, "First Name", null, " ");
        insertMergeField(builder, "Last Name", null, null);
        builder.insertParagraph();
    }

    /// <summary>
    /// Uses a document builder to insert a merge field.
    /// </summary>
    @Test(enabled = false)
    public void insertMergeField(final DocumentBuilder builder, final String fieldName, final String textBefore, final String textAfter) throws Exception {
        FieldMergeField field = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        field.setFieldName(fieldName);
        field.setTextBefore(textBefore);
        field.setTextAfter(textAfter);
    }
    //ExEnd

    private void testFieldNext(Document doc) throws Exception {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals("First row: Mr. John Doe\r" +
                "Second row: Mrs. Jane Cardholder\r" +
                "Third row: Mr. Joe Bloggs\r\f", doc.getText());
    }

    //ExStart
    //ExFor:FieldNoteRef
    //ExFor:FieldNoteRef.BookmarkName
    //ExFor:FieldNoteRef.InsertHyperlink
    //ExFor:FieldNoteRef.InsertReferenceMark
    //ExFor:FieldNoteRef.InsertRelativePosition
    //ExSummary:Shows to insert NOTEREF fields and modify their appearance.
    @Test(enabled = false, description = "WORDSNET-17845") //ExSkip
    public void fieldNoteRef() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a boomkark with a footnote for the NOTEREF field to reference
        insertBookmarkWithFootnote(builder, "MyBookmark1", "Contents of MyBookmark1", "Footnote from MyBookmark1");

        // This NOTEREF field will display just the number of the footnote inside the referenced bookmark
        // Setting the InsertHyperlink attribute lets us jump to the bookmark by Ctrl + clicking the field
        Assert.assertEquals(" NOTEREF  MyBookmark2 \\h",
                insertFieldNoteRef(builder, "MyBookmark2", true, false, false, "Hyperlink to Bookmark2, with footnote number ").getFieldCode());

        // When using the \p flag, after the footnote number the field also displays the position of the bookmark relative to the field
        // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update
        Assert.assertEquals(" NOTEREF  MyBookmark1 \\h \\p",
                insertFieldNoteRef(builder, "MyBookmark1", true, true, false, "Bookmark1, with footnote number ").getFieldCode());

        // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below"
        // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text
        Assert.assertEquals(" NOTEREF  MyBookmark2 \\h \\p \\f",
                insertFieldNoteRef(builder, "MyBookmark2", true, true, true, "Bookmark2, with footnote number ").getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        insertBookmarkWithFootnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.NOTEREF.docx");
        testNoteRef(new Document(getArtifactsDir() + "Field.NOTEREF.docx")); //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a NOTEREF field and sets its attributes.
    /// </summary>
    private static FieldNoteRef insertFieldNoteRef(DocumentBuilder builder, String bookmarkName, boolean insertHyperlink, boolean insertRelativePosition, boolean insertReferenceMark, String textBefore) throws Exception {
        builder.write(textBefore);

        FieldNoteRef field = (FieldNoteRef) builder.insertField(FieldType.FIELD_NOTE_REF, true);
        field.setBookmarkName(bookmarkName);
        field.setInsertHyperlink(insertHyperlink);
        field.setInsertRelativePosition(insertRelativePosition);
        field.setInsertReferenceMark(insertReferenceMark);
        builder.writeln();

        return field;
    }

    /// <summary>
    /// Uses a document builder to insert a named bookmark with a footnote at the end.
    /// </summary>
    private void insertBookmarkWithFootnote(final DocumentBuilder builder, final String bookmarkName,
                                            final String bookmarkText, final String footnoteText) {
        builder.startBookmark(bookmarkName);
        builder.write(bookmarkText);
        builder.insertFootnote(FootnoteType.FOOTNOTE, footnoteText);
        builder.endBookmark(bookmarkName);
        builder.writeln();
    }
    //ExEnd

    private void testNoteRef(Document doc) {
        FieldNoteRef field = (FieldNoteRef) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark2 \\h", "2", field);
        Assert.assertEquals("MyBookmark2", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertFalse(field.getInsertRelativePosition());
        Assert.assertFalse(field.getInsertReferenceMark());

        field = (FieldNoteRef) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark1 \\h \\p", "1 above", field);
        Assert.assertEquals("MyBookmark1", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());
        Assert.assertFalse(field.getInsertReferenceMark());

        field = (FieldNoteRef) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark2 \\h \\p \\f", "2 below", field);
        Assert.assertEquals("MyBookmark2", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());
        Assert.assertTrue(field.getInsertReferenceMark());
    }

    @Test(enabled = false, description = "WORDSNET-17845")
    public void footnoteRef() throws Exception {
        //ExStart
        //ExFor:FieldFootnoteRef
        //ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field.
        // Create a blank document and a document builder for it
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some text, and a footnote, all inside a bookmark named "CrossRefBookmark"
        builder.startBookmark("CrossRefBookmark");
        builder.write("Hello world!");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Cross referenced footnote.");
        builder.endBookmark("CrossRefBookmark");

        builder.insertParagraph();
        builder.write("CrossReference: ");

        // Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker
        FieldFootnoteRef field = (FieldFootnoteRef) builder.insertField(FieldType.FIELD_FOOTNOTE_REF, true);

        // Get this field to reference a bookmark
        // The bookmark that we chose contains a footnote marker belonging to the footnote we inserted, which will be displayed by the field, just by itself
        builder.moveTo(field.getSeparator());
        builder.write("CrossRefBookmark");

        Assert.assertEquals(field.getFieldCode(), " FOOTNOTEREF CrossRefBookmark");

        doc.updateFields();

        // This field works only in older versions of Microsoft Word
        doc.save(getArtifactsDir() + "Field.FOOTNOTEREF.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.FOOTNOTEREF.doc");
        field = (FieldFootnoteRef) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_FOOTNOTE_REF, " FOOTNOTEREF CrossRefBookmark", "1", field);
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "Cross referenced footnote.",
                (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
    }

    //ExStart
    //ExFor:FieldPageRef
    //ExFor:FieldPageRef.BookmarkName
    //ExFor:FieldPageRef.InsertHyperlink
    //ExFor:FieldPageRef.InsertRelativePosition
    //ExSummary:Shows to insert PAGEREF fields and present them in different ways.
    @Test(enabled = false, description = "WORDSNET-17836") //ExSkip
    public void fieldPageRef() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        insertAndNameBookmark(builder, "MyBookmark1");

        // This field will display just the page number where the bookmark starts
        // Setting InsertHyperlink attribute makes the field function as a link to the bookmark
        Assert.assertEquals(" PAGEREF  MyBookmark3 \\h",
                insertFieldPageRef(builder, "MyBookmark3", true, false, "Hyperlink to Bookmark3, on page: ").getFieldCode());

        // Setting the \p flag makes the field display the relative position of the bookmark to the field instead of a page number
        // Bookmark1 is on the same page and above this field, so the result will be "above" on update
        Assert.assertEquals(" PAGEREF  MyBookmark1 \\h \\p",
                insertFieldPageRef(builder, "MyBookmark1", true, true, "Bookmark1 is ").getFieldCode());

        // Bookmark2 will be on the same page and below this field, so the field will display "below"
        Assert.assertEquals(" PAGEREF  MyBookmark2 \\h \\p",
                insertFieldPageRef(builder, "MyBookmark2", true, true, "Bookmark2 is ").getFieldCode());

        // Bookmark3 will be on a different page, so the field will display "on page 2"
        Assert.assertEquals(" PAGEREF  MyBookmark3 \\h \\p",
                insertFieldPageRef(builder, "MyBookmark3", true, true, "Bookmark3 is ").getFieldCode());

        insertAndNameBookmark(builder, "MyBookmark2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        insertAndNameBookmark(builder, "MyBookmark3");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.PAGEREF.docx");
        testPageRef(new Document(getArtifactsDir() + "Field.PAGEREF.docx")); //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its attributes.
    /// </summary>
    private FieldPageRef insertFieldPageRef(final DocumentBuilder builder, final String bookmarkName, final boolean insertHyperlink,
                                            final boolean insertRelativePosition, final String textBefore) throws Exception {
        builder.write(textBefore);

        FieldPageRef field = (FieldPageRef) builder.insertField(FieldType.FIELD_PAGE_REF, true);
        field.setBookmarkName(bookmarkName);
        field.setInsertHyperlink(insertHyperlink);
        field.setInsertRelativePosition(insertRelativePosition);
        builder.writeln();

        return field;
    }

    /// <summary>
    /// Uses a document builder to insert a named bookmark.
    /// </summary>
    private void insertAndNameBookmark(final DocumentBuilder builder, final String bookmarkName) {
        builder.startBookmark(bookmarkName);
        builder.writeln(MessageFormat.format("Contents of bookmark \"{0}\".", bookmarkName));
        builder.endBookmark(bookmarkName);
    }
    //ExEnd

    private void testPageRef(Document doc) {
        FieldPageRef field = (FieldPageRef) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark3 \\h", "2", field);
        Assert.assertEquals("MyBookmark3", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertFalse(field.getInsertRelativePosition());

        field = (FieldPageRef) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark1 \\h \\p", "above", field);
        Assert.assertEquals("MyBookmark1", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());

        field = (FieldPageRef) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark2 \\h \\p", "below", field);
        Assert.assertEquals("MyBookmark2", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());

        field = (FieldPageRef) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark3 \\h \\p", "on page 2", field);
        Assert.assertEquals("MyBookmark3", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());
    }

    //ExStart
    //ExFor:FieldRef
    //ExFor:FieldRef.BookmarkName
    //ExFor:FieldRef.IncludeNoteOrComment
    //ExFor:FieldRef.InsertHyperlink
    //ExFor:FieldRef.InsertParagraphNumber
    //ExFor:FieldRef.InsertParagraphNumberInFullContext
    //ExFor:FieldRef.InsertParagraphNumberInRelativeContext
    //ExFor:FieldRef.InsertRelativePosition
    //ExFor:FieldRef.NumberSeparator
    //ExFor:FieldRef.SuppressNonDelimiters
    //ExSummary:Shows how to insert REF fields to reference bookmarks and present them in various ways.
    @Test(enabled = false, description = "WORDSNET-18067") //ExSkip
    public void fieldRef() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the bookmark that all our REF fields will reference and leave it at the end of the document
        builder.startBookmark("MyBookmark");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "MyBookmark footnote #1");
        builder.write("Text that will appear in REF field");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "MyBookmark footnote #2");
        builder.endBookmark("MyBookmark");
        builder.moveToDocumentStart();

        // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at
        // Note that the angle brackets count as non-delimiter characters
        builder.getListFormat().applyNumberDefault();
        builder.getListFormat().getListLevel().setNumberFormat("> \u0000");

        // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes
        FieldRef field = insertFieldRef(builder, "MyBookmark", "", "\n");
        field.setIncludeNoteOrComment(true);
        field.setInsertHyperlink(true);

        Assert.assertEquals(field.getFieldCode(), " REF  MyBookmark \\f \\h");

        // Insert a REF field and display whether the referenced bookmark is above or below it
        field = insertFieldRef(builder, "MyBookmark", "The referenced paragraph is ", " this field.\n");
        field.setInsertRelativePosition(true);

        Assert.assertEquals(field.getFieldCode(), " REF  MyBookmark \\p");

        // Display the list number of the bookmark, as it appears in the document
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number is ", "\n");
        field.setInsertParagraphNumber(true);

        Assert.assertEquals(" REF  MyBookmark \\n", field.getFieldCode());

        // Display the list number of the bookmark, but with non-delimiter characters omitted
        // In this case they are the angle brackets
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number, non-delimiters suppressed, is ", "\n");
        field.setInsertParagraphNumber(true);
        field.setSuppressNonDelimiters(true);

        Assert.assertEquals(field.getFieldCode(), " REF  MyBookmark \\n \\t");

        // Move down one list level
        builder.getListFormat().setListLevelNumber(builder.getListFormat().getListLevelNumber() + 1)/*Property++*/;
        builder.getListFormat().getListLevel().setNumberFormat(">> \u0001");

        // Display the list number of the bookmark as well as the numbers of all the list levels above it
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's full context paragraph number is ", "\n");
        field.setInsertParagraphNumberInFullContext(true);

        Assert.assertEquals(field.getFieldCode(), " REF  MyBookmark \\w");

        builder.insertBreak(BreakType.PAGE_BREAK);

        // Display the list level numbers between this REF field and the bookmark that it is referencing
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's relative paragraph number is ", "\n");
        field.setInsertParagraphNumberInRelativeContext(true);

        Assert.assertEquals(field.getFieldCode(), " REF  MyBookmark \\r");

        // The bookmark, which is at the end of the document, will show up as a list item here
        builder.writeln("List level above bookmark");
        builder.getListFormat().setListLevelNumber(builder.getListFormat().getListLevelNumber() + 1)/*Property++*/;
        builder.getListFormat().getListLevel().setNumberFormat(">>> \u0002");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.REF.docx");
        testFieldRef(new Document(getArtifactsDir() + "Field.REF.docx")); //ExSkip
    }

    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after.
    /// </summary>
    private FieldRef insertFieldRef(final DocumentBuilder builder, final String bookmarkName,
                                    final String textBefore, final String textAfter) throws Exception {
        builder.write(textBefore);
        FieldRef field = (FieldRef) builder.insertField(FieldType.FIELD_REF, true);
        field.setBookmarkName(bookmarkName);
        builder.write(textAfter);
        return field;
    }
    //ExEnd

    private void testFieldRef(Document doc) throws Exception {
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "MyBookmark footnote #1",
                (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "MyBookmark footnote #2",
                (Footnote) doc.getChild(NodeType.FOOTNOTE, 0, true));

        FieldRef field = (FieldRef) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\f \\h",
                "\u0002 MyBookmark footnote #1\r" +
                        "Text that will appear in REF field\u0002 MyBookmark footnote #2\r", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getIncludeNoteOrComment());
        Assert.assertTrue(field.getInsertHyperlink());

        field = (FieldRef) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\p", "below", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertRelativePosition());

        field = (FieldRef) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\n", ">>> i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumber());
        Assert.assertEquals(" REF  MyBookmark \\n", field.getFieldCode());
        Assert.assertEquals(">>> i", field.getResult());

        field = (FieldRef) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\n \\t", "i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumber());
        Assert.assertTrue(field.getSuppressNonDelimiters());

        field = (FieldRef) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\w", "> 4>> c>>> i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumberInFullContext());

        field = (FieldRef) doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\r", ">> c>>> i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumberInRelativeContext());
    }

    @Test(enabled = false, description = "WORDSNET-18068")
    public void fieldRD() throws Exception {
        //ExStart
        //ExFor:FieldRD
        //ExFor:FieldRD.FileName
        //ExFor:FieldRD.IsPathRelative
        //ExSummary:Shows to insert an RD field to source table of contents entries from an external document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a table of contents and, on the following page, one entry
        builder.insertField(FieldType.FIELD_TOC, true);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.getCurrentParagraph().getParagraphFormat().setStyleName("Heading 1");
        builder.writeln("TOC entry from within this document");

        // Insert an RD field, designating an external document that our TOC field will look in for more entries
        FieldRD field = (FieldRD) builder.insertField(FieldType.FIELD_REF_DOC, true);
        field.setFileName("ReferencedDocument.docx");
        field.isPathRelative(true);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " RD  ReferencedDocument.docx \\f");

        // Create the document and insert a TOC entry, which will end up in the TOC of our original document
        Document referencedDoc = new Document();
        DocumentBuilder refDocBuilder = new DocumentBuilder(referencedDoc);
        refDocBuilder.getCurrentParagraph().getParagraphFormat().setStyleName("Heading 1");
        refDocBuilder.writeln("TOC entry from referenced document");
        referencedDoc.save(getArtifactsDir() + "ReferencedDocument.docx");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.RD.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.RD.docx");

        FieldToc fieldToc = (FieldToc) doc.getRange().getFields().get(0);

        Assert.assertEquals("TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r" +
                "TOC entry from referenced document\t1\r", fieldToc.getResult());

        FieldPageRef fieldPageRef = (FieldPageRef) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF _Toc36149519 \\h ", "2", fieldPageRef);

        field = (FieldRD) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_REF_DOC, " RD  ReferencedDocument.docx \\f", "", field);
        Assert.assertEquals("ReferencedDocument.docx", field.getFileName());
        Assert.assertTrue(field.isPathRelative());
    }

    @Test
    public void skipIf() throws Exception {
        //ExStart
        //ExFor:FieldSkipIf
        //ExFor:FieldSkipIf.ComparisonOperator
        //ExFor:FieldSkipIf.LeftExpression
        //ExFor:FieldSkipIf.RightExpression
        //ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a data table that will be the source for our mail merge
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Name");
        table.getColumns().add("Department");
        table.getRows().add("John Doe", "Sales");
        table.getRows().add("Jane Doe", "Accounting");
        table.getRows().add("John Cardholder", "HR");

        // Insert a SKIPIF field, which will skip a page of a mail merge if the condition is fulfilled
        // We will move to the SKIPIF field's separator character and insert a MERGEFIELD at that place to create a nested field
        FieldSkipIf fieldSkipIf = (FieldSkipIf) builder.insertField(FieldType.FIELD_SKIP_IF, true);
        builder.moveTo(fieldSkipIf.getSeparator());
        FieldMergeField fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Department");

        // The MERGEFIELD refers to the "Department" column in our data table, and our SKIPIF field will check if its value equals to "HR"
        // One of three rows satisfies that condition, so we will expect the result of our mail merge to have two pages
        fieldSkipIf.setLeftExpression("=");
        fieldSkipIf.setRightExpression("HR");

        // Add some content to our mail merge and execute it
        builder.moveToDocumentEnd();
        builder.write("Dear ");
        fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Name");
        builder.writeln(", ");

        doc.getMailMerge().execute(table);
        doc.save(getArtifactsDir() + "Field.SKIPIF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SKIPIF.docx");

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals("Dear John Doe, \r" +
                "\fDear Jane Doe, \r\f", doc.getText());
    }

    @Test
    public void fieldSet() throws Exception {
        //ExStart
        //ExFor:FieldSet
        //ExFor:FieldSet.BookmarkName
        //ExFor:FieldSet.BookmarkText
        //ExSummary:Shows to alter a bookmark's text with a SET field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.writeln("Bookmark contents");
        builder.endBookmark("MyBookmark");

        Bookmark bookmark = doc.getRange().getBookmarks().get("MyBookmark");
        bookmark.setText("Old text");

        FieldSet field = (FieldSet) builder.insertField(FieldType.FIELD_SET, false);
        field.setBookmarkName("MyBookmark");
        field.setBookmarkText("New text");

        Assert.assertEquals(field.getFieldCode(), " SET  MyBookmark \"New text\"");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SET.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SET.docx");

        Assert.assertEquals("New text", doc.getRange().getBookmarks().get(0).getText());

        field = (FieldSet) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SET, " SET  MyBookmark \"New text\"", "New text", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertEquals("New text", field.getBookmarkText());
    }

    @Test(enabled = false, description = "WORDSNET-18137")
    public void fieldTemplate() throws Exception {
        //ExStart
        //ExFor:FieldTemplate
        //ExFor:FieldTemplate.IncludeFullPath
        //ExSummary:Shows how to display the location of the document's template with a TEMPLATE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldTemplate field = (FieldTemplate) builder.insertField(FieldType.FIELD_TEMPLATE, false);
        Assert.assertEquals(field.getFieldCode(), " TEMPLATE ");

        builder.writeln();
        field = (FieldTemplate) builder.insertField(FieldType.FIELD_TEMPLATE, false);
        field.setIncludeFullPath(true);

        Assert.assertEquals(field.getFieldCode(), " TEMPLATE  \\p");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TEMPLATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.TEMPLATE.docx");

        field = (FieldTemplate) doc.getRange().getFields().get(0);
        Assert.assertEquals(" TEMPLATE ", field.getFieldCode());
        Assert.assertEquals("Normal.dotm", field.getResult());

        field = (FieldTemplate) doc.getRange().getFields().get(1);
        Assert.assertEquals(" TEMPLATE  \\p", field.getFieldCode());
        Assert.assertTrue(field.getResult().endsWith("\\Microsoft\\Templates\\Normal.dotm"));

    }

    @Test
    public void fieldSymbol() throws Exception {
        //ExStart
        //ExFor:FieldSymbol
        //ExFor:FieldSymbol.CharacterCode
        //ExFor:FieldSymbol.DontAffectsLineSpacing
        //ExFor:FieldSymbol.FontName
        //ExFor:FieldSymbol.FontSize
        //ExFor:FieldSymbol.IsAnsi
        //ExFor:FieldSymbol.IsShiftJis
        //ExFor:FieldSymbol.IsUnicode
        //ExSummary:Shows how to use the SYMBOL field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a SYMBOL field to display a symbol, designated by a character code
        FieldSymbol field = (FieldSymbol) builder.insertField(FieldType.FIELD_SYMBOL, true);

        // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol
        field.setCharacterCode(Integer.toString(0x00a9));
        field.isAnsi(true);

        Assert.assertEquals(field.getFieldCode(), " SYMBOL  169 \\a");

        builder.writeln(" Line 1");

        // In Unicode, the "221E" code is reserved for ths infinity symbol
        field = (FieldSymbol) builder.insertField(FieldType.FIELD_SYMBOL, true);
        field.setCharacterCode(Integer.toString(0x221E));
        field.isUnicode(true);

        // Change the appearance of our symbol
        // Note that some symbols can change from font to font
        // The full list of symbols and their fonts can be looked up in the Windows Character Map
        field.setFontName("Calibri");
        field.setFontSize("24");

        // A tall symbol like the one we placed can also be made to not push down the text on its line
        field.setDontAffectsLineSpacing(true);

        Assert.assertEquals(field.getFieldCode(), " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h");

        builder.writeln("Line 2");

        // Display a symbol from the Shift-JIS, also known as the Windows-932 code page
        // With a font that supports Shift-JIS, this symbol will display "あ"
        field = (FieldSymbol) builder.insertField(FieldType.FIELD_SYMBOL, true);
        field.setFontName("MS Gothic");
        field.setCharacterCode(Integer.toString(0x82A0));
        field.isShiftJis(true);

        Assert.assertEquals(field.getFieldCode(), " SYMBOL  33440 \\f \"MS Gothic\" \\j");

        builder.write("Line 3");

        doc.save(getArtifactsDir() + "Field.SYMBOL.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SYMBOL.docx");

        field = (FieldSymbol) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL  169 \\a", "", field);
        Assert.assertEquals(Integer.toString(0x00a9), field.getCharacterCode());
        Assert.assertTrue(field.isAnsi());
        Assert.assertEquals("©", field.getDisplayResult());

        field = (FieldSymbol) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", "", field);
        Assert.assertEquals(Integer.toString(0x221E), field.getCharacterCode());
        Assert.assertEquals("Calibri", field.getFontName());
        Assert.assertEquals("24", field.getFontSize());
        Assert.assertTrue(field.isUnicode());
        Assert.assertTrue(field.getDontAffectsLineSpacing());
        Assert.assertEquals("∞", field.getDisplayResult());

        field = (FieldSymbol) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL  33440 \\f \"MS Gothic\" \\j", "", field);
        Assert.assertEquals(Integer.toString(0x82A0), field.getCharacterCode());
        Assert.assertEquals("MS Gothic", field.getFontName());
        Assert.assertTrue(field.isShiftJis());
    }

    @Test
    public void fieldTitle() throws Exception {
        //ExStart
        //ExFor:FieldTitle
        //ExFor:FieldTitle.Text
        //ExSummary:Shows how to use the TITLE field.
        Document doc = new Document();

        // A TITLE field will display the value assigned to this variable
        doc.getBuiltInDocumentProperties().setTitle("My Title");

        // Insert a TITLE field using a document builder
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldTitle field = (FieldTitle) builder.insertField(FieldType.FIELD_TITLE, false);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " TITLE ");
        Assert.assertEquals(field.getResult(), "My Title");

        // Set the Text attribute to display a different value
        builder.writeln();
        field = (FieldTitle) builder.insertField(FieldType.FIELD_TITLE, false);
        field.setText("My New Title");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " TITLE  \"My New Title\"");
        Assert.assertEquals(field.getResult(), "My New Title");

        // In doing that we've also changed the title in the document properties
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getTitle(), "My New Title");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TITLE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.TITLE.docx");

        Assert.assertEquals("My New Title", doc.getBuiltInDocumentProperties().getTitle());

        field = (FieldTitle) doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_TITLE, " TITLE ", "My New Title", field);

        field = (FieldTitle) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_TITLE, " TITLE  \"My New Title\"", "My New Title", field);
        Assert.assertEquals("My New Title", field.getText());
    }

    //ExStart
    //ExFor:FieldToa
    //ExFor:FieldToa.BookmarkName
    //ExFor:FieldToa.EntryCategory
    //ExFor:FieldToa.EntrySeparator
    //ExFor:FieldToa.PageNumberListSeparator
    //ExFor:FieldToa.PageRangeSeparator
    //ExFor:FieldToa.RemoveEntryFormatting
    //ExFor:FieldToa.SequenceName
    //ExFor:FieldToa.SequenceSeparator
    //ExFor:FieldToa.UseHeading
    //ExFor:FieldToa.UsePassim
    //ExFor:FieldTA
    //ExFor:FieldTA.EntryCategory
    //ExFor:FieldTA.IsBold
    //ExFor:FieldTA.IsItalic
    //ExFor:FieldTA.LongCitation
    //ExFor:FieldTA.PageRangeBookmarkName
    //ExFor:FieldTA.ShortCitation
    //ExSummary:Shows how to build and customize a table of authorities using TOA and TA fields.
    @Test //ExSkip
    public void fieldTOA() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOA field, which will list all the TA entries in the document,
        // displaying long citations and page numbers for each
        FieldToa fieldToa = (FieldToa) builder.insertField(FieldType.FIELD_TOA, false);

        // Set the entry category for our table
        // For a TA field to be included in this table, it will have to have a matching entry category
        fieldToa.setEntryCategory("1");

        // Moreover, the Table of Authorities category at index 1 is "Cases",
        // which will show up as the title of our table if we set this variable to true
        fieldToa.setUseHeading(true);

        // We can further filter TA fields by designating a named bookmark that they have to be inside of
        fieldToa.setBookmarkName("MyBookmark");

        // By default, a dotted line page-wide tab appears between the TA field's citation and its page number
        // We can replace it with any text we put in this attribute, even preserving the tab if we use tab character
        fieldToa.setEntrySeparator(" \t p.");

        // If we have multiple TA entries that share the same long citation,
        // all their respective page numbers will show up on one row,
        // and the page numbers separated by a string specified here
        fieldToa.setPageNumberListSeparator(" & p. ");

        // To reduce clutter, we can set this to true to get our table to display the word "passim"
        // if there would be 5 or more page numbers in one row
        fieldToa.setUsePassim(true);

        // One TA field can refer to a range of pages, and the sequence specified here will be between the start and end page numbers
        fieldToa.setPageRangeSeparator(" to ");

        // The format from the TA fields will carry over into our table, and we can stop it from doing so by setting this variable
        fieldToa.setRemoveEntryFormatting(true);
        builder.getFont().setColor(Color.GREEN);
        builder.getFont().setName("Arial Black");

        Assert.assertEquals(fieldToa.getFieldCode(), " TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f");

        builder.insertBreak(BreakType.PAGE_BREAK);

        // We will insert a TA entry using a document builder
        // This entry is outside the bookmark specified by our table, so it won't be displayed
        FieldTA fieldTA = insertToaEntry(builder, "1", "Source 1");

        Assert.assertEquals(fieldTA.getFieldCode(), " TA  \\c 1 \\l \"Source 1\"");

        // This entry is inside the bookmark,
        // but the entry category doesn't match that of the table, so it will also be omitted
        builder.startBookmark("MyBookmark");
        fieldTA = insertToaEntry(builder, "2", "Source 2");

        // This entry will appear in the table
        fieldTA = insertToaEntry(builder, "1", "Source 3");

        // Short citations aren't displayed by a TOA table,
        // but they can be used as a shorthand to refer to bulky source names that multiple TA fields reference
        fieldTA.setShortCitation("S.3");

        Assert.assertEquals(fieldTA.getFieldCode(), " TA  \\c 1 \\l \"Source 3\" \\s S.3");

        // The page number can be made to appear bold and/or italic
        // This will still be displayed if our table is set to ignore formatting
        fieldTA = insertToaEntry(builder, "1", "Source 2");
        fieldTA.isBold(true);
        fieldTA.isItalic(true);

        Assert.assertEquals(fieldTA.getFieldCode(), " TA  \\c 1 \\l \"Source 2\" \\b \\i");

        // We can get TA fields to refer to a range of pages that a bookmark spans across instead of the page that they are on
        // Note that this entry refers to the same source as the one above, so they will share one row in our table,
        // displaying the page number of the entry above as well as the page range of this entry,
        // with the table's page list and page number range separators between page numbers
        fieldTA = insertToaEntry(builder, "1", "Source 3");
        fieldTA.setPageRangeBookmarkName("MyMultiPageBookmark");

        builder.startBookmark("MyMultiPageBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.endBookmark("MyMultiPageBookmark");

        Assert.assertEquals(fieldTA.getFieldCode(), " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark");

        // Having 5 or more TA entries with the same source invokes the "passim" feature of our table, if we enabled it
        for (int i = 0; i < 5; i++) {
            insertToaEntry(builder, "1", "Source 4");
        }

        builder.endBookmark("MyBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOA.TA.docx");
        testFieldTOA(new Document(getArtifactsDir() + "Field.TOA.TA.docx")); //ExSKip
    }

    /// <summary>
    /// Get a builder to insert a TA field, specifying its long citation and category,
    /// then insert a page break and return the field we created.
    /// </summary>
    private FieldTA insertToaEntry(final DocumentBuilder builder, final String entryCategory, final String longCitation) throws Exception {
        FieldTA field = (FieldTA) builder.insertField(FieldType.FIELD_TOA_ENTRY, false);
        field.setEntryCategory(entryCategory);
        field.setLongCitation(longCitation);

        builder.insertBreak(BreakType.PAGE_BREAK);

        return field;
    }
    //ExEnd

    private void testFieldTOA(Document doc) {
        FieldToa fieldTOA = (FieldToa) doc.getRange().getFields().get(0);

        Assert.assertEquals("1", fieldTOA.getEntryCategory());
        Assert.assertTrue(fieldTOA.getUseHeading());
        Assert.assertEquals("MyBookmark", fieldTOA.getBookmarkName());
        Assert.assertEquals(" \t p.", fieldTOA.getEntrySeparator());
        Assert.assertEquals(" & p. ", fieldTOA.getPageNumberListSeparator());
        Assert.assertTrue(fieldTOA.getUsePassim());
        Assert.assertEquals(" to ", fieldTOA.getPageRangeSeparator());
        Assert.assertTrue(fieldTOA.getRemoveEntryFormatting());
        Assert.assertEquals(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldTOA.getFieldCode());
        Assert.assertEquals("Cases\r" +
                "Source 2 \t p.5\r" +
                "Source 3 \t p.4 & p. 7 to 10\r" +
                "Source 4 \t p.passim\r", fieldTOA.getResult());

        FieldTA fieldTA = (FieldTA) doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 1\"", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 1", fieldTA.getLongCitation());

        fieldTA = (FieldTA) doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 2 \\l \"Source 2\"", "", fieldTA);
        Assert.assertEquals("2", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 2", fieldTA.getLongCitation());

        fieldTA = (FieldTA) doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 3\" \\s S.3", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 3", fieldTA.getLongCitation());
        Assert.assertEquals("S.3", fieldTA.getShortCitation());

        fieldTA = (FieldTA) doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 2\" \\b \\i", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 2", fieldTA.getLongCitation());
        Assert.assertTrue(fieldTA.isBold());
        Assert.assertTrue(fieldTA.isItalic());

        fieldTA = (FieldTA) doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 3", fieldTA.getLongCitation());
        Assert.assertEquals("MyMultiPageBookmark", fieldTA.getPageRangeBookmarkName());

        for (int i = 6; i < 11; i++) {
            fieldTA = (FieldTA) doc.getRange().getFields().get(i);

            TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 4\"", "", fieldTA);
            Assert.assertEquals("1", fieldTA.getEntryCategory());
            Assert.assertEquals("Source 4", fieldTA.getLongCitation());
        }
    }

    @Test
    public void fieldAddIn() throws Exception {
        //ExStart
        //ExFor:FieldAddIn
        //ExSummary:Shows how to process an ADDIN field.
        // Open a document that contains an ADDIN field
        Document doc = new Document(getMyDir() + "Field sample - ADDIN.docx");

        // Aspose.Words does not support inserting ADDIN fields, they can be read
        FieldAddIn field = (FieldAddIn) doc.getRange().getFields().get(0);

        Assert.assertEquals(" ADDIN \"My value\" ", field.getFieldCode());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.verifyField(FieldType.FIELD_ADDIN, " ADDIN \"My value\" ", "", doc.getRange().getFields().get(0));
    }

    @Test
    public void fieldEditTime() throws Exception {
        //ExStart
        //ExFor:FieldEditTime
        //ExSummary:Shows how to use the EDITTIME field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert an EDITTIME field in the header
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("You've been editing this document for ");
        FieldEditTime field = (FieldEditTime) builder.insertField(FieldType.FIELD_EDIT_TIME, true);
        builder.writeln(" minutes.");

        // The EDITTIME field will show, in minutes only,
        // the time spent with the document open in a Microsoft Word window
        // The minutes are tracked in a document property, which we can change like this
        doc.getBuiltInDocumentProperties().setTotalEditingTime(10);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " EDITTIME ");
        Assert.assertEquals(field.getResult(), "10");

        // The field does not update in real time and will have to be manually updated in Microsoft Word also
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.EDITTIME.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.EDITTIME.docx");

        Assert.assertEquals(10, doc.getBuiltInDocumentProperties().getTotalEditingTime());

        TestUtil.verifyField(FieldType.FIELD_EDIT_TIME, " EDITTIME ", "10", doc.getRange().getFields().get(0));
    }

    //ExStart
    //ExFor:FieldEQ
    //ExSummary:Shows how to use the EQ field to display a variety of mathematical equations.
    @Test //ExSkip
    public void fieldEQ() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // An EQ field displays a mathematical equation consisting of one or many elements
        // Each element takes the following form: [switch][options][arguments]
        // One switch, several possible options, followed by a set of argument values inside round braces

        // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction"
        // No options are invoked, and the values 1 and 4 are passed as arguments
        // This field will display a fraction with 1 as the numerator and 4 as the denominator
        FieldEQ field = insertFieldEQ(builder, "\\f(1,4)");

        Assert.assertEquals(" EQ \\f(1,4)", field.getFieldCode());

        // One EQ field may contain multiple elements placed sequentially,
        // and elements may also be nested by being placed inside the argument brackets of other elements
        // The full list of switches and their corresponding options can be found here:
        // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

        // Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing
        insertFieldEQ(builder, "\\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)");

        // Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces
        // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output
        insertFieldEQ(builder, "\\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))");

        // Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline
        insertFieldEQ(builder, "A \\d \\fo30 \\li() B");

        // Formula consisting of multiple fractions
        insertFieldEQ(builder, "\\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)");

        // Integral switch "\i", with a summation symbol
        insertFieldEQ(builder, "\\i \\su(n=1,5,n)");

        // List switch "\l"
        insertFieldEQ(builder, "\\l(1,1,2,3,n,8,13)");

        // Radical switch "\r", displaying a cubed root of x
        insertFieldEQ(builder, "\\r (3,x)");

        // Subscript/superscript switch "/s", first as a superscript and then as a subscript
        insertFieldEQ(builder, "\\s \\up8(Superscript) Text \\s \\do8(Subscript)");

        // Box switch "\x", with lines at the top, bottom, left and right of the input
        insertFieldEQ(builder, "\\x \\to \\bo \\le \\ri(5)");

        // More complex combinations
        insertFieldEQ(builder, "\\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))");
        insertFieldEQ(builder, "\\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)");
        insertFieldEQ(builder, "\\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)");

        doc.save(getArtifactsDir() + "Field.EQ.docx");
        testFieldEQ(new Document(getArtifactsDir() + "Field.EQ.docx")); //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
    /// </summary>
    private static FieldEQ insertFieldEQ(DocumentBuilder builder, String args) throws Exception {
        FieldEQ field = (FieldEQ) builder.insertField(FieldType.FIELD_EQUATION, true);
        builder.moveTo(field.getSeparator());
        builder.write(args);
        builder.moveTo(field.getStart().getParentNode());

        builder.insertParagraph();
        return field;
    }
    //ExEnd

    private void testFieldEQ(Document doc) {
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\f(1,4)", "", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)", "", doc.getRange().getFields().get(1));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))", "", doc.getRange().getFields().get(2));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ A \\d \\fo30 \\li() B", "", doc.getRange().getFields().get(3));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)", "", doc.getRange().getFields().get(4));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\i \\su(n=1,5,n)", "", doc.getRange().getFields().get(5));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\l(1,1,2,3,n,8,13)", "", doc.getRange().getFields().get(6));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\r (3,x)", "", doc.getRange().getFields().get(7));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\s \\up8(Superscript) Text \\s \\do8(Subscript)", "", doc.getRange().getFields().get(8));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\x \\to \\bo \\le \\ri(5)", "", doc.getRange().getFields().get(9));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))", "", doc.getRange().getFields().get(10));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)", "", doc.getRange().getFields().get(11));
        TestUtil.verifyField(FieldType.FIELD_EQUATION, " EQ \\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)", "", doc.getRange().getFields().get(12));
    }

    @Test
    public void fieldForms() throws Exception {
        //ExStart
        //ExFor:FieldFormCheckBox
        //ExFor:FieldFormDropDown
        //ExFor:FieldFormText
        //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
        // These fields are legacy equivalents of the FormField, and they can be read but not inserted by Aspose.Words,
        // and can be inserted in Microsoft Word 2019 via the Legacy Tools menu in the Developer tab
        Document doc = new Document(getMyDir() + "Form fields.docx");

        FieldFormCheckBox fieldFormCheckBox = (FieldFormCheckBox) doc.getRange().getFields().get(1);
        Assert.assertEquals(" FORMCHECKBOX \u0001", fieldFormCheckBox.getFieldCode());

        FieldFormDropDown fieldFormDropDown = (FieldFormDropDown) doc.getRange().getFields().get(2);
        Assert.assertEquals(" FORMDROPDOWN \u0001", fieldFormDropDown.getFieldCode());

        FieldFormText fieldFormText = (FieldFormText) doc.getRange().getFields().get(0);
        Assert.assertEquals(" FORMTEXT \u0001", fieldFormText.getFieldCode());
        //ExEnd
    }

    @Test
    public void fieldFormula() throws Exception {
        //ExStart
        //ExFor:FieldFormula
        //ExSummary:Shows how to use the "=" field.
        Document doc = new Document();

        // Create a formula field using a field builder
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_FORMULA);
        fieldBuilder.addArgument(2);
        fieldBuilder.addArgument("*");
        fieldBuilder.addArgument(5);

        FieldFormula field = (FieldFormula) fieldBuilder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph());
        field.update();

        Assert.assertEquals(field.getFieldCode(), " = 2 * 5 ");
        Assert.assertEquals(field.getResult(), "10");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FORMULA.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.FORMULA.docx");

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2 * 5 ", "10", doc.getRange().getFields().get(0));
    }

    @Test
    public void fieldLastSavedBy() throws Exception {
        //ExStart
        //ExFor:FieldLastSavedBy
        //ExSummary:Shows how to use the LASTSAVEDBY field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" property
        // This is the property that a LASTSAVEDBY field looks up and displays
        // If we make a document programmatically, this property is null and needs to have a value assigned to it first
        doc.getBuiltInDocumentProperties().setLastSavedBy("John Doe");

        // Insert a LASTSAVEDBY field using a document builder
        FieldLastSavedBy field = (FieldLastSavedBy) builder.insertField(FieldType.FIELD_LAST_SAVED_BY, true);

        // The value from our document property appears here
        Assert.assertEquals(field.getFieldCode(), " LASTSAVEDBY ");
        Assert.assertEquals(field.getResult(), "John Doe");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LASTSAVEDBY.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.LASTSAVEDBY.docx");

        Assert.assertEquals("John Doe", doc.getBuiltInDocumentProperties().getLastSavedBy());
        TestUtil.verifyField(FieldType.FIELD_LAST_SAVED_BY, " LASTSAVEDBY ", "John Doe", doc.getRange().getFields().get(0));
    }

    @Test(enabled = false, description = "WORDSNET-18173")
    public void fieldMergeRec() throws Exception {
        //ExStart
        //ExFor:FieldMergeRec
        //ExFor:FieldMergeSeq
        //ExSummary:Shows how to number and count mail merge records in the output documents of a mail merge using MERGEREC and MERGESEQ fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a merge field
        builder.write("Dear ");
        FieldMergeField fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Name");
        builder.writeln(",");

        // A MERGEREC field will print the row number of the data being merged
        builder.write("\nRow number of record in data source: ");
        FieldMergeRec fieldMergeRec = (FieldMergeRec) builder.insertField(FieldType.FIELD_MERGE_REC, true);

        Assert.assertEquals(fieldMergeRec.getFieldCode(), " MERGEREC ");

        // A MERGESEQ field will count the number of successful merges and print the current value on each respective page
        // If no rows are skipped and the data source is not sorted, and no SKIP/SKIPIF/NEXT/NEXTIF fields are invoked,
        // the MERGESEQ and MERGEREC fields will function the same
        builder.write("\nSuccessful merge number: ");
        FieldMergeSeq fieldMergeSeq = (FieldMergeSeq) builder.insertField(FieldType.FIELD_MERGE_SEQ, true);

        Assert.assertEquals(fieldMergeSeq.getFieldCode(), " MERGESEQ ");

        // Insert a SKIPIF field, which will skip a merge if the name is "John Doe"
        FieldSkipIf fieldSkipIf = (FieldSkipIf) builder.insertField(FieldType.FIELD_SKIP_IF, true);
        builder.moveTo(fieldSkipIf.getSeparator());
        fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Name");
        fieldSkipIf.setLeftExpression("=");
        fieldSkipIf.setRightExpression("John Doe");

        // Create a data source with 3 rows, one of them having "John Doe" as a value for the "Name" column
        // Since a SKIPIF field will be triggered once by that value, the output of our mail merge will have 2 pages instead of 3
        // On page 1, the MERGESEQ and MERGEREC fields will both display "1"
        // On page 2, the MERGEREC field will display "3" and the MERGESEQ field will display "2"
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Name");
        table.getRows().add(new String[][]{{"Jane Doe"}, {"John Doe"}, {"Joe Bloggs"}});

        // Execute the mail merge and save document
        doc.getMailMerge().execute(table);
        doc.save(getArtifactsDir() + "Field.MERGEREC.MERGESEQ.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEREC.MERGESEQ.docx");

        Assert.assertEquals(0, doc.getRange().getFields().getCount());

        Assert.assertEquals("Dear Jane Doe,\r" +
                "\r" +
                "Row number of record in data source: 1\r" +
                "Successful merge number: 1\fDear Joe Bloggs,\r" +
                "\r" +
                "Row number of record in data source: 2\r" +
                "Successful merge number: 3", doc.getText().trim());
    }

    @Test
    public void fieldOcx() throws Exception {
        //ExStart
        //ExFor:FieldOcx
        //ExSummary:Shows how to insert an OCX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert an OCX field
        FieldOcx field = (FieldOcx) builder.insertField(FieldType.FIELD_OCX, true);

        Assert.assertEquals(field.getFieldCode(), " OCX ");
        //ExEnd

        TestUtil.verifyField(FieldType.FIELD_OCX, " OCX ", "", field);
    }

    //ExStart
    //ExFor:FieldPrivate
    //ExSummary:Shows how to process PRIVATE fields.
    @Test //ExSkip
    public void fieldPrivate() throws Exception {
        // Open a Corel WordPerfect document that was converted to .docx format
        Document doc = new Document(getMyDir() + "Field sample - PRIVATE.docx");

        // WordPerfect 5.x/6.x documents like the one we opened may contain PRIVATE fields
        // The PRIVATE field is a WordPerfect artifact that is preserved when a file is opened and saved in Microsoft Word
        // However, they have no functionality in Microsoft Word
        FieldPrivate field = (FieldPrivate) doc.getRange().getFields().get(0);

        Assert.assertEquals(" PRIVATE \"My value\" ", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_PRIVATE, field.getType());

        // PRIVATE fields can also be inserted by a document builder
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField(FieldType.FIELD_PRIVATE, true);

        // It is strongly advised against using them to attempt to hide or store private information
        // Unless backward compatibility with older versions of WordPerfect is necessary, these fields can safely be removed
        // This can be done using a document visitor implementation
        Assert.assertEquals(doc.getRange().getFields().getCount(), 2);

        FieldPrivateRemover remover = new FieldPrivateRemover();
        doc.accept(remover);

        Assert.assertEquals(remover.getFieldsRemovedCount(), 2);
        Assert.assertEquals(doc.getRange().getFields().getCount(), 0);
    }

    /// <summary>
    /// Visitor implementation that removes all PRIVATE fields that it encounters.
    /// </summary>
    public static class FieldPrivateRemover extends DocumentVisitor {
        public FieldPrivateRemover() {
            mFieldsRemovedCount = 0;
        }

        public int getFieldsRemovedCount() {
            return mFieldsRemovedCount;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// If the node belongs to a PRIVATE field, the entire field is removed.
        /// </summary>
        public int visitFieldEnd(final FieldEnd fieldEnd) throws Exception {
            if (fieldEnd.getFieldType() == FieldType.FIELD_PRIVATE) {
                fieldEnd.getField().remove();
                mFieldsRemovedCount++;
            }

            return VisitorAction.CONTINUE;
        }

        private int mFieldsRemovedCount;
    }
    //ExEnd

    @Test
    public void fieldSection() throws Exception {
        //ExStart
        //ExFor:FieldSection
        //ExFor:FieldSectionPages
        //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to facilitate page numbering separated by sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the document builder to the header that appears across all pages and align to the top right
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        // A SECTION field displays the number of the section it is placed in
        builder.write("Section ");
        FieldSection fieldSection = (FieldSection) builder.insertField(FieldType.FIELD_SECTION, true);

        Assert.assertEquals(fieldSection.getFieldCode(), " SECTION ");

        // A PAGE field displays the number of the page it is placed in
        builder.write("\nPage ");
        FieldPage fieldPage = (FieldPage) builder.insertField(FieldType.FIELD_PAGE, true);

        Assert.assertEquals(fieldPage.getFieldCode(), " PAGE ");

        // A SECTIONPAGES field displays the number of pages that the section it is in spans across
        builder.write(" of ");
        FieldSectionPages fieldSectionPages = (FieldSectionPages) builder.insertField(FieldType.FIELD_SECTION_PAGES, true);

        Assert.assertEquals(fieldSectionPages.getFieldCode(), " SECTIONPAGES ");

        // Move out of the header back into the main document and insert two pages
        // Both these pages will be in the first section and our three fields will keep track of the numbers in each header
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // We can insert a new section with the document builder like this
        // This will change the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        // The PAGE field will keep counting pages across the whole document
        // We can manually reset its count after a new section is added to keep track of pages section-by-section
        builder.getCurrentSection().getPageSetup().setRestartPageNumbering(true);
        builder.insertBreak(BreakType.PAGE_BREAK);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SECTION.SECTIONPAGES.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SECTION.SECTIONPAGES.docx");

        TestUtil.verifyField(FieldType.FIELD_SECTION, " SECTION ", "2", doc.getRange().getFields().get(0));
        TestUtil.verifyField(FieldType.FIELD_PAGE, " PAGE ", "2", doc.getRange().getFields().get(1));
        TestUtil.verifyField(FieldType.FIELD_SECTION_PAGES, " SECTIONPAGES ", "2", doc.getRange().getFields().get(2));
    }

    //ExStart
    //ExFor:FieldTime
    //ExSummary:Shows how to display the current time using the TIME field.
    @Test //ExSkip
    public void fieldTime() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, time is displayed in the "h:mm am/pm" format
        FieldTime field = insertFieldTime(builder, "");
        Assert.assertEquals(field.getFieldCode(), " TIME ");

        // By using the \@ flag, we can change the appearance of our time
        field = insertFieldTime(builder, "\\@ HHmm");
        Assert.assertEquals(field.getFieldCode(), " TIME \\@ HHmm");

        // We can even display the date, according to the gregorian calendar
        field = insertFieldTime(builder, "\\@ \"M/d/yyyy h mm:ss am/pm\"");
        Assert.assertEquals(field.getFieldCode(), " TIME \\@ \"M/d/yyyy h mm:ss am/pm\"");

        doc.save(getArtifactsDir() + "Field.TIME.docx");
        testFieldTime(new Document(getArtifactsDir() + "Field.TIME.docx")); //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field
    /// </summary>
    private FieldTime insertFieldTime(final DocumentBuilder builder, final String format) throws Exception {
        FieldTime field = (FieldTime) builder.insertField(FieldType.FIELD_TIME, true);
        builder.moveTo(field.getSeparator());
        builder.write(format);
        builder.moveTo(field.getStart().getParentNode());

        builder.insertParagraph();
        return field;
    }
    //ExEnd

    private void testFieldTime(Document doc) throws Exception {
        String docLoadingTime = LocalTime.now().format(DateTimeFormatter.ofPattern("h:mm a"));
        doc = DocumentHelper.saveOpen(doc);

        FieldTime field = (FieldTime) doc.getRange().getFields().get(0);

        Assert.assertEquals(" TIME ", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_TIME, field.getType());
        Assert.assertEquals(field.getResult(), docLoadingTime);

        field = (FieldTime) doc.getRange().getFields().get(1);

        Assert.assertEquals(" TIME \\@ HHmm", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_TIME, field.getType());
        Assert.assertEquals(field.getResult(), docLoadingTime);

        field = (FieldTime) doc.getRange().getFields().get(2);

        Assert.assertEquals(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_TIME, field.getType());
        Assert.assertEquals(field.getResult(), docLoadingTime);
    }

    @Test
    public void bidiOutline() throws Exception {
        //ExStart
        //ExFor:FieldBidiOutline
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExFor:ParagraphFormat.Bidi
        //ExSummary:Shows how to create RTL lists with BIDIOUTLINE fields.
        // Create a blank document and a document builder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use our builder to insert a BIDIOUTLINE field
        // This field numbers paragraphs like the AUTONUM/LISTNUM fields,
        // but is only visible when a RTL editing language is enabled, such as Hebrew or Arabic
        // The following field will display ".1", the RTL equivalent of list number "1."
        FieldBidiOutline field = (FieldBidiOutline) builder.insertField(FieldType.FIELD_BIDI_OUTLINE, true);
        builder.writeln("שלום");

        Assert.assertEquals(" BIDIOUTLINE ", field.getFieldCode());

        // Add two more BIDIOUTLINE fields, which will be automatically numbered ".2" and ".3"
        builder.insertField(FieldType.FIELD_BIDI_OUTLINE, true);
        builder.writeln("שלום");
        builder.insertField(FieldType.FIELD_BIDI_OUTLINE, true);
        builder.writeln("שלום");

        // Set the horizontal text alignment for every paragraph in the document to RTL
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            para.getParagraphFormat().setBidi(true);
        }

        // If a RTL editing language is enabled in Microsoft Word, out fields will display numbers
        // Otherwise, they will appear as "###"
        doc.save(getArtifactsDir() + "Field.BIDIOUTLINE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.BIDIOUTLINE.docx");

        for (Field fieldBidiOutline : doc.getRange().getFields())
            TestUtil.verifyField(FieldType.FIELD_BIDI_OUTLINE, " BIDIOUTLINE ", "", fieldBidiOutline);
    }

    @Test
    public void legacy() throws Exception {
        //ExStart
        //ExFor:FieldEmbed
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled.
        // Open a document that was created in Microsoft Word 2003
        Document doc = new Document(getMyDir() + "Legacy fields.doc");

        // If we open the document in Word and press Alt+F9, we will see a SHAPE and an EMBED field
        // A SHAPE field is the anchor/canvas for an autoshape object with the "In line with text" wrapping style enabled
        // An EMBED field has the same function, but for an embedded object, such as a spreadsheet from an external Excel document
        // However, these fields will not appear in the document's Fields collection
        Assert.assertEquals(doc.getRange().getFields().getCount(), 0);

        // These fields are supported only by old versions of Microsoft Word
        // As such, they are converted into shapes during the document importation process and can instead be found in the collection of Shape nodes
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        Assert.assertEquals(shapes.getCount(), 3);

        // The first Shape node corresponds to what was the SHAPE field in the input document: the inline canvas for an autoshape
        Shape shape = (Shape) shapes.get(0);
        Assert.assertEquals(shape.getShapeType(), ShapeType.IMAGE);

        // The next Shape node is the autoshape that is within the canvas
        shape = (Shape) shapes.get(1);
        Assert.assertEquals(shape.getShapeType(), ShapeType.CAN);

        // The third Shape is what was the EMBED field that contained the external spreadsheet
        shape = (Shape) shapes.get(2);
        Assert.assertEquals(ShapeType.OLE_OBJECT, shape.getShapeType());
        //ExEnd
    }
}
