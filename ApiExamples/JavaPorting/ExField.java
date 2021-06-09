// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.msString;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldDate;
import com.aspose.words.FieldType;
import com.aspose.words.FieldChar;
import org.testng.Assert;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.FieldIf;
import com.aspose.words.FieldAuthor;
import com.aspose.words.FieldBuilder;
import com.aspose.words.FieldRevNum;
import com.aspose.words.Field;
import com.aspose.words.FieldUnknown;
import com.aspose.words.FindReplaceOptions;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.ms.System.msConsole;
import com.aspose.words.FieldUpdateCultureSource;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.Run;
import com.aspose.words.FieldArgumentBuilder;
import com.aspose.barcode.License;
import com.aspose.words.FieldDatabase;
import com.aspose.words.Table;
import com.aspose.words.NodeType;
import com.aspose.words.FieldIncludePicture;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.FieldFormat;
import com.aspose.words.GeneralFormat;
import java.util.Iterator;
import com.aspose.words.Section;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.Paragraph;
import com.aspose.words.ControlChar;
import com.aspose.words.FieldStart;
import com.aspose.words.CompositeNode;
import com.aspose.words.FieldRef;
import com.aspose.words.FieldAsk;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.FieldMergeField;
import com.aspose.words.IFieldUserPromptRespondent;
import com.aspose.words.FieldAdvance;
import com.aspose.words.FieldAddressBlock;
import com.aspose.words.FieldCollection;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.FieldSeparator;
import com.aspose.words.FieldEnd;
import com.aspose.words.FieldCompare;
import com.aspose.words.FieldIfComparisonResult;
import com.aspose.words.FieldAutoNum;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.FieldAutoNumLgl;
import com.aspose.words.FieldAutoNumOut;
import com.aspose.words.GlossaryDocument;
import com.aspose.words.BuildingBlock;
import com.aspose.words.BuildingBlockGallery;
import com.aspose.words.BuildingBlockBehavior;
import com.aspose.words.FieldAutoText;
import com.aspose.words.FieldGlossary;
import com.aspose.words.FieldAutoTextList;
import com.aspose.words.Body;
import com.aspose.words.FieldGreetingLine;
import com.aspose.words.FieldListNum;
import com.aspose.words.FieldToc;
import com.aspose.words.BreakType;
import com.aspose.words.FieldTC;
import com.aspose.words.FieldSeq;
import com.aspose.words.FieldPageRef;
import com.aspose.words.FieldCitation;
import com.aspose.words.FieldBibliography;
import com.aspose.words.FieldData;
import com.aspose.words.FieldInclude;
import com.aspose.words.FieldImport;
import com.aspose.words.Shape;
import com.aspose.words.FieldIncludeText;
import com.aspose.XmlUtilPal;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.Xml.Schema.XmlNamespaceManager;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.FieldHyperlink;
import com.aspose.words.net.System.Data.DataColumn;
import com.aspose.words.MergeFieldImageDimensionUnit;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.MergeFieldImageDimension;
import com.aspose.words.ImageType;
import java.util.HashMap;
import com.aspose.ms.System.Collections.msDictionary;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import com.aspose.words.FieldIndex;
import com.aspose.words.FieldXE;
import com.aspose.words.FieldBarcode;
import com.aspose.words.FieldDisplayBarcode;
import com.aspose.words.FieldMergeBarcode;
import com.aspose.words.FieldLink;
import com.aspose.words.FieldDde;
import com.aspose.words.FieldDdeAuto;
import com.aspose.words.UserInformation;
import com.aspose.words.FieldUserAddress;
import com.aspose.words.FieldUserInitials;
import com.aspose.words.FieldUserName;
import com.aspose.words.List;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.FieldStyleRef;
import com.aspose.words.FieldCreateDate;
import com.aspose.ms.System.Globalization.Calendar;
import com.aspose.ms.System.Globalization.UmAlQuraCalendar;
import com.aspose.words.FieldSaveDate;
import com.aspose.words.FieldSymbol;
import com.aspose.words.FieldDocProperty;
import com.aspose.words.FieldDocVariable;
import com.aspose.words.FieldSubject;
import com.aspose.words.FieldComments;
import com.aspose.words.FieldFileSize;
import com.aspose.words.FieldGoToButton;
import com.aspose.words.FieldFillIn;
import com.aspose.words.FieldInfo;
import com.aspose.words.FieldMacroButton;
import com.aspose.words.FieldKeywords;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.FieldNumChars;
import com.aspose.words.FieldNumWords;
import com.aspose.words.FieldPage;
import com.aspose.words.FieldNumPages;
import com.aspose.words.FieldPrint;
import com.aspose.words.FieldPrintDate;
import com.aspose.words.FieldQuote;
import com.aspose.words.FieldNext;
import com.aspose.words.FieldNextIf;
import com.aspose.words.FieldNoteRef;
import com.aspose.words.FootnoteType;
import com.aspose.words.FieldFootnoteRef;
import com.aspose.words.Footnote;
import com.aspose.words.FieldRD;
import com.aspose.words.FieldSkipIf;
import com.aspose.words.FieldSet;
import com.aspose.words.FieldTemplate;
import com.aspose.words.FieldTitle;
import com.aspose.words.FieldToa;
import com.aspose.words.FieldTA;
import com.aspose.words.FieldAddIn;
import com.aspose.words.FieldEditTime;
import com.aspose.words.FieldEQ;
import com.aspose.words.FieldFormCheckBox;
import com.aspose.words.FieldFormDropDown;
import com.aspose.words.FieldFormText;
import com.aspose.words.FieldFormula;
import com.aspose.words.FieldLastSavedBy;
import com.aspose.words.FieldMergeRec;
import com.aspose.words.FieldMergeSeq;
import com.aspose.words.FieldOcx;
import com.aspose.words.FieldPrivate;
import com.aspose.words.FieldSection;
import com.aspose.words.FieldSectionPages;
import com.aspose.words.FieldTime;
import com.aspose.words.FieldBidiOutline;
import com.aspose.words.ShapeType;
import com.aspose.words.FieldIndexFormat;
import com.aspose.words.ComparisonEvaluationResult;
import com.aspose.words.IComparisonExpressionEvaluator;
import com.aspose.words.ComparisonExpression;
import java.util.ArrayList;
import org.testng.annotations.DataProvider;


@Test
public class ExField extends ApiExampleBase
{
    @Test
    public void getFieldFromDocument() throws Exception
    {
        //ExStart
        //ExFor:FieldType
        //ExFor:FieldChar
        //ExFor:FieldChar.FieldType
        //ExFor:FieldChar.IsDirty
        //ExFor:FieldChar.IsLocked
        //ExFor:FieldChar.GetField
        //ExFor:Field.IsLocked
        //ExSummary:Shows how to work with a FieldStart node.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldDate field = (FieldDate)builder.insertField(FieldType.FIELD_DATE, true);
        field.getFormat().setDateTimeFormat("dddd, MMMM dd, yyyy");
        field.update();
        
        FieldChar fieldStart = field.getStart();

        Assert.assertEquals(FieldType.FIELD_DATE, fieldStart.getFieldType());
        Assert.assertEquals(false, fieldStart.isDirty());
        Assert.assertEquals(false, fieldStart.isLocked());

        // Retrieve the facade object which represents the field in the document.
        field = (FieldDate)fieldStart.getField();

        Assert.assertEquals(false, field.isLocked());
        Assert.assertEquals(" DATE  \\@ \"dddd, MMMM dd, yyyy\"", field.getFieldCode());

        // Update the field to show the current date.
        field.update();         
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\@ \"dddd, MMMM dd, yyyy\"", new Date().toString("dddd, MMMM dd, yyyy"), doc.getRange().getFields().get(0));
    }
    
    @Test
    public void getFieldCode() throws Exception
    {
        //ExStart
        //ExFor:Field.GetFieldCode
        //ExFor:Field.GetFieldCode(bool)
        //ExSummary:Shows how to get a field's field code.
        // Open a document which contains a MERGEFIELD inside an IF field.
        Document doc = new Document(getMyDir() + "Nested fields.docx");
        FieldIf fieldIf = (FieldIf)doc.getRange().getFields().get(0);

        // There are two ways of getting a field's field code:
        // 1 -  Omit its inner fields:
        Assert.assertEquals(" IF  > 0 \" (surplus of ) \" \"\" ", fieldIf.getFieldCode(false));

        // 2 -  Include its inner fields:
        Assert.assertEquals($" IF \u0013 MERGEFIELD NetIncome \u0014\u0015 > 0 \" (surplus of \u0013 MERGEFIELD  NetIncome \\f $ \u0014\u0015) \" \"\" ",
            fieldIf.getFieldCode(true));

        // By default, the GetFieldCode method displays inner fields.
        Assert.assertEquals(fieldIf.getFieldCode(), fieldIf.getFieldCode(true));
        //ExEnd
    }

    @Test
    public void displayResult() throws Exception
    {
        //ExStart
        //ExFor:Field.DisplayResult
        //ExSummary:Shows how to get the real text that a field displays in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("This document was written by ");
        FieldAuthor fieldAuthor = (FieldAuthor)builder.insertField(FieldType.FIELD_AUTHOR, true);
        fieldAuthor.setAuthorName("John Doe");

        // We can use the DisplayResult property to verify what exact text
        // a field would display in its place in the document.
        Assert.assertEquals("", fieldAuthor.getDisplayResult());

        // Fields do not maintain accurate result values in real-time. 
        // To make sure our fields display accurate results at any given time,
        // such as right before a save operation, we need to update them manually.
        fieldAuthor.update();

        Assert.assertEquals("John Doe", fieldAuthor.getDisplayResult());

        doc.save(getArtifactsDir() + "Field.DisplayResult.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DisplayResult.docx");

        Assert.assertEquals("John Doe", doc.getRange().getFields().get(0).getDisplayResult());
    }

    @Test
    public void createWithFieldBuilder() throws Exception
    {
        //ExStart
        //ExFor:FieldBuilder.#ctor(FieldType)
        //ExFor:FieldBuilder.BuildAndInsert(Inline)
        //ExSummary:Shows how to create and insert a field using a field builder.
        Document doc = new Document();

        // A convenient way of adding text content to a document is with a document builder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write(" Hello world! This text is one Run, which is an inline node.");

        // Fields have their builder, which we can use to construct a field code piece by piece.
        // In this case, we will construct a BARCODE field representing a US postal code,
        // and then insert it in front of a Run.
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_BARCODE);
        fieldBuilder.addArgument("90210");
        fieldBuilder.addSwitch("\\f", "A");
        fieldBuilder.addSwitch("\\u");

        fieldBuilder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0));

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CreateWithFieldBuilder.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.CreateWithFieldBuilder.docx");

        TestUtil.verifyField(FieldType.FIELD_BARCODE, " BARCODE 90210 \\f A \\u ", "", doc.getRange().getFields().get(0));

        Assert.assertEquals(doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(11).getPreviousSibling(), doc.getRange().getFields().get(0).getEnd());
        Assert.assertEquals($"{ControlChar.FieldStartChar} BARCODE 90210 \\f A \\u {ControlChar.FieldEndChar} Hello world! This text is one Run, which is an inline node.", 
            doc.getText().trim());
    }

    @Test
    public void revNum() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.RevisionNumber
        //ExFor:FieldRevNum
        //ExSummary:Shows how to work with REVNUM fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Current revision #");

        // Insert a REVNUM field, which displays the document's current revision number property.
        FieldRevNum field = (FieldRevNum)builder.insertField(FieldType.FIELD_REVISION_NUM, true);

        Assert.assertEquals(" REVNUM ", field.getFieldCode());
        Assert.assertEquals("1", field.getResult());
        Assert.assertEquals(1, doc.getBuiltInDocumentProperties().getRevisionNumber());

        // This property counts how many times a document has been saved in Microsoft Word,
        // and is unrelated to tracked revisions. We can find it by right clicking the document in Windows Explorer
        // via Properties -> Details. We can update this property manually.
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
    public void insertFieldNone() throws Exception
    {
        //ExStart
        //ExFor:FieldUnknown
        //ExSummary:Shows how to work with 'FieldNone' field in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field that does not denote an objective field type in its field code.
        Field field = builder.insertField(" NOTAREALFIELD //a");

        // The "FieldNone" field type is reserved for fields such as these.
        Assert.assertEquals(FieldType.FIELD_NONE, field.getType());

        // We can also still work with these fields and assign them as instances of the FieldUnknown class.
        FieldUnknown fieldUnknown = (FieldUnknown)field;
        Assert.assertEquals(" NOTAREALFIELD //a", fieldUnknown.getFieldCode());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.verifyField(FieldType.FIELD_NONE, " NOTAREALFIELD //a", "Error! Bookmark not defined.", doc.getRange().getFields().get(0));
    }

    @Test
    public void insertTcField() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TC field at the current document builder position.
        builder.insertField("TC \"Entry Text\" \\f t");
    }

    @Test
    public void insertTcFieldsAtText() throws Exception
    {
        Document doc = new Document();

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new InsertTcFieldHandler("Chapter 1", "\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc.getRange().replaceInternal(new Regex("The Beginning"), "", options);
    }

    private static class InsertTcFieldHandler implements IReplacingCallback
    {
        // Store the text and switches to be used for the TC fields.
        private /*final*/ String mFieldText;
        private /*final*/ String mFieldSwitches;

        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty String or null.
        /// </summary>
        public InsertTcFieldHandler(String text, String switches)
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs args) throws Exception
        {
            DocumentBuilder builder = new DocumentBuilder((Document)args.getMatchNode().getDocument());
            builder.moveTo(args.getMatchNode());

            // If the user-specified text is used in the field as display text, use that, otherwise
            // use the match String as the display text.
            String insertText = !msString.isNullOrEmpty(mFieldText) ? mFieldText : args.getMatchInternal().getValue();

            // Insert the TC field before this node using the specified String
            // as the display text and user-defined switches.
            builder.insertField($"TC \"{insertText}\" {mFieldSwitches}");

            return ReplaceAction.SKIP;
        }
    }

    @Test
    public void fieldLocale() throws Exception
    {
        //ExStart
        //ExFor:Field.LocaleId
        //ExSummary:Shows how to insert a field and work with its locale.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a DATE field, and then print the date it will display.
        // Your thread's current culture determines the formatting of the date.
        Field field = builder.insertField("DATE");
        System.out.println("Today's date, as displayed in the \"{CultureInfo.CurrentCulture.EnglishName}\" culture: {field.Result}");

        Assert.assertEquals(1033, field.getLocaleId());
        Assert.assertEquals(FieldUpdateCultureSource.CURRENT_THREAD, doc.getFieldOptions().getFieldUpdateCultureSource()); //ExSkip

        // Changing the culture of our thread will impact the result of the DATE field.
        // Another way to get the DATE field to display a date in a different culture is to use its LocaleId property.
        // This way allows us to avoid changing the thread's culture to get this effect.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        msCultureInfo de = new msCultureInfo("de-DE");
        field.setLocaleId(de.getLCID());
        field.update();

        System.out.println("Today's date, as displayed according to the \"{CultureInfo.GetCultureInfo(field.LocaleId).EnglishName}\" culture: {field.Result}");
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        field = doc.getRange().getFields().get(0); 

        TestUtil.verifyField(FieldType.FIELD_DATE, "DATE", new Date().toString(de.getDateTimeFormat().getShortDatePattern()), field);
        Assert.assertEquals(new msCultureInfo("de-DE").getLCID(), field.getLocaleId());
    }

    @Test (enabled = false, description = "WORDSNET-16037", dataProvider = "updateDirtyFieldsDataProvider")
    public void updateDirtyFields(boolean updateDirtyFields) throws Exception
    {
        //ExStart
        //ExFor:Field.IsDirty
        //ExFor:LoadOptions.UpdateDirtyFields
        //ExSummary:Shows how to use special property for updating field result.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Give the document's built-in "Author" property value, and then display it with a field.
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");
        FieldAuthor field = (FieldAuthor)builder.insertField(FieldType.FIELD_AUTHOR, true);

        Assert.assertFalse(field.isDirty());
        Assert.assertEquals("John Doe", field.getResult());

        // Update the property. The field still displays the old value.
        doc.getBuiltInDocumentProperties().setAuthor("John & Jane Doe");

        Assert.assertEquals("John Doe", field.getResult());

        // Since the field's value is out of date, we can mark it as "dirty".
        // This value will stay out of date until we update the field manually with the Field.Update() method.
        field.isDirty(true);
        
        MemoryStream docStream = new MemoryStream();
        try /*JAVA: was using*/
        {
            // If we save without calling an update method,
            // the field will keep displaying the out of date value in the output document.
            doc.save(docStream, SaveFormat.DOCX);

            // The LoadOptions object has an option to update all fields
            // marked as "dirty" when loading the document.
            LoadOptions options = new LoadOptions();
            options.setUpdateDirtyFields(updateDirtyFields);
            doc = new Document(docStream, options);
            
            Assert.assertEquals("John & Jane Doe", doc.getBuiltInDocumentProperties().getAuthor());

            field = (FieldAuthor)doc.getRange().getFields().get(0);

            // Updating dirty fields like this automatically set their "IsDirty" flag to false.
            if (updateDirtyFields)
            {
                Assert.assertEquals("John & Jane Doe", field.getResult());
                Assert.assertFalse(field.isDirty());
            }
            else
            {
                Assert.assertEquals("John Doe", field.getResult());
                Assert.assertTrue(field.isDirty());
            }
        }
        finally { if (docStream != null) docStream.close(); }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "updateDirtyFieldsDataProvider")
	public static Object[][] updateDirtyFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void insertFieldWithFieldBuilderException() throws Exception
    {
        Document doc = new Document();

        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
        argumentBuilder.addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD));
        argumentBuilder.addNode(run);
        argumentBuilder.addText("Text argument builder");

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_INCLUDE_TEXT);

        Assert.That(
            () => fieldBuilder.addArgument(argumentBuilder).addArgument("=").addArgument("BestField")
                .addArgument(10).addArgument(20.0).buildAndInsert(run), Throws.<IllegalArgumentException>TypeOf());
    }

    @Test
    public void barCodeWord2Pdf() throws Exception
    {
        Document doc = new Document(getMyDir() + "Field sample - BARCODE.docx");

        doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());

        doc.save(getArtifactsDir() + "Field.BarCodeWord2Pdf.pdf");

        BarCodeReader barCodeReader = barCodeReaderPdf(getArtifactsDir() + "Field.BarCodeWord2Pdf.pdf");
        try /*JAVA: was using*/
        {
            Assert.AreEqual("QR", barCodeReader.FoundBarCodes[0].CodeTypeName);
        }
        finally { if (barCodeReader != null) barCodeReader.close(); }
    }

    private BarCodeReader barCodeReaderPdf(String filename) throws Exception
    {
        // Set license for Aspose.BarCode.
        License licenceBarCode = new License();
        licenceBarCode.setLicense(getLicenseDir() + "Aspose.Total.NET.lic");

        Aspose.Pdf.Facades.PdfExtractor pdfExtractor = new Aspose.Pdf.Facades.PdfExtractor();
        pdfExtractor.BindPdf(filename);

        // Set page range for image extraction.
        pdfExtractor.StartPage = 1;
        pdfExtractor.EndPage = 1;

        pdfExtractor.ExtractImage();

        MemoryStream imageStream = new MemoryStream();
        pdfExtractor.GetNextImage(imageStream);
        imageStream.setPosition(0);

        // Recognize the barcode from the image stream above.
        BarCodeReader barcodeReader = new BarCodeReader(imageStream, DecodeType.QR);

        for (BarCodeResult result : barcodeReader.ReadBarCodes() !!Autoporter error: Undefined expression type )
            msConsole.WriteLine("Codetext found: " + result.CodeText + ", Symbology: " + result.CodeTypeName);

        return barcodeReader;
    }

    @Test (enabled = false, description = "WORDSNET-13854")
    public void fieldDatabase() throws Exception
    {
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
        
        // This DATABASE field will run a query on a database, and display the result in a table.
        FieldDatabase field = (FieldDatabase)builder.insertField(FieldType.FIELD_DATABASE, true);
        field.setFileName(getMyDir() + "Database\\Northwind.mdb");
        field.setConnection("DSN=MS Access Databases");
        field.setQuery("SELECT * FROM [Products]");

        Assert.AreEqual($" DATABASE  \\d \"{DatabaseDir.Replace("\\", "\\\\") + "Northwind.mdb"}\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"", 
            field.getFieldCode());

        // Insert another DATABASE field with a more complex query that sorts all products in descending order by gross sales.
        field = (FieldDatabase)builder.insertField(FieldType.FIELD_DATABASE, true);
        field.setFileName(getMyDir() + "Database\\Northwind.mdb");
        field.setConnection("DSN=MS Access Databases");
        field.setQuery("SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
            "FROM([Products] " +
            "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
            "GROUP BY[Products].ProductName " +
            "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC");

        // These properties have the same function as LIMIT and TOP clauses.
        // Configure them to display only rows 1 to 10 of the query result in the field's table.
        field.setFirstRecord("1");
        field.setLastRecord("10");

        // This property is the index of the format we want to use for our table. The list of table formats is in the "Table AutoFormat..." menu
        // that shows up when we create a DATABASE field in Microsoft Word. Index #10 corresponds to the "Colorful 3" format.
        field.setTableFormat("10");

        // The FormatAttribute property is a string representation of an integer which stores multiple flags.
        // We can patrially apply the format which the TableFormat property points to by setting different flags in this property.
        // The number we use is the sum of a combination of values corresponding to different aspects of the table style.
        // 63 represents 1 (borders) + 2 (shading) + 4 (font) + 8 (color) + 16 (autofit) + 32 (heading rows).
        field.setFormatAttributes("63");
        field.setInsertHeadings(true);
        field.setInsertOnceOnMailMerge(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.DATABASE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DATABASE.docx");

        Assert.assertEquals(2, doc.getRange().getFields().getCount());
        
        Table table = doc.getFirstSection().getBody().getTables().get(0);

        Assert.assertEquals(77, table.getRows().getCount());
        Assert.assertEquals(10, table.getRows().get(0).getCells().getCount());

        field = (FieldDatabase)doc.getRange().getFields().get(0);

        Assert.AreEqual($" DATABASE  \\d \"{DatabaseDir.Replace("\\", "\\\\") + "Northwind.mdb"}\" \\c \"DSN=MS Access Databases\" \\s \"SELECT * FROM [Products]\"",
            field.getFieldCode());

        TestUtil.tableMatchesQueryResult(table, getDatabaseDir() + "Northwind.mdb", field.getQuery());

        table = (Table)doc.getChild(NodeType.TABLE, 1, true);
        field = (FieldDatabase)doc.getRange().getFields().get(1);

        Assert.assertEquals(11, table.getRows().getCount());
        Assert.assertEquals(2, table.getRows().get(0).getCells().getCount());
        Assert.assertEquals("ProductName\u0007", table.getRows().get(0).getCells().get(0).getText());
        Assert.assertEquals("GrossSales\u0007", table.getRows().get(0).getCells().get(1).getText());

        Assert.AreEqual($" DATABASE  \\d \"{DatabaseDir.Replace("\\", "\\\\") + "Northwind.mdb"}\" \\c \"DSN=MS Access Databases\" " +
                        $"\\s \"SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales " +
                        "FROM([Products] " +
                        "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) " +
                        "GROUP BY[Products].ProductName " +
                        "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC\" \\f 1 \\t 10 \\l 10 \\b 63 \\h \\o",
            field.getFieldCode());

        table.getRows().get(0).remove();

        TestUtil.tableMatchesQueryResult(table, getDatabaseDir() + "Northwind.mdb", msString.insert(field.getQuery(), 7, " TOP 10 "));
    }

    @Test (dataProvider = "preserveIncludePictureDataProvider")
    public void preserveIncludePicture(boolean preserveIncludePictureField) throws Exception
    {
        //ExStart
        //ExFor:Field.Update(bool)
        //ExFor:LoadOptions.PreserveIncludePictureField
        //ExSummary:Shows how to preserve or discard INCLUDEPICTURE fields when loading a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldIncludePicture includePicture = (FieldIncludePicture)builder.insertField(FieldType.FIELD_INCLUDE_PICTURE, true);
        includePicture.setSourceFullName(getImageDir() + "Transparent background logo.png");
        includePicture.update(true);

        MemoryStream docStream = new MemoryStream();
        try /*JAVA: was using*/
        {
            doc.save(docStream, new OoxmlSaveOptions(SaveFormat.DOCX));

            // We can set a flag in a LoadOptions object to decide whether to convert all INCLUDEPICTURE fields
            // into image shapes when loading a document that contains them.
            LoadOptions loadOptions = new LoadOptions();
            {
                loadOptions.setPreserveIncludePictureField(preserveIncludePictureField);
            }

            doc = new Document(docStream, loadOptions);

            if (preserveIncludePictureField)
            {
                Assert.True(doc.getRange().getFields().Any(f => f.Type == FieldType.FieldIncludePicture));

                doc.updateFields();
                doc.save(getArtifactsDir() + "Field.PreserveIncludePicture.docx");
            }
            else
            {
                Assert.False(doc.getRange().getFields().Any(f => f.Type == FieldType.FieldIncludePicture));
            }
        }
        finally { if (docStream != null) docStream.close(); }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "preserveIncludePictureDataProvider")
	public static Object[][] preserveIncludePictureDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void fieldFormat() throws Exception
    {
        //ExStart
        //ExFor:Field.Format
        //ExFor:Field.Update
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
        //ExSummary:Shows how to format field results.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a field that displays a result with no format applied.
        Field field = builder.insertField("= 2 + 3");

        Assert.assertEquals("= 2 + 3", field.getFieldCode());
        Assert.assertEquals("5", field.getResult());

        // We can apply a format to a field's result using the field's properties.
        // Below are three types of formats that we can apply to a field's result.
        // 1 -  Numeric format:
        FieldFormat format = field.getFormat();
        format.setNumericFormat("$###.00");
        field.update();

        Assert.assertEquals("= 2 + 3 \\# $###.00", field.getFieldCode());
        Assert.assertEquals("$  5.00", field.getResult());

        // 2 -  Date/time format:
        field = builder.insertField("DATE");
        format = field.getFormat();
        format.setDateTimeFormat("dddd, MMMM dd, yyyy");
        field.update();

        Assert.assertEquals("DATE \\@ \"dddd, MMMM dd, yyyy\"", field.getFieldCode());
        System.out.println("Today's date, in {format.DateTimeFormat} format:\n\t{field.Result}");

        // 3 -  General format:
        field = builder.insertField("= 25 + 33");
        format = field.getFormat();
        format.getGeneralFormats().add(GeneralFormat.LOWERCASE_ROMAN);
        format.getGeneralFormats().add(GeneralFormat.UPPER);
        field.update();

        int index = 0;
        Iterator<Integer> generalFormatEnumerator = format.getGeneralFormats().iterator();
        try /*JAVA: was using*/
    	{
            while (generalFormatEnumerator.hasNext())
                System.out.println("General format index {index++}: {generalFormatEnumerator.Current}");
    	}
        finally { if (generalFormatEnumerator != null) generalFormatEnumerator.close(); }

        Assert.assertEquals("= 25 + 33 \\* roman \\* Upper", field.getFieldCode());
        Assert.assertEquals("LVIII", field.getResult());
        Assert.assertEquals(2, format.getGeneralFormats().getCount());
        Assert.assertEquals(GeneralFormat.LOWERCASE_ROMAN, format.getGeneralFormats().get(0));

        // We can remove our formats to revert the field's result to its original form.
        format.getGeneralFormats().remove(GeneralFormat.LOWERCASE_ROMAN);
        format.getGeneralFormats().removeAt(0);
        Assert.assertEquals(0, format.getGeneralFormats().getCount());
        field.update();

        Assert.assertEquals("= 25 + 33  ", field.getFieldCode());
        Assert.assertEquals("58", field.getResult());
        Assert.assertEquals(0, format.getGeneralFormats().getCount());
        //ExEnd
    }

    @Test
    public void unlink() throws Exception
    {
        //ExStart
        //ExFor:Document.UnlinkFields
        //ExSummary:Shows how to unlink all fields in the document.
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        doc.unlinkFields();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);

        Assert.assertEquals("Fields.Docx   Элементы указателя не найдены.     1.\r", paraWithFields);
    }

    @Test
    public void unlinkAllFieldsInRange() throws Exception
    {
        //ExStart
        //ExFor:Range.UnlinkFields
        //ExSummary:Shows how to unlink all fields in a range.
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        Section newSection = (Section)doc.getSections().get(0).deepClone(true);
        doc.getSections().add(newSection);

        doc.getSections().get(1).getRange().unlinkFields();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        String secWithFields = DocumentHelper.getSectionText(doc, 1);

        Assert.assertTrue(secWithFields.trim().endsWith(
            "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4."));
    }

    @Test
    public void unlinkSingleField() throws Exception
    {
        //ExStart
        //ExFor:Field.Unlink
        //ExSummary:Shows how to unlink a field.
        Document doc = new Document(getMyDir() + "Linked fields.docx");
        doc.getRange().getFields().get(1).unlink();
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);

        Assert.assertTrue(paraWithFields.trim().endsWith(
            "FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015"));
    }

    @Test
    public void updateTocPageNumbers() throws Exception
    {
        Document doc = new Document(getMyDir() + "Field sample - TOC.docx");

        Node startNode = DocumentHelper.getParagraph(doc, 2);
        Node endNode = null;

        NodeCollection paragraphCollection = doc.getChildNodes(NodeType.PARAGRAPH, true);

        for (Paragraph para : paragraphCollection.<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            for (Run run : para.getRuns().<Run>OfType() !!Autoporter error: Undefined expression type )
            {
                if (run.getText().contains(ControlChar.PAGE_BREAK))
                {
                    endNode = run;
                    break;
                }
            }
        }

        if (startNode != null && endNode != null)
        {
            removeSequence(startNode, endNode);

            startNode.remove();
            endNode.remove();
        }

        NodeCollection fStart = doc.getChildNodes(NodeType.FIELD_START, true);

        for (FieldStart field : fStart.<FieldStart>OfType() !!Autoporter error: Undefined expression type )
        {
            /*FieldType*/int fType = field.getFieldType();
            if (fType == FieldType.FIELD_TOC)
            {
                Paragraph para = (Paragraph)field.getAncestor(NodeType.PARAGRAPH);
                para.getRange().updateFields();
                break;
            }
        }

        doc.save(getArtifactsDir() + "Field.UpdateTocPageNumbers.docx");
    }

    private static void removeSequence(Node start, Node end)
    {
        Node curNode = start.nextPreOrder(start.getDocument());
        while (curNode != null && !curNode.equals(end))
        {
            Node nextNode = curNode.nextPreOrder(start.getDocument());

            if (curNode.isComposite())
            {
                CompositeNode curComposite = (CompositeNode)curNode;
                if (!curComposite.getChildNodes(NodeType.ANY, true).contains(end) &&
                    !curComposite.getChildNodes(NodeType.ANY, true).contains(start))
                {
                    nextNode = curNode.getNextSibling();
                    curNode.remove();
                }
            }
            else
            {
                curNode.remove();
            }

            curNode = nextNode;
        }
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
    //ExSummary:Shows how to create an ASK field, and set its properties.
    @Test
    public void fieldAsk() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Place a field where the response to our ASK field will be placed.
        FieldRef fieldRef = (FieldRef)builder.insertField(FieldType.FIELD_REF, true);
        fieldRef.setBookmarkName("MyAskField");
        builder.writeln();

        Assert.assertEquals(" REF  MyAskField", fieldRef.getFieldCode());

        // Insert the ASK field and edit its properties to reference our REF field by bookmark name.
        FieldAsk fieldAsk = (FieldAsk)builder.insertField(FieldType.FIELD_ASK, true);
        fieldAsk.setBookmarkName("MyAskField");
        fieldAsk.setPromptText("Please provide a response for this ASK field");
        fieldAsk.setDefaultResponse("Response from within the field.");
        fieldAsk.setPromptOnceOnMailMerge(true);
        builder.writeln();

        Assert.assertEquals(
            " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o",
            fieldAsk.getFieldCode());

        // ASK fields apply the default response to their respective REF fields during a mail merge.
        DataTable table = new DataTable("My Table");
        table.getColumns().add("Column 1");
        table.getRows().add("Row 1");
        table.getRows().add("Row 2");

        FieldMergeField fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Column 1");

        // We can modify or override the default response in our ASK fields with a custom prompt responder,
        // which will occur during a mail merge.
        doc.getFieldOptions().setUserPromptRespondent(new MyPromptRespondent());
        doc.getMailMerge().execute(table);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.ASK.docx");
        testFieldAsk(table, doc); //ExSkip
    }

    /// <summary>
    /// Prepends text to the default response of an ASK field during a mail merge.
    /// </summary>
    private static class MyPromptRespondent implements IFieldUserPromptRespondent
    {
        public String respond(String promptText, String defaultResponse)
        {
            return "Response from MyPromptRespondent. " + defaultResponse;
        }
    }
    //ExEnd

    private void testFieldAsk(DataTable dataTable, Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        FieldRef fieldRef = (FieldRef)doc.getRange().getFields().First(f => f.Type == FieldType.FieldRef);
        TestUtil.verifyField(FieldType.FIELD_REF, 
            " REF  MyAskField", "Response from MyPromptRespondent. Response from within the field.", fieldRef);

        FieldAsk fieldAsk = (FieldAsk)doc.getRange().getFields().First(f => f.Type == FieldType.FieldAsk);
        TestUtil.verifyField(FieldType.FIELD_ASK, 
            " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o", 
            "Response from MyPromptRespondent. Response from within the field.", fieldAsk);
        
        Assert.assertEquals("MyAskField", fieldAsk.getBookmarkName());
        Assert.assertEquals("Please provide a response for this ASK field", fieldAsk.getPromptText());
        Assert.assertEquals("Response from within the field.", fieldAsk.getDefaultResponse());
        Assert.assertEquals(true, fieldAsk.getPromptOnceOnMailMerge());

        TestUtil.mailMergeMatchesDataTable(dataTable, doc, true);
    }

    @Test
    public void fieldAdvance() throws Exception
    {
        //ExStart
        //ExFor:Fields.FieldAdvance
        //ExFor:Fields.FieldAdvance.DownOffset
        //ExFor:Fields.FieldAdvance.HorizontalPosition
        //ExFor:Fields.FieldAdvance.LeftOffset
        //ExFor:Fields.FieldAdvance.RightOffset
        //ExFor:Fields.FieldAdvance.UpOffset
        //ExFor:Fields.FieldAdvance.VerticalPosition
        //ExSummary:Shows how to insert an ADVANCE field, and edit its properties. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("This text is in its normal place.");

        // Below are two ways of using the ADVANCE field to adjust the position of text that follows it.
        // The effects of an ADVANCE field continue to be applied until the paragraph ends,
        // or another ADVANCE field updates the offset/coordinate values.
        // 1 -  Specify a directional offset:
        FieldAdvance field = (FieldAdvance)builder.insertField(FieldType.FIELD_ADVANCE, true);
        Assert.assertEquals(FieldType.FIELD_ADVANCE, field.getType()); //ExSkip
        Assert.assertEquals(" ADVANCE ", field.getFieldCode()); //ExSkip
        field.setRightOffset("5");
        field.setUpOffset("5");

        Assert.assertEquals(" ADVANCE  \\r 5 \\u 5", field.getFieldCode());

        builder.write("This text will be moved up and to the right.");
        
        field = (FieldAdvance)builder.insertField(FieldType.FIELD_ADVANCE, true);
        field.setDownOffset("5");
        field.setLeftOffset("100");

        Assert.assertEquals(" ADVANCE  \\d 5 \\l 100", field.getFieldCode());

        builder.writeln("This text is moved down and to the left, overlapping the previous text.");

        // 2 -  Move text to a position specified by coordinates:
        field = (FieldAdvance)builder.insertField(FieldType.FIELD_ADVANCE, true);
        field.setHorizontalPosition("-100");
        field.setVerticalPosition("200");

        Assert.assertEquals(" ADVANCE  \\x -100 \\y 200", field.getFieldCode());

        builder.write("This text is in a custom position.");

        doc.save(getArtifactsDir() + "Field.ADVANCE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.ADVANCE.docx");

        field = (FieldAdvance)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_ADVANCE, " ADVANCE  \\r 5 \\u 5", "", field);
        Assert.assertEquals("5", field.getRightOffset());
        Assert.assertEquals("5", field.getUpOffset());

        field = (FieldAdvance)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_ADVANCE, " ADVANCE  \\d 5 \\l 100", "", field);
        Assert.assertEquals("5", field.getDownOffset());
        Assert.assertEquals("100", field.getLeftOffset());

        field = (FieldAdvance)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_ADVANCE, " ADVANCE  \\x -100 \\y 200", "", field);
        Assert.assertEquals("-100", field.getHorizontalPosition());
        Assert.assertEquals("200", field.getVerticalPosition());
    }

    @Test
    public void fieldAddressBlock() throws Exception
    {
        //ExStart
        //ExFor:Fields.FieldAddressBlock.ExcludedCountryOrRegionName
        //ExFor:Fields.FieldAddressBlock.FormatAddressOnCountryOrRegion
        //ExFor:Fields.FieldAddressBlock.IncludeCountryOrRegionName
        //ExFor:Fields.FieldAddressBlock.LanguageId
        //ExFor:Fields.FieldAddressBlock.NameAndAddressFormat
        //ExSummary:Shows how to insert an ADDRESSBLOCK field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldAddressBlock field = (FieldAddressBlock)builder.insertField(FieldType.FIELD_ADDRESS_BLOCK, true);

        Assert.assertEquals(" ADDRESSBLOCK ", field.getFieldCode());

        // Setting this to "2" will include all countries and regions,
        // unless it is the one specified in the ExcludedCountryOrRegionName property.
        field.setIncludeCountryOrRegionName("2");
        field.setFormatAddressOnCountryOrRegion(true);
        field.setExcludedCountryOrRegionName("United States");
        field.setNameAndAddressFormat("<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");

        // By default, this property will contain the language ID of the first character of the document.
        // We can set a different culture for the field to format the result with like this.
        field.setLanguageId(Integer.toString(new msCultureInfo("en-US").getLCID()));

        Assert.assertEquals(
            " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033",
            field.getFieldCode());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);
        field = (FieldAddressBlock)doc.getRange().getFields().get(0);

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
    //ExFor:FieldCollection.Count
    //ExFor:FieldCollection.GetEnumerator
    //ExFor:FieldStart
    //ExFor:FieldStart.Accept(DocumentVisitor)
    //ExFor:FieldSeparator
    //ExFor:FieldSeparator.Accept(DocumentVisitor)
    //ExFor:FieldEnd
    //ExFor:FieldEnd.Accept(DocumentVisitor)
    //ExFor:FieldEnd.HasSeparator
    //ExFor:Field.End
    //ExFor:Field.Separator
    //ExFor:Field.Start
    //ExSummary:Shows how to work with a collection of fields.
    @Test //ExSkip
    public void fieldCollection() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder.insertField(" TIME ");
        builder.insertField(" REVNUM ");
        builder.insertField(" AUTHOR  \"John Doe\" ");
        builder.insertField(" SUBJECT \"My Subject\" ");
        builder.insertField(" QUOTE \"Hello world!\" ");
        doc.updateFields();

        FieldCollection fields = doc.getRange().getFields();

        Assert.assertEquals(6, fields.getCount());

        // Iterate over the field collection, and print contents and type
        // of every field using a custom visitor implementation.
        FieldVisitor fieldVisitor = new FieldVisitor();

        Iterator<Field> fieldEnumerator = fields.iterator();
        try /*JAVA: was using*/
        {
            while (fieldEnumerator.hasNext())
            {
                if (fieldEnumerator.next() != null)
                {
                    fieldEnumerator.next().getStart().accept(fieldVisitor);
                    fieldEnumerator.next().getSeparator()?.Accept(fieldVisitor);
                    fieldEnumerator.next().getEnd().accept(fieldVisitor);
                }
                else
                {
                    System.out.println("There are no fields in the document.");
                }
            }
        }
        finally { if (fieldEnumerator != null) fieldEnumerator.close(); }

        System.out.println(fieldVisitor.getText());
        testFieldCollection(fieldVisitor.getText()); //ExSkip
    }

    /// <summary>
    /// Document visitor implementation that prints field info.
    /// </summary>
    public static class FieldVisitor extends DocumentVisitor
    {
        public FieldVisitor()
        {
            mBuilder = new StringBuilder();
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText()
        {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldStart(FieldStart fieldStart)
        {
            msStringBuilder.appendLine(mBuilder, "Found field: " + fieldStart.getFieldType());
            msStringBuilder.appendLine(mBuilder, "\tField code: " + fieldStart.getField().getFieldCode());
            msStringBuilder.appendLine(mBuilder, "\tDisplayed as: " + fieldStart.getField().getResult());

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldSeparator(FieldSeparator fieldSeparator)
        {
            msStringBuilder.appendLine(mBuilder, "\tFound separator: " + fieldSeparator.getText());

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldEnd(FieldEnd fieldEnd)
        {
            msStringBuilder.appendLine(mBuilder, "End of field: " + fieldEnd.getFieldType());

            return VisitorAction.CONTINUE;
        }

        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testFieldCollection(String fieldVisitorText)
    {
        Assert.assertTrue(fieldVisitorText.contains("Found field: FieldDate"));
        Assert.assertTrue(fieldVisitorText.contains("Found field: FieldTime"));
        Assert.assertTrue(fieldVisitorText.contains("Found field: FieldRevisionNum"));
        Assert.assertTrue(fieldVisitorText.contains("Found field: FieldAuthor"));
        Assert.assertTrue(fieldVisitorText.contains("Found field: FieldSubject"));
        Assert.assertTrue(fieldVisitorText.contains("Found field: FieldQuote"));
    }

    @Test
    public void removeFields() throws Exception
    {
        //ExStart
        //ExFor:FieldCollection
        //ExFor:FieldCollection.Count
        //ExFor:FieldCollection.Clear
        //ExFor:FieldCollection.Item(Int32)
        //ExFor:FieldCollection.Remove(Field)
        //ExFor:FieldCollection.RemoveAt(Int32)
        //ExFor:Field.Remove
        //ExSummary:Shows how to remove fields from a field collection.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" DATE \\@ \"dddd, d MMMM yyyy\" ");
        builder.insertField(" TIME ");
        builder.insertField(" REVNUM ");
        builder.insertField(" AUTHOR  \"John Doe\" ");
        builder.insertField(" SUBJECT \"My Subject\" ");
        builder.insertField(" QUOTE \"Hello world!\" ");
        doc.updateFields();

        FieldCollection fields = doc.getRange().getFields();

        Assert.assertEquals(6, fields.getCount());

        // Below are four ways of removing fields from a field collection.
        // 1 -  Get a field to remove itself:
        fields.get(0).remove();
        Assert.assertEquals(5, fields.getCount());

        // 2 -  Get the collection to remove a field that we pass to its removal method:
        Field lastField = fields.get(3);
        fields.remove(lastField);
        Assert.assertEquals(4, fields.getCount());

        // 3 -  Remove a field from a collection at an index:
        fields.removeAt(2);
        Assert.assertEquals(3, fields.getCount());

        // 4 -  Remove all the fields from the collection at once:
        fields.clear();
        Assert.assertEquals(0, fields.getCount());
        //ExEnd
    }

    @Test
    public void fieldCompare() throws Exception
    {
        //ExStart
        //ExFor:FieldCompare
        //ExFor:FieldCompare.ComparisonOperator
        //ExFor:FieldCompare.LeftExpression
        //ExFor:FieldCompare.RightExpression
        //ExSummary:Shows how to compare expressions using a COMPARE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldCompare field = (FieldCompare)builder.insertField(FieldType.FIELD_COMPARE, true);
        field.setLeftExpression("3");
        field.setComparisonOperator("<");
        field.setRightExpression("2");
        field.update();

        // The COMPARE field displays a "0" or a "1", depending on its statement's truth.
        // The result of this statement is false so that this field will display a "0".
        Assert.assertEquals(" COMPARE  3 < 2", field.getFieldCode());
        Assert.assertEquals("0", field.getResult());

        builder.writeln();

        field = (FieldCompare)builder.insertField(FieldType.FIELD_COMPARE, true);
        field.setLeftExpression("5");
        field.setComparisonOperator("=");
        field.setRightExpression("2 + 3");
        field.update();

        // This field displays a "1" since the statement is true.
        Assert.assertEquals(" COMPARE  5 = \"2 + 3\"", field.getFieldCode());
        Assert.assertEquals("1", field.getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.COMPARE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.COMPARE.docx");

        field = (FieldCompare)doc.getRange().getFields().get(0);
        
        TestUtil.verifyField(FieldType.FIELD_COMPARE, " COMPARE  3 < 2", "0", field);
        Assert.assertEquals("3", field.getLeftExpression());
        Assert.assertEquals("<", field.getComparisonOperator());
        Assert.assertEquals("2", field.getRightExpression());

        field = (FieldCompare)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_COMPARE, " COMPARE  5 = \"2 + 3\"", "1", field);
        Assert.assertEquals("5", field.getLeftExpression());
        Assert.assertEquals("=", field.getComparisonOperator());
        Assert.assertEquals("\"2 + 3\"", field.getRightExpression());
    }

    @Test
    public void fieldIf() throws Exception
    {
        //ExStart
        //ExFor:FieldIf
        //ExFor:FieldIf.ComparisonOperator
        //ExFor:FieldIf.EvaluateCondition
        //ExFor:FieldIf.FalseText
        //ExFor:FieldIf.LeftExpression
        //ExFor:FieldIf.RightExpression
        //ExFor:FieldIf.TrueText
        //ExFor:FieldIfComparisonResult
        //ExSummary:Shows how to insert an IF field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Statement 1: ");
        FieldIf field = (FieldIf)builder.insertField(FieldType.FIELD_IF, true);
        field.setLeftExpression("0");
        field.setComparisonOperator("=");
        field.setRightExpression("1");

        // The IF field will display a string from either its "TrueText" property,
        // or its "FalseText" property, depending on the truth of the statement that we have constructed.
        field.setTrueText("True");
        field.setFalseText("False");
        field.update();

        // In this case, "0 = 1" is incorrect, so the displayed result will be "False".
        Assert.assertEquals(" IF  0 = 1 True False", field.getFieldCode());
        Assert.assertEquals(FieldIfComparisonResult.FALSE, field.evaluateCondition());
        Assert.assertEquals("False", field.getResult());

        builder.write("\nStatement 2: ");
        field = (FieldIf)builder.insertField(FieldType.FIELD_IF, true);
        field.setLeftExpression("5");
        field.setComparisonOperator("=");
        field.setRightExpression("2 + 3");
        field.setTrueText("True");
        field.setFalseText("False");
        field.update();

        // This time the statement is correct, so the displayed result will be "True".
        Assert.assertEquals(" IF  5 = \"2 + 3\" True False", field.getFieldCode());
        Assert.assertEquals(FieldIfComparisonResult.TRUE, field.evaluateCondition());
        Assert.assertEquals("True", field.getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.IF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.IF.docx");
        field = (FieldIf)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_IF, " IF  0 = 1 True False", "False", field);
        Assert.assertEquals("0", field.getLeftExpression());
        Assert.assertEquals("=", field.getComparisonOperator());
        Assert.assertEquals("1", field.getRightExpression());
        Assert.assertEquals("True", field.getTrueText());
        Assert.assertEquals("False", field.getFalseText());

        field = (FieldIf)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_IF, " IF  5 = \"2 + 3\" True False", "True", field);
        Assert.assertEquals("5", field.getLeftExpression());
        Assert.assertEquals("=", field.getComparisonOperator());
        Assert.assertEquals("\"2 + 3\"", field.getRightExpression());
        Assert.assertEquals("True", field.getTrueText());
        Assert.assertEquals("False", field.getFalseText());
    }

    @Test
    public void fieldAutoNum() throws Exception
    {
        //ExStart
        //ExFor:FieldAutoNum
        //ExFor:FieldAutoNum.SeparatorCharacter
        //ExSummary:Shows how to number paragraphs using autonum fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Each AUTONUM field displays the current value of a running count of AUTONUM fields,
        // allowing us to automatically number items like a numbered list.
        // This field will display a number "1.".
        FieldAutoNum field = (FieldAutoNum)builder.insertField(FieldType.FIELD_AUTO_NUM, true);
        builder.writeln("\tParagraph 1.");

        Assert.assertEquals(" AUTONUM ", field.getFieldCode());

        field = (FieldAutoNum)builder.insertField(FieldType.FIELD_AUTO_NUM, true);
        builder.writeln("\tParagraph 2.");

        // The separator character, which appears in the field result immediately after the number,is a full stop by default.
        // If we leave this property null, our second AUTONUM field will display "2." in the document.
        Assert.assertNull(field.getSeparatorCharacter());

        // We can set this property to apply the first character of its string as the new separator character.
        // In this case, our AUTONUM field will now display "2:".
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
    public void fieldAutoNumLgl() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        final String FILLER_TEXT = "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                                  "\nUt enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ";

        // AUTONUMLGL fields display a number that increments at each AUTONUMLGL field within its current heading level.
        // These fields maintain a separate count for each heading level,
        // and each field also displays the AUTONUMLGL field counts for all heading levels below its own. 
        // Changing the count for any heading level resets the counts for all levels above that level to 1.
        // This allows us to organize our document in the form of an outline list.
        // This is the first AUTONUMLGL field at a heading level of 1, displaying "1." in the document.
        insertNumberedClause(builder, "\tHeading 1", FILLER_TEXT, StyleIdentifier.HEADING_1);

        // This is the second AUTONUMLGL field at a heading level of 1, so it will display "2.".
        insertNumberedClause(builder, "\tHeading 2", FILLER_TEXT, StyleIdentifier.HEADING_1);

        // This is the first AUTONUMLGL field at a heading level of 2,
        // and the AUTONUMLGL count for the heading level below it is "2", so it will display "2.1.".
        insertNumberedClause(builder, "\tHeading 3", FILLER_TEXT, StyleIdentifier.HEADING_2);

        // This is the first AUTONUMLGL field at a heading level of 3. 
        // Working in the same way as the field above, it will display "2.1.1.".
        insertNumberedClause(builder, "\tHeading 4", FILLER_TEXT, StyleIdentifier.HEADING_3);

        // This field is at a heading level of 2, and its respective AUTONUMLGL count is at 2, so the field will display "2.2.".
        insertNumberedClause(builder, "\tHeading 5", FILLER_TEXT, StyleIdentifier.HEADING_2);

        // Incrementing the AUTONUMLGL count for a heading level below this one
        // has reset the count for this level so that this field will display "2.2.1.".
        insertNumberedClause(builder, "\tHeading 6", FILLER_TEXT, StyleIdentifier.HEADING_3);

        for (FieldAutoNumLgl field : doc.getRange().getFields().Where(f => f.Type == FieldType.FieldAutoNumLegal) !!Autoporter error: Undefined expression type )
        {
            // The separator character, which appears in the field result immediately after the number,
            // is a full stop by default. If we leave this property null,
            // our last AUTONUMLGL field will display "2.2.1." in the document.
            Assert.assertNull(field.getSeparatorCharacter());

            // Setting a custom separator character and removing the trailing period
            // will change that field's appearance from "2.2.1." to "2:2:1".
            // We will apply this to all the fields that we have created.
            field.setSeparatorCharacter(":");
            field.setRemoveTrailingPeriod(true);
            Assert.assertEquals(" AUTONUMLGL  \\s : \\e", field.getFieldCode());
        }

        doc.save(getArtifactsDir() + "Field.AUTONUMLGL.docx");
        testFieldAutoNumLgl(doc); //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a clause numbered by an AUTONUMLGL field.
    /// </summary>
    private static void insertNumberedClause(DocumentBuilder builder, String heading, String contents, /*StyleIdentifier*/int headingStyle) throws Exception
    {
        builder.insertField(FieldType.FIELD_AUTO_NUM_LEGAL, true);
        builder.getCurrentParagraph().getParagraphFormat().setStyleIdentifier(headingStyle);
        builder.writeln(heading);

        // This text will belong to the auto num legal field above it.
        // It will collapse when we click the arrow next to the corresponding AUTONUMLGL field in Microsoft Word.
        builder.getCurrentParagraph().getParagraphFormat().setStyleIdentifier(StyleIdentifier.BODY_TEXT);
        builder.writeln(contents);
    }
    //ExEnd

    private void testFieldAutoNumLgl(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        for (FieldAutoNumLgl field : doc.getRange().getFields().Where(f => f.Type == FieldType.FieldAutoNumLegal) !!Autoporter error: Undefined expression type )
        {
            TestUtil.verifyField(FieldType.FIELD_AUTO_NUM_LEGAL, " AUTONUMLGL  \\s : \\e", "", field);
            
            Assert.assertEquals(":", field.getSeparatorCharacter());
            Assert.assertTrue(field.getRemoveTrailingPeriod());
        }
    }

    @Test
    public void fieldAutoNumOut() throws Exception
    {
        //ExStart
        //ExFor:FieldAutoNumOut
        //ExSummary:Shows how to number paragraphs using AUTONUMOUT fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // AUTONUMOUT fields display a number that increments at each AUTONUMOUT field.
        // Unlike AUTONUM fields, AUTONUMOUT fields use the outline numbering scheme,
        // which we can define in Microsoft Word via Format -> Bullets & Numbering -> "Outline Numbered".
        // This allows us to automatically number items like a numbered list.
        // LISTNUM fields are a newer alternative to AUTONUMOUT fields.
        // This field will display "1.".
        builder.insertField(FieldType.FIELD_AUTO_NUM_OUTLINE, true);
        builder.writeln("\tParagraph 1.");

        // This field will display "2.".
        builder.insertField(FieldType.FIELD_AUTO_NUM_OUTLINE, true);
        builder.writeln("\tParagraph 2.");

        for (FieldAutoNumOut field : doc.getRange().getFields().Where(f => f.Type == FieldType.FieldAutoNumOutline) !!Autoporter error: Undefined expression type )
            Assert.assertEquals(" AUTONUMOUT ", field.getFieldCode());

        doc.save(getArtifactsDir() + "Field.AUTONUMOUT.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTONUMOUT.docx");

        for (Field field : doc.getRange().getFields())
            TestUtil.verifyField(FieldType.FIELD_AUTO_NUM_OUTLINE, " AUTONUMOUT ", "", field);
    }

    @Test
    public void fieldAutoText() throws Exception
    {
        //ExStart
        //ExFor:Fields.FieldAutoText
        //ExFor:FieldAutoText.EntryName
        //ExFor:FieldOptions.BuiltInTemplatesPaths
        //ExFor:FieldGlossary
        //ExFor:FieldGlossary.EntryName
        //ExSummary:Shows how to display a building block with AUTOTEXT and GLOSSARY fields. 
        Document doc = new Document();

        // Create a glossary document and add an AutoText building block to it.
        doc.setGlossaryDocument(new GlossaryDocument());
        BuildingBlock buildingBlock = new BuildingBlock(doc.getGlossaryDocument());
        buildingBlock.setName("MyBlock");
        buildingBlock.setGallery(BuildingBlockGallery.AUTO_TEXT);
        buildingBlock.setCategory("General");
        buildingBlock.setDescription("MyBlock description");
        buildingBlock.setBehavior(BuildingBlockBehavior.PARAGRAPH);
        doc.getGlossaryDocument().appendChild(buildingBlock);

        // Create a source and add it as text to our building block.
        Document buildingBlockSource = new Document();
        DocumentBuilder buildingBlockSourceBuilder = new DocumentBuilder(buildingBlockSource);
        buildingBlockSourceBuilder.writeln("Hello World!");

        Node buildingBlockContent = doc.getGlossaryDocument().importNode(buildingBlockSource.getFirstSection(), true);
        buildingBlock.appendChild(buildingBlockContent);

        // Set a file which contains parts that our document, or its attached template may not contain.
        doc.getFieldOptions().setBuiltInTemplatesPaths(new String[] { getMyDir() + "Busniess brochure.dotx" });

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways to use fields to display the contents of our building block.
        // 1 -  Using an AUTOTEXT field:
        FieldAutoText fieldAutoText = (FieldAutoText)builder.insertField(FieldType.FIELD_AUTO_TEXT, true);
        fieldAutoText.setEntryName("MyBlock");

        Assert.assertEquals(" AUTOTEXT  MyBlock", fieldAutoText.getFieldCode());
        
        // 2 -  Using a GLOSSARY field:
        FieldGlossary fieldGlossary = (FieldGlossary)builder.insertField(FieldType.FIELD_GLOSSARY, true);
        fieldGlossary.setEntryName("MyBlock");

        Assert.assertEquals(" GLOSSARY  MyBlock", fieldGlossary.getFieldCode());

		doc.updateFields();
        doc.save(getArtifactsDir() + "Field.AUTOTEXT.GLOSSARY.dotx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTOTEXT.GLOSSARY.dotx");
        
        Assert.That(doc.getFieldOptions().getBuiltInTemplatesPaths(), Is.Empty);

        fieldAutoText = (FieldAutoText)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_AUTO_TEXT, " AUTOTEXT  MyBlock", "Hello World!\r", fieldAutoText);
        Assert.assertEquals("MyBlock", fieldAutoText.getEntryName());

        fieldGlossary = (FieldGlossary)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_GLOSSARY, " GLOSSARY  MyBlock", "Hello World!\r", fieldGlossary);
        Assert.assertEquals("MyBlock", fieldGlossary.getEntryName());
    }

    //ExStart
    //ExFor:Fields.FieldAutoTextList
    //ExFor:Fields.FieldAutoTextList.EntryName
    //ExFor:Fields.FieldAutoTextList.ListStyle
    //ExFor:Fields.FieldAutoTextList.ScreenTip
    //ExSummary:Shows how to use an AUTOTEXTLIST field to select from a list of AutoText entries.
    @Test //ExSkip
    public void fieldAutoTextList() throws Exception
    {
        Document doc = new Document();

        // Create a glossary document and populate it with auto text entries.
        doc.setGlossaryDocument(new GlossaryDocument());
        appendAutoTextEntry(doc.getGlossaryDocument(), "AutoText 1", "Contents of AutoText 1");
        appendAutoTextEntry(doc.getGlossaryDocument(), "AutoText 2", "Contents of AutoText 2");
        appendAutoTextEntry(doc.getGlossaryDocument(), "AutoText 3", "Contents of AutoText 3");

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an AUTOTEXTLIST field and set the text that the field will display in Microsoft Word.
        // Set the text to prompt the user to right-click this field to select an AutoText building block,
        // whose contents the field will display.
        FieldAutoTextList field = (FieldAutoTextList)builder.insertField(FieldType.FIELD_AUTO_TEXT_LIST, true);
        field.setEntryName("Right click here to select an AutoText block");
        field.setListStyle("Heading 1");
        field.setScreenTip("Hover tip text for AutoTextList goes here");

        Assert.assertEquals(" AUTOTEXTLIST  \"Right click here to select an AutoText block\" " +
                        "\\s \"Heading 1\" " +
                        "\\t \"Hover tip text for AutoTextList goes here\"", field.getFieldCode());

        doc.save(getArtifactsDir() + "Field.AUTOTEXTLIST.dotx");
        testFieldAutoTextList(doc); //ExSkip
    }

    /// <summary>
    /// Create an AutoText-type building block and add it to a glossary document.
    /// </summary>
    private static void appendAutoTextEntry(GlossaryDocument glossaryDoc, String name, String contents)
    {
        BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc);
        buildingBlock.setName(name);
        buildingBlock.setGallery(BuildingBlockGallery.AUTO_TEXT);
        buildingBlock.setCategory("General");
        buildingBlock.setBehavior(BuildingBlockBehavior.PARAGRAPH);

        Section section = new Section(glossaryDoc);
        section.appendChild(new Body(glossaryDoc));
        section.getBody().appendParagraph(contents);
        buildingBlock.appendChild(section);

        glossaryDoc.appendChild(buildingBlock);
    }
    //ExEnd

    private void testFieldAutoTextList(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(3, doc.getGlossaryDocument().getCount());
        Assert.assertEquals("AutoText 1", doc.getGlossaryDocument().getBuildingBlocks().get(0).getName());
        Assert.assertEquals("Contents of AutoText 1", doc.getGlossaryDocument().getBuildingBlocks().get(0).getText().trim());
        Assert.assertEquals("AutoText 2", doc.getGlossaryDocument().getBuildingBlocks().get(1).getName());
        Assert.assertEquals("Contents of AutoText 2", doc.getGlossaryDocument().getBuildingBlocks().get(1).getText().trim());
        Assert.assertEquals("AutoText 3", doc.getGlossaryDocument().getBuildingBlocks().get(2).getName());
        Assert.assertEquals("Contents of AutoText 3", doc.getGlossaryDocument().getBuildingBlocks().get(2).getText().trim());

        FieldAutoTextList field = (FieldAutoTextList)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_AUTO_TEXT_LIST,
            " AUTOTEXTLIST  \"Right click here to select an AutoText block\" \\s \"Heading 1\" \\t \"Hover tip text for AutoTextList goes here\"",
            "", field);
        Assert.assertEquals("Right click here to select an AutoText block", field.getEntryName());
        Assert.assertEquals("Heading 1", field.getListStyle());
        Assert.assertEquals("Hover tip text for AutoTextList goes here", field.getScreenTip());
    }

    @Test
    public void fieldGreetingLine() throws Exception
    {
        //ExStart
        //ExFor:FieldGreetingLine
        //ExFor:FieldGreetingLine.AlternateText
        //ExFor:FieldGreetingLine.GetFieldNames
        //ExFor:FieldGreetingLine.LanguageId
        //ExFor:FieldGreetingLine.NameFormat
        //ExSummary:Shows how to insert a GREETINGLINE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a generic greeting using a GREETINGLINE field, and some text after it.
        FieldGreetingLine field = (FieldGreetingLine)builder.insertField(FieldType.FIELD_GREETING_LINE, true);
        builder.writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!");

        // A GREETINGLINE field accepts values from a data source during a mail merge, like a MERGEFIELD.
        // It can also format how the source's data is written in its place once the mail merge is complete.
        // The field names collection corresponds to the columns from the data source
        // that the field will take values from.
        Assert.assertEquals(0, field.getFieldNames().length);

        // To populate that array, we need to specify a format for our greeting line.
        field.setNameFormat("<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> ");

        // Now, our field will accept values from these two columns in the data source.
        Assert.assertEquals("Courtesy Title", field.getFieldNames()[0]);
        Assert.assertEquals("Last Name", field.getFieldNames()[1]);
        Assert.assertEquals(2, field.getFieldNames().length);

        // This string will cover any cases where the data table data is invalid
        // by substituting the malformed name with a string.
        field.setAlternateText("Sir or Madam");

        // Set a locale to format the result.
        field.setLanguageId(Integer.toString(new msCultureInfo("en-US").getLCID()));

        Assert.assertEquals(" GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033", 
            field.getFieldCode());

        // Create a data table with columns whose names match elements
        // from the field's field names collection, and then carry out the mail merge.
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");

        // This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text.
        table.getRows().add("", "No", "Name");

        doc.getMailMerge().execute(table);

        Assert.That(doc.getRange().getFields(), Is.Empty);
        Assert.assertEquals("Dear Mr. Doe,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                        "\fDear Mrs. Cardholder,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!\r" +
                        "\fDear Sir or Madam,\r\r\tThis is your custom greeting, created programmatically using Aspose Words!",
            doc.getText().trim());
        //ExEnd
    }

    @Test
    public void fieldListNum() throws Exception
    {
        //ExStart
        //ExFor:FieldListNum
        //ExFor:FieldListNum.HasListName
        //ExFor:FieldListNum.ListLevel
        //ExFor:FieldListNum.ListName
        //ExFor:FieldListNum.StartingNumber
        //ExSummary:Shows how to number paragraphs with LISTNUM fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // LISTNUM fields display a number that increments at each LISTNUM field.
        // These fields also have a variety of options that allow us to use them to emulate numbered lists.
        FieldListNum field = (FieldListNum)builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // Lists start counting at 1 by default, but we can set this number to a different value, such as 0.
        // This field will display "0)".
        field.setStartingNumber("0");
        builder.writeln("Paragraph 1");

        Assert.assertEquals(" LISTNUM  \\s 0", field.getFieldCode());

        // LISTNUM fields maintain separate counts for each list level. 
        // Inserting a LISTNUM field in the same paragraph as another LISTNUM field
        // increases the list level instead of the count.
        // The next field will continue the count we started above and display a value of "1" at list level 1.
        builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // This field will start a count at list level 2. It will display a value of "1".
        builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // This field will start a count at list level 3. It will display a value of "1".
        // Different list levels have different formatting,
        // so these fields combined will display a value of "1)a)i)".
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.writeln("Paragraph 2");

        // The next LISTNUM field that we insert will continue the count at the list level
        // that the previous LISTNUM field was on.
        // We can use the "ListLevel" property to jump to a different list level.
        // If this LISTNUM field stayed on list level 3, it would display "ii)",
        // but, since we have moved it to list level 2, it carries on the count at that level and displays "b)".
        field = (FieldListNum)builder.insertField(FieldType.FIELD_LIST_NUM, true);
        field.setListLevel("2");
        builder.writeln("Paragraph 3");

        Assert.assertEquals(" LISTNUM  \\l 2", field.getFieldCode());

        // We can set the ListName property to get the field to emulate a different AUTONUM field type.
        // "NumberDefault" emulates AUTONUM, "OutlineDefault" emulates AUTONUMOUT,
        // and "LegalDefault" emulates AUTONUMLGL fields.
        // The "OutlineDefault" list name with 1 as the starting number will result in displaying "I.".
        field = (FieldListNum)builder.insertField(FieldType.FIELD_LIST_NUM, true);
        field.setStartingNumber("1");
        field.setListName("OutlineDefault");
        builder.writeln("Paragraph 4");

        Assert.assertTrue(field.hasListName());
        Assert.assertEquals(" LISTNUM  OutlineDefault \\s 1", field.getFieldCode());

        // The ListName does not carry over from the previous field, so we will need to set it for each new field.
        // This field continues the count with the different list name and displays "II.".
        field = (FieldListNum)builder.insertField(FieldType.FIELD_LIST_NUM, true);
        field.setListName("OutlineDefault");
        builder.writeln("Paragraph 5");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LISTNUM.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.LISTNUM.docx");

        Assert.assertEquals(7, doc.getRange().getFields().getCount());

        field = (FieldListNum)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM  \\s 0", "", field);
        Assert.assertEquals("0", field.getStartingNumber());
        Assert.assertNull(field.getListLevel());
        Assert.assertFalse(field.hasListName());
        Assert.assertNull(field.getListName());

        for (int i = 1; i < 4; i++)
        {
            field = (FieldListNum)doc.getRange().getFields().get(i);

            TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM ", "", field);
            Assert.assertNull(field.getStartingNumber());
            Assert.assertNull(field.getListLevel());
            Assert.assertFalse(field.hasListName());
            Assert.assertNull(field.getListName());
        }

        field = (FieldListNum)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM  \\l 2", "", field);
        Assert.assertNull(field.getStartingNumber());
        Assert.assertEquals("2", field.getListLevel());
        Assert.assertFalse(field.hasListName());
        Assert.assertNull(field.getListName());

        field = (FieldListNum)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_LIST_NUM, " LISTNUM  OutlineDefault \\s 1", "", field);
        Assert.assertEquals("1", field.getStartingNumber());
        Assert.assertNull(field.getListLevel());
        Assert.assertTrue(field.hasListName());
        Assert.assertEquals("OutlineDefault", field.getListName());
    }

    @Test
    public void mergeField() throws Exception
    {
        //ExStart
        //ExFor:FieldMergeField
        //ExFor:FieldMergeField.FieldName
        //ExFor:FieldMergeField.FieldNameNoPrefix
        //ExFor:FieldMergeField.IsMapped
        //ExFor:FieldMergeField.IsVerticalFormatting
        //ExFor:FieldMergeField.TextAfter
        //ExFor:FieldMergeField.TextBefore
        //ExSummary:Shows how to use MERGEFIELD fields to perform a mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a data table to be used as a mail merge data source.
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");

        // Insert a MERGEFIELD with a FieldName property set to the name of a column in the data source.
        FieldMergeField fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Courtesy Title");
        fieldMergeField.isMapped(true);
        fieldMergeField.isVerticalFormatting(false);

        // We can apply text before and after the value that this field accepts when the merge takes place.
        fieldMergeField.setTextBefore("Dear ");
        fieldMergeField.setTextAfter(" ");

        Assert.assertEquals(" MERGEFIELD  \"Courtesy Title\" \\m \\b \"Dear \" \\f \" \"", fieldMergeField.getFieldCode());

        // Insert another MERGEFIELD for a different column in the data source.
        fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Last Name");
        fieldMergeField.setTextAfter(":");

        doc.updateFields();
        doc.getMailMerge().execute(table);

        Assert.assertEquals("Dear Mr. Doe:\fDear Mrs. Cardholder:", doc.getText().trim());
        //ExEnd

        Assert.That(doc.getRange().getFields(), Is.Empty);
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
    //ExSummary:Shows how to insert a TOC, and populate it with entries based on heading styles.
    @Test //ExSkip
    public void fieldToc() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");

        // Insert a TOC field, which will compile all headings into a table of contents.
        // For each heading, this field will create a line with the text in that heading style to the left,
        // and the page the heading appears on to the right.
        FieldToc field = (FieldToc)builder.insertField(FieldType.FIELD_TOC, true);

        // Use the BookmarkName property to only list headings
        // that appear within the bounds of a bookmark with the "MyBookmark" name.
        field.setBookmarkName("MyBookmark");

        // Text with a built-in heading style, such as "Heading 1", applied to it will count as a heading.
        // We can name additional styles to be picked up as headings by the TOC in this property and their TOC levels.
        field.setCustomStyles("Quote; 6; Intense Quote; 7");

        // By default, Styles/TOC levels are separated in the CustomStyles property by a comma,
        // but we can set a custom delimiter in this property.
        doc.getFieldOptions().setCustomTocStyleSeparator(";");

        // Configure the field to exclude any headings that have TOC levels outside of this range.
        field.setHeadingLevelRange("1-3");

        // The TOC will not display the page numbers of headings whose TOC levels are within this range.
        field.setPageNumberOmittingLevelRange("2-5");

        // Set a custom string that will separate every heading from its page number. 
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

        // These two headings will have the page numbers omitted because they are within the "2-5" range.
        insertNewPageWithHeading(builder, "Fifth entry", "Heading 2");
        insertNewPageWithHeading(builder, "Sixth entry", "Heading 3");

        // This entry does not appear because "Heading 4" is outside of the "1-3" range that we have set earlier.
        insertNewPageWithHeading(builder, "Seventh entry", "Heading 4");

        builder.endBookmark("MyBookmark");
        builder.writeln("Paragraph text.");

        // This entry does not appear because it is outside the bookmark specified by the TOC.
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
    @Test (enabled = false)
    public void insertNewPageWithHeading(DocumentBuilder builder, String captionText, String styleName)
    {
        builder.insertBreak(BreakType.PAGE_BREAK);
        String originalStyle = builder.getParagraphFormat().getStyleName();
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get(styleName));
        builder.writeln(captionText);
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get(originalStyle));
    }
    //ExEnd

    private void testFieldToc(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);
        FieldToc field = (FieldToc)doc.getRange().getFields().get(0);

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
    //ExSummary:Shows how to insert a TOC field, and filter which TC fields end up as entries.
    @Test //ExSkip
    public void fieldTocEntryIdentifier() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field, which will compile all TC fields into a table of contents.
        FieldToc fieldToc = (FieldToc)builder.insertField(FieldType.FIELD_TOC, true);

        // Configure the field only to pick up TC entries of the "A" type, and an entry-level between 1 and 3.
        fieldToc.setEntryIdentifier("A");
        fieldToc.setEntryLevelRange("1-3");

        Assert.assertEquals(" TOC  \\f A \\l 1-3", fieldToc.getFieldCode());

        // These two entries will appear in the table.
        builder.insertBreak(BreakType.PAGE_BREAK);
        insertTocEntry(builder, "TC field 1", "A", "1");
        insertTocEntry(builder, "TC field 2", "A", "2");

        Assert.assertEquals(" TC  \"TC field 1\" \\n \\f A \\l 1", doc.getRange().getFields().get(1).getFieldCode());

        // This entry will be omitted from the table because it has a different type from "A".
        insertTocEntry(builder, "TC field 3", "B", "1");

        // This entry will be omitted from the table because it has an entry-level outside of the 1-3 range.
        insertTocEntry(builder, "TC field 4", "A", "5");
        
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TC.docx");
        testFieldTocEntryIdentifier(doc); //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert a TC field.
    /// </summary>
    @Test (enabled = false)
    public void insertTocEntry(DocumentBuilder builder, String text, String typeIdentifier, String entryLevel) throws Exception
    {
        FieldTC fieldTc = (FieldTC)builder.insertField(FieldType.FIELD_TOC_ENTRY, true);
        fieldTc.setOmitPageNumber(true);
        fieldTc.setText(text);
        fieldTc.setTypeIdentifier(typeIdentifier);
        fieldTc.setEntryLevel(entryLevel);
    }
    //ExEnd

    private void testFieldTocEntryIdentifier(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);
        FieldToc fieldToc = (FieldToc)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_TOC, " TOC  \\f A \\l 1-3", "TC field 1\rTC field 2\r", fieldToc);
        Assert.assertEquals("A", fieldToc.getEntryIdentifier());
        Assert.assertEquals("1-3", fieldToc.getEntryLevelRange());

        FieldTC fieldTc = (FieldTC)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 1\" \\n \\f A \\l 1", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 1", fieldTc.getText());
        Assert.assertEquals("A", fieldTc.getTypeIdentifier());
        Assert.assertEquals("1", fieldTc.getEntryLevel());

        fieldTc = (FieldTC)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 2\" \\n \\f A \\l 2", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 2", fieldTc.getText());
        Assert.assertEquals("A", fieldTc.getTypeIdentifier());
        Assert.assertEquals("2", fieldTc.getEntryLevel());

        fieldTc = (FieldTC)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 3\" \\n \\f B \\l 1", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 3", fieldTc.getText());
        Assert.assertEquals("B", fieldTc.getTypeIdentifier());
        Assert.assertEquals("1", fieldTc.getEntryLevel());

        fieldTc = (FieldTC)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_TOC_ENTRY, " TC  \"TC field 4\" \\n \\f A \\l 5", "", fieldTc);
        Assert.assertTrue(fieldTc.getOmitPageNumber());
        Assert.assertEquals("TC field 4", fieldTc.getText());
        Assert.assertEquals("A", fieldTc.getTypeIdentifier());
        Assert.assertEquals("5", fieldTc.getEntryLevel());
    }

    @Test
    public void tocSeqPrefix() throws Exception
    {
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

        // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
        // Each entry contains the paragraph that includes the SEQ field and the page's number that the field appears on.
        FieldToc fieldToc = (FieldToc)builder.insertField(FieldType.FIELD_TOC, true);

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Use the "TableOfFiguresLabel" property to name a main sequence for the TOC.
        // Now, this TOC will only create entries out of SEQ fields with their "SequenceIdentifier" set to "MySequence".
        fieldToc.setTableOfFiguresLabel("MySequence");

        // We can name another SEQ field sequence in the "PrefixedSequenceIdentifier" property.
        // SEQ fields from this prefix sequence will not create TOC entries. 
        // Every TOC entry created from a main sequence SEQ field will now also display the count that
        // the prefix sequence is currently on at the primary sequence SEQ field that made the entry.
        fieldToc.setPrefixedSequenceIdentifier("PrefixSequence");

        // Each TOC entry will display the prefix sequence count immediately to the left
        // of the page number that the main sequence SEQ field appears on.
        // We can specify a custom separator that will appear between these two numbers.
        fieldToc.setSequenceSeparator(">");

        Assert.assertEquals(" TOC  \\c MySequence \\s PrefixSequence \\d >", fieldToc.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);

        // There are two ways of using SEQ fields to populate this TOC.
        // 1 -  Inserting a SEQ field that belongs to the TOC's prefix sequence:
        // This field will increment the SEQ sequence count for the "PrefixSequence" by 1.
        // Since this field does not belong to the main sequence identified
        // by the "TableOfFiguresLabel" property of the TOC, it will not appear as an entry.
        FieldSeq fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("PrefixSequence");
        builder.insertParagraph();

        Assert.assertEquals(" SEQ  PrefixSequence", fieldSeq.getFieldCode());

        // 2 -  Inserting a SEQ field that belongs to the TOC's main sequence:
        // This SEQ field will create an entry in the TOC.
        // The TOC entry will contain the paragraph that the SEQ field is in and the number of the page that it appears on.
        // This entry will also display the count that the prefix sequence is currently at,
        // separated from the page number by the value in the TOC's SeqenceSeparator property.
        // The "PrefixSequence" count is at 1, this main sequence SEQ field is on page 2,
        // and the separator is ">", so entry will display "1>2".
        builder.write("First TOC entry, MySequence #");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");

        Assert.assertEquals(" SEQ  MySequence", fieldSeq.getFieldCode());

        // Insert a page, advance the prefix sequence by 2, and insert a SEQ field to create a TOC entry afterwards.
        // The prefix sequence is now at 2, and the main sequence SEQ field is on page 3,
        // so the TOC entry will display "2>3" at its page count.
        builder.insertBreak(BreakType.PAGE_BREAK);
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("PrefixSequence");
        builder.insertParagraph();
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        builder.write("Second TOC entry, MySequence #");
        fieldSeq.setSequenceIdentifier("MySequence");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOC.SEQ.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.TOC.SEQ.docx");

        Assert.assertEquals(9, doc.getRange().getFields().getCount());

        fieldToc = (FieldToc)doc.getRange().getFields().get(0);
        System.out.println(fieldToc.getDisplayResult());
        TestUtil.verifyField(FieldType.FIELD_TOC, " TOC  \\c MySequence \\s PrefixSequence \\d >",
            "First TOC entry, MySequence #12\t\u0013 SEQ PrefixSequence _Toc256000000 \\* ARABIC \u00141\u0015>\u0013 PAGEREF _Toc256000000 \\h \u00142\u0015\r2" +
            "Second TOC entry, MySequence #\t\u0013 SEQ PrefixSequence _Toc256000001 \\* ARABIC \u00142\u0015>\u0013 PAGEREF _Toc256000001 \\h \u00143\u0015\r", 
            fieldToc);
        Assert.assertEquals("MySequence", fieldToc.getTableOfFiguresLabel());
        Assert.assertEquals("PrefixSequence", fieldToc.getPrefixedSequenceIdentifier());
        Assert.assertEquals(">", fieldToc.getSequenceSeparator());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ PrefixSequence _Toc256000000 \\* ARABIC ", "1", fieldSeq);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());

        // Byproduct field created by Aspose.Words
        FieldPageRef fieldPageRef = (FieldPageRef)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF _Toc256000000 \\h ", "2", fieldPageRef);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());
        Assert.assertEquals("_Toc256000000", fieldPageRef.getBookmarkName());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ PrefixSequence _Toc256000001 \\* ARABIC ", "2", fieldSeq);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());

        fieldPageRef = (FieldPageRef)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF _Toc256000001 \\h ", "3", fieldPageRef);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());
        Assert.assertEquals("_Toc256000001", fieldPageRef.getBookmarkName());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  PrefixSequence", "1", fieldSeq);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(6);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "1", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(7);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  PrefixSequence", "2", fieldSeq);
        Assert.assertEquals("PrefixSequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(8);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "2", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
    }

    @Test
    public void tocSeqNumbering() throws Exception
    {
        //ExStart
        //ExFor:FieldSeq
        //ExFor:FieldSeq.InsertNextNumber
        //ExFor:FieldSeq.ResetHeadingLevel
        //ExFor:FieldSeq.ResetNumber
        //ExFor:FieldSeq.SequenceIdentifier
        //ExSummary:Shows create numbering using SEQ fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Insert a SEQ field that will display the current count value of "MySequence",
        // after using the "ResetNumber" property to set it to 100.
        builder.write("#");
        FieldSeq fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setResetNumber("100");
        fieldSeq.update();

        Assert.assertEquals(" SEQ  MySequence \\r 100", fieldSeq.getFieldCode());
        Assert.assertEquals("100", fieldSeq.getResult());

        // Display the next number in this sequence with another SEQ field.
        builder.write(", #");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.update();

        Assert.assertEquals("101", fieldSeq.getResult());

        // Insert a level 1 heading.
        builder.insertBreak(BreakType.PARAGRAPH_BREAK);
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("This level 1 heading will reset MySequence to 1");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));

        // Insert another SEQ field from the same sequence and configure it to reset the count at every heading with 1.
        builder.write("\n#");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setResetHeadingLevel("1");
        fieldSeq.update();

        // The above heading is a level 1 heading, so the count for this sequence is reset to 1.
        Assert.assertEquals(" SEQ  MySequence \\s 1", fieldSeq.getFieldCode());
        Assert.assertEquals("1", fieldSeq.getResult());

        // Move to the next number of this sequence.
        builder.write(", #");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setInsertNextNumber(true);
        fieldSeq.update();

        Assert.assertEquals(" SEQ  MySequence \\n", fieldSeq.getFieldCode());
        Assert.assertEquals("2", fieldSeq.getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SEQ.ResetNumbering.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SEQ.ResetNumbering.docx");

        Assert.assertEquals(4, doc.getRange().getFields().getCount());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\r 100", "100", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "101", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\s 1", "1", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence \\n", "2", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
    }

    @Test (enabled = false, description = "WORDSNET-18083")
    public void tocSeqBookmark() throws Exception
    {
        //ExStart
        //ExFor:FieldSeq
        //ExFor:FieldSeq.BookmarkName
        //ExSummary:Shows how to combine table of contents and sequence fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // A TOC field can create an entry in its table of contents for each SEQ field found in the document.
        // Each entry contains the paragraph that contains the SEQ field,
        // and the number of the page that the field appears on.
        FieldToc fieldToc = (FieldToc)builder.insertField(FieldType.FIELD_TOC, true);

        // Configure this TOC field to have a SequenceIdentifier property with a value of "MySequence".
        fieldToc.setTableOfFiguresLabel("MySequence");

        // Configure this TOC field to only pick up SEQ fields that are within the bounds of a bookmark
        // named "TOCBookmark".
        fieldToc.setBookmarkName("TOCBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);

        Assert.assertEquals(" TOC  \\c MySequence \\b TOCBookmark", fieldToc.getFieldCode());

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Insert a SEQ field that has a sequence identifier that matches the TOC's
        // TableOfFiguresLabel property. This field will not create an entry in the TOC since it is outside
        // the bookmark's bounds designated by "BookmarkName".
        builder.write("MySequence #");
        FieldSeq fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        builder.writeln(", will not show up in the TOC because it is outside of the bookmark.");

        builder.startBookmark("TOCBookmark");

        // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bookmark's bounds.
        // The paragraph that contains this field will show up in the TOC as an entry.
        builder.write("MySequence #");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        builder.writeln(", will show up in the TOC next to the entry for the above caption.");

        // This SEQ field's sequence does not match the TOC's "TableOfFiguresLabel" property,
        // and is within the bounds of the bookmark. Its paragraph will not show up in the TOC as an entry.
        builder.write("MySequence #");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("OtherSequence");
        builder.writeln(", will not show up in the TOC because it's from a different sequence identifier.");

        // This SEQ field's sequence matches the TOC's "TableOfFiguresLabel" property and is within the bounds of the bookmark.
        // This field also references another bookmark. The contents of that bookmark will appear in the TOC entry for this SEQ field.
        // The SEQ field itself will not display the contents of that bookmark.
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        fieldSeq.setBookmarkName("SEQBookmark");
        Assert.assertEquals(" SEQ  MySequence SEQBookmark", fieldSeq.getFieldCode());

        // Create a bookmark with contents that will show up in the TOC entry due to the above SEQ field referencing it.
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark("SEQBookmark");
        builder.write("MySequence #");
        fieldSeq = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier("MySequence");
        builder.writeln(", text from inside SEQBookmark.");
        builder.endBookmark("SEQBookmark");

        builder.endBookmark("TOCBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SEQ.Bookmark.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SEQ.Bookmark.docx");

        Assert.assertEquals(8, doc.getRange().getFields().getCount());

        fieldToc = (FieldToc)doc.getRange().getFields().get(0);
        String[] pageRefIds = msString.split(fieldToc.getResult(), ' ').Where(s => s.StartsWith("_Toc")).ToArray();

        Assert.assertEquals(FieldType.FIELD_TOC, fieldToc.getType());
        Assert.assertEquals("MySequence", fieldToc.getTableOfFiguresLabel());
        TestUtil.verifyField(FieldType.FIELD_TOC, " TOC  \\c MySequence \\b TOCBookmark",
            $"MySequence #2, will show up in the TOC next to the entry for the above caption.\t\u0013 PAGEREF {pageRefIds[0]} \\h \u00142\u0015\r" +
            $"3MySequence #3, text from inside SEQBookmark.\t\u0013 PAGEREF {pageRefIds[1]} \\h \u00142\u0015\r", fieldToc);

        FieldPageRef fieldPageRef = (FieldPageRef)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, $" PAGEREF {pageRefIds[0]} \\h ", "2", fieldPageRef);
        Assert.assertEquals(pageRefIds[0], fieldPageRef.getBookmarkName());
        
        fieldPageRef = (FieldPageRef)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, $" PAGEREF {pageRefIds[1]} \\h ", "2", fieldPageRef);
        Assert.assertEquals(pageRefIds[1], fieldPageRef.getBookmarkName());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "1", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "2", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  OtherSequence", "1", fieldSeq);
        Assert.assertEquals("OtherSequence", fieldSeq.getSequenceIdentifier());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(6);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence SEQBookmark", "3", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
        Assert.assertEquals("SEQBookmark", fieldSeq.getBookmarkName());

        fieldSeq = (FieldSeq)doc.getRange().getFields().get(7);

        TestUtil.verifyField(FieldType.FIELD_SEQUENCE, " SEQ  MySequence", "3", fieldSeq);
        Assert.assertEquals("MySequence", fieldSeq.getSequenceIdentifier());
    }

    @Test (enabled = false, description = "WORDSNET-13854")
    public void fieldCitation() throws Exception
    {
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
        // Open a document containing bibliographical sources that we can find in
        // Microsoft Word via References -> Citations & Bibliography -> Manage Sources.
        Document doc = new Document(getMyDir() + "Bibliography.docx");
        Assert.assertEquals(2, doc.getRange().getFields().getCount()); //ExSkip

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Text to be cited with one source.");

        // Create a citation with just the page number and the author of the referenced book.
        FieldCitation fieldCitation = (FieldCitation)builder.insertField(FieldType.FIELD_CITATION, true);

        // We refer to sources using their tag names.
        fieldCitation.setSourceTag("Book1");
        fieldCitation.setPageNumber("85");
        fieldCitation.setSuppressAuthor(false);
        fieldCitation.setSuppressTitle(true);
        fieldCitation.setSuppressYear(true);

        Assert.assertEquals(" CITATION  Book1 \\p 85 \\t \\y", fieldCitation.getFieldCode());

        // Create a more detailed citation which cites two sources.
        builder.insertParagraph();
        builder.write("Text to be cited with two sources.");
        fieldCitation = (FieldCitation)builder.insertField(FieldType.FIELD_CITATION, true);
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

        // We can use a BIBLIOGRAPHY field to display all the sources within the document.
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldBibliography fieldBibliography = (FieldBibliography)builder.insertField(FieldType.FIELD_BIBLIOGRAPHY, true);
        fieldBibliography.setFormatLanguageId("1124");

        Assert.assertEquals(" BIBLIOGRAPHY  \\l 1124", fieldBibliography.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CITATION.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.CITATION.docx");

        Assert.assertEquals(5, doc.getRange().getFields().getCount());

        fieldCitation = (FieldCitation)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_CITATION, " CITATION  Book1 \\p 85 \\t \\y", " (Doe, p. 85)", fieldCitation);
        Assert.assertEquals("Book1", fieldCitation.getSourceTag());
        Assert.assertEquals("85", fieldCitation.getPageNumber());
        Assert.assertFalse(fieldCitation.getSuppressAuthor());
        Assert.assertTrue(fieldCitation.getSuppressTitle());
        Assert.assertTrue(fieldCitation.getSuppressYear());

        fieldCitation = (FieldCitation)doc.getRange().getFields().get(1);

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

        fieldBibliography = (FieldBibliography)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_BIBLIOGRAPHY, " BIBLIOGRAPHY  \\l 1124",
            "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);
        Assert.assertEquals("1124", fieldBibliography.getFormatLanguageId());

        fieldCitation = (FieldCitation)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_CITATION, " CITATION Book1 \\l 1033 ", "(Doe, 2018)", fieldCitation);
        Assert.assertEquals("Book1", fieldCitation.getSourceTag());
        Assert.assertEquals("1033", fieldCitation.getFormatLanguageId());

        fieldBibliography = (FieldBibliography)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_BIBLIOGRAPHY, " BIBLIOGRAPHY ", 
            "Cardholder, A. (2018). My Book, Vol. II. New York: Doe Co. Ltd.\rDoe, J. (2018). My Book, Vol I. London: Doe Co. Ltd.\r", fieldBibliography);
    }

    @Test
    public void fieldData() throws Exception
    {
        //ExStart
        //ExFor:FieldData
        //ExSummary:Shows how to insert a DATA field into a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldData field = (FieldData)builder.insertField(FieldType.FIELD_DATA, true);
        Assert.assertEquals(" DATA ", field.getFieldCode());
        //ExEnd
        
        TestUtil.verifyField(FieldType.FIELD_DATA, " DATA ", "", DocumentHelper.saveOpen(doc).getRange().getFields().get(0));
    }

    @Test
    public void fieldInclude() throws Exception
    {
        //ExStart
        //ExFor:FieldInclude
        //ExFor:FieldInclude.BookmarkName
        //ExFor:FieldInclude.LockFields
        //ExFor:FieldInclude.SourceFullName
        //ExFor:FieldInclude.TextConverter
        //ExSummary:Shows how to create an INCLUDE field, and set its properties.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We can use an INCLUDE field to import a portion of another document in the local file system.
        // The bookmark from the other document that we reference with this field contains this imported portion.
        FieldInclude field = (FieldInclude)builder.insertField(FieldType.FIELD_INCLUDE, true);
        field.setSourceFullName(getMyDir() + "Bookmarks.docx");
        field.setBookmarkName("MyBookmark1");
        field.setLockFields(false);
        field.setTextConverter("Microsoft Word");

        Assert.assertTrue(Regex.match(field.getFieldCode(), " INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").getSuccess());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INCLUDE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INCLUDE.docx");
        field = (FieldInclude)doc.getRange().getFields().get(0);

        Assert.assertEquals(FieldType.FIELD_INCLUDE, field.getType());
        Assert.assertEquals("First bookmark.", field.getResult());
        Assert.assertTrue(Regex.match(field.getFieldCode(), " INCLUDE .* MyBookmark1 \\\\c \"Microsoft Word\"").getSuccess());

        Assert.assertEquals(getMyDir() + "Bookmarks.docx", field.getSourceFullName());
        Assert.assertEquals("MyBookmark1", field.getBookmarkName());
        Assert.assertFalse(field.getLockFields());
        Assert.assertEquals("Microsoft Word", field.getTextConverter());
    }

    @Test
    public void fieldIncludePicture() throws Exception
    {
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

        // Below are two similar field types that we can use to display images linked from the local file system.
        // 1 -  The INCLUDEPICTURE field:
        FieldIncludePicture fieldIncludePicture = (FieldIncludePicture)builder.insertField(FieldType.FIELD_INCLUDE_PICTURE, true);
        fieldIncludePicture.setSourceFullName(getImageDir() + "Transparent background logo.png");

        Assert.assertTrue(Regex.match(fieldIncludePicture.getFieldCode(), " INCLUDEPICTURE  .*").getSuccess());

        // Apply the PNG32.FLT filter.
        fieldIncludePicture.setGraphicFilter("PNG32");
        fieldIncludePicture.isLinked(true);
        fieldIncludePicture.setResizeHorizontally(true);
        fieldIncludePicture.setResizeVertically(true);

        // 2 -  The IMPORT field:
        FieldImport fieldImport = (FieldImport)builder.insertField(FieldType.FIELD_IMPORT, true);
        fieldImport.setSourceFullName(getImageDir() + "Transparent background logo.png");
        fieldImport.setGraphicFilter("PNG32");
        fieldImport.isLinked(true);

        Assert.assertTrue(Regex.match(fieldImport.getFieldCode(), " IMPORT  .* \\\\c PNG32 \\\\d").getSuccess());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.IMPORT.INCLUDEPICTURE.docx");
        //ExEnd

        Assert.assertEquals(getImageDir() + "Transparent background logo.png", fieldIncludePicture.getSourceFullName());
        Assert.assertEquals("PNG32", fieldIncludePicture.getGraphicFilter());
        Assert.assertTrue(fieldIncludePicture.isLinked());
        Assert.assertTrue(fieldIncludePicture.getResizeHorizontally());
        Assert.assertTrue(fieldIncludePicture.getResizeVertically());

        Assert.assertEquals(getImageDir() + "Transparent background logo.png", fieldImport.getSourceFullName());
        Assert.assertEquals("PNG32", fieldImport.getGraphicFilter());
        Assert.assertTrue(fieldImport.isLinked());
        
        doc = new Document(getArtifactsDir() + "Field.IMPORT.INCLUDEPICTURE.docx");

        // The INCLUDEPICTURE fields have been converted into shapes with linked images during loading.
        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape image = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        Assert.assertTrue(image.isImage());
        Assert.assertNull(image.getImageData().getImageBytes());
        Assert.assertEquals(getImageDir() + "Transparent background logo.png", image.getImageData().getSourceFullName().replace("%20", " "));

        image = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        Assert.assertTrue(image.isImage());
        Assert.assertNull(image.getImageData().getImageBytes());
        Assert.assertEquals(getImageDir() + "Transparent background logo.png", image.getImageData().getSourceFullName().replace("%20", " "));
    }

    //ExStart
    //ExFor:FieldIncludeText
    //ExFor:FieldIncludeText.BookmarkName
    //ExFor:FieldIncludeText.Encoding
    //ExFor:FieldIncludeText.LockFields
    //ExFor:FieldIncludeText.MimeType
    //ExFor:FieldIncludeText.NamespaceMappings
    //ExFor:FieldIncludeText.SourceFullName
    //ExFor:FieldIncludeText.TextConverter
    //ExFor:FieldIncludeText.XPath
    //ExFor:FieldIncludeText.XslTransformation
    //ExSummary:Shows how to create an INCLUDETEXT field, and set its properties.
    @Test (enabled = false, description = "WORDSNET-17543") //ExSkip
    public void fieldIncludeText() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways to use INCLUDETEXT fields to display the contents of an XML file in the local file system.
        // 1 -  Perform an XSL transformation on an XML document:
        FieldIncludeText fieldIncludeText = createFieldIncludeText(builder, getMyDir() + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1");
        fieldIncludeText.setXslTransformation(getMyDir() + "CD collection XSL transformation.xsl");

        builder.writeln();

        // 2 -  Use an XPath to take specific elements from an XML document:
        fieldIncludeText = createFieldIncludeText(builder, getMyDir() + "CD collection data.xml", false, "text/xml", "XML", "ISO-8859-1");
        fieldIncludeText.setNamespaceMappings("xmlns:n='myNamespace'");
        fieldIncludeText.setXPath("/catalog/cd/title");

        doc.save(getArtifactsDir() + "Field.INCLUDETEXT.docx");
        testFieldIncludeText(new Document(getArtifactsDir() + "Field.INCLUDETEXT.docx")); //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field with custom properties.
    /// </summary>
    @Test (enabled = false)
    public FieldIncludeText createFieldIncludeText(DocumentBuilder builder, String sourceFullName, boolean lockFields, String mimeType, String textConverter, String encoding) throws Exception
    {
        FieldIncludeText fieldIncludeText = (FieldIncludeText)builder.insertField(FieldType.FIELD_INCLUDE_TEXT, true);
        fieldIncludeText.setSourceFullName(sourceFullName);
        fieldIncludeText.setLockFields(lockFields);
        fieldIncludeText.setMimeType(mimeType);
        fieldIncludeText.setTextConverter(textConverter);
        fieldIncludeText.setEncoding(encoding);

        return fieldIncludeText;
    }
    //ExEnd

    private void testFieldIncludeText(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        FieldIncludeText fieldIncludeText = (FieldIncludeText)doc.getRange().getFields().get(0);
        Assert.assertEquals(getMyDir() + "CD collection data.xml", fieldIncludeText.getSourceFullName());
        Assert.assertEquals(getMyDir() + "CD collection XSL transformation.xsl", fieldIncludeText.getXslTransformation());
        Assert.assertFalse(fieldIncludeText.getLockFields());
        Assert.assertEquals("text/xml", fieldIncludeText.getMimeType());
        Assert.assertEquals("XML", fieldIncludeText.getTextConverter());
        Assert.assertEquals("ISO-8859-1", fieldIncludeText.getEncoding());
        Assert.assertEquals(" INCLUDETEXT  \"" + getMyDir().replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\t \"" + 
                        getMyDir().replace("\\", "\\\\") + "CD collection XSL transformation.xsl\"", 
            fieldIncludeText.getFieldCode());
        Assert.assertTrue(fieldIncludeText.getResult().startsWith("My CD Collection"));

        org.w3c.dom.Document cdCollectionData = XmlUtilPal.newXmlDocument();
        cdCollectionData.LoadXml(File.readAllText(getMyDir() + "CD collection data.xml"));
        org.w3c.dom.Node catalogData = cdCollectionData.getChildNodes().item(0);

        org.w3c.dom.Document cdCollectionXslTransformation = XmlUtilPal.newXmlDocument();
        cdCollectionXslTransformation.LoadXml(File.readAllText(getMyDir() + "CD collection XSL transformation.xsl"));

        Table table = doc.getFirstSection().getBody().getTables().get(0);

        XmlNamespaceManager manager = new XmlNamespaceManager(cdCollectionXslTransformation.NameTable);
        manager.addNamespace("xsl", "http://www.w3.org/1999/XSL/Transform");

        for (int i = 0; i < table.getRows().getCount(); i++)
            for (int j = 0; j < table.getRows().get(i).getCount(); j++)
            {
                if (i == 0)
                {
                    // When on the first row from the input document's table, ensure that all table's cells match all XML element Names.
                    for (int k = 0; k < table.getRows().getCount() - 1; k++)
                        Assert.assertEquals(catalogData.getChildNodes().item(k).getChildNodes().item(j).getNodeName(),
                            table.getRows().get(i).getCells().get(j).getText().replace(ControlChar.CELL, "").toLowerCase());

                    // Also, make sure that the whole first row has the same color as the XSL transform.
                    Assert.assertEquals(cdCollectionXslTransformation.SelectNodes("//xsl:stylesheet/xsl:template/html/body/table/tr", manager).item(0).getAttributes().getNamedItem("bgcolor").getNodeValue(),
                        ColorTranslator.ToHtml(table.getRows().get(i).getCells().get(j).getCellFormat().getShading().getBackgroundPatternColor()).toLowerCase());
                }
                else
                {
                    // When on all other rows of the input document's table, ensure that cell contents match XML element Values.
                    Assert.assertEquals(catalogData.getChildNodes().item(i - 1).getChildNodes().item(j).getFirstChild().getNodeValue(),
                        table.getRows().get(i).getCells().get(j).getText().replace(ControlChar.CELL, ""));
                    Assert.assertEquals(msColor.Empty, table.getRows().get(i).getCells().get(j).getCellFormat().getShading().getBackgroundPatternColor());
                }

                Assert.assertEquals(
                    double.Parse(cdCollectionXslTransformation.SelectNodes("//xsl:stylesheet/xsl:template/html/body/table", manager).item(0).getAttributes().getNamedItem("border").getNodeValue()) * 0.75, 
                    table.getFirstRow().getRowFormat().getBorders().getBottom().getLineWidth());
            }

        fieldIncludeText = (FieldIncludeText)doc.getRange().getFields().get(1);
        Assert.assertEquals(getMyDir() + "CD collection data.xml", fieldIncludeText.getSourceFullName());
        Assert.assertNull(fieldIncludeText.getXslTransformation());
        Assert.assertFalse(fieldIncludeText.getLockFields());
        Assert.assertEquals("text/xml", fieldIncludeText.getMimeType());
        Assert.assertEquals("XML", fieldIncludeText.getTextConverter());
        Assert.assertEquals("ISO-8859-1", fieldIncludeText.getEncoding());
        Assert.assertEquals(" INCLUDETEXT  \"" + getMyDir().replace("\\", "\\\\") + "CD collection data.xml\" \\m text/xml \\c XML \\e ISO-8859-1 \\n xmlns:n='myNamespace' \\x /catalog/cd/title", 
            fieldIncludeText.getFieldCode());

        String expectedFieldResult = "";
        for (int i = 0; i < catalogData.getChildNodes().getLength(); i++)
        {
            expectedFieldResult = msString.plusEqOperator(expectedFieldResult, catalogData.getChildNodes().item(i).getChildNodes().item(0).getChildNodes().item(0).getNodeValue());
        }

        Assert.assertEquals(expectedFieldResult, fieldIncludeText.getResult());
    }

    @Test (enabled = false, description = "WORDSNET-17545")
    public void fieldHyperlink() throws Exception
    {
        //ExStart
        //ExFor:FieldHyperlink
        //ExFor:FieldHyperlink.Address
        //ExFor:FieldHyperlink.IsImageMap
        //ExFor:FieldHyperlink.OpenInNewWindow
        //ExFor:FieldHyperlink.ScreenTip
        //ExFor:FieldHyperlink.SubAddress
        //ExFor:FieldHyperlink.Target
        //ExSummary:Shows how to use HYPERLINK fields to link to documents in the local file system.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldHyperlink field = (FieldHyperlink)builder.insertField(FieldType.FIELD_HYPERLINK, true);

        // When we click this HYPERLINK field in Microsoft Word,
        // it will open the linked document and then place the cursor at the specified bookmark.
        field.setAddress(getMyDir() + "Bookmarks.docx");
        field.setSubAddress("MyBookmark3");
        field.setScreenTip("Open " + field.getAddress() + " on bookmark " + field.getSubAddress() + " in a new window");

        builder.writeln();

        // When we click this HYPERLINK field in Microsoft Word,
        // it will open the linked document, and automatically scroll down to the specified iframe.
        field = (FieldHyperlink)builder.insertField(FieldType.FIELD_HYPERLINK, true);
        field.setAddress(getMyDir() + "Iframes.html");
        field.setScreenTip("Open " + field.getAddress());
        field.setTarget("iframe_3");
        field.setOpenInNewWindow(true);
        field.isImageMap(false);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.HYPERLINK.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.HYPERLINK.docx");
        field = (FieldHyperlink)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_HYPERLINK, 
            " HYPERLINK \"" + getMyDir().replace("\\", "\\\\") + "Bookmarks.docx\" \\l \"MyBookmark3\" \\o \"Open " + getMyDir() + "Bookmarks.docx on bookmark MyBookmark3 in a new window\" ",
            getMyDir() + "Bookmarks.docx - MyBookmark3", field);
        Assert.assertEquals(getMyDir() + "Bookmarks.docx", field.getAddress());
        Assert.assertEquals("MyBookmark3", field.getSubAddress());
        Assert.assertEquals("Open " + field.getAddress().replace("\\", "") + " on bookmark " + field.getSubAddress() + " in a new window", field.getScreenTip());

        field = (FieldHyperlink)doc.getRange().getFields().get(1);

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
    //ExSummary:Shows how to set the dimensions of images as MERGEFIELDS accepts them during a mail merge.
    @Test //ExSkip
    public void mergeFieldImageDimension() throws Exception
    {
        Document doc = new Document();

        // Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
        // a column in the data source containing local system filenames of images we wish to use in the mail merge.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldMergeField field = (FieldMergeField)builder.insertField("MERGEFIELD Image:ImageColumn");

        // The data source should have such a column named "ImageColumn".
        Assert.assertEquals("Image:ImageColumn", field.getFieldName());

        // Create a suitable data source.
        DataTable dataTable = new DataTable("Images");
        dataTable.getColumns().add(new DataColumn("ImageColumn"));
        dataTable.getRows().add(getImageDir() + "Logo.jpg");
        dataTable.getRows().add(getImageDir() + "Transparent background logo.png");
        dataTable.getRows().add(getImageDir() + "Enhanced Windows MetaFile.emf");
        
        // Configure a callback to modify the sizes of images at merge time, then execute the mail merge.
        doc.getMailMerge().setFieldMergingCallback(new MergedImageResizer(200.0, 200.0, MergeFieldImageDimensionUnit.POINT));
        doc.getMailMerge().execute(dataTable);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.MERGEFIELD.ImageDimension.docx");
        testMergeFieldImageDimension(doc); //ExSkip
    }

    /// <summary>
    /// Sets the size of all mail merged images to one defined width and height.
    /// </summary>
    private static class MergedImageResizer implements IFieldMergingCallback
    {
        public MergedImageResizer(double imageWidth, double imageHeight, /*MergeFieldImageDimensionUnit*/int unit)
        {
            mImageWidth = imageWidth;
            mImageHeight = imageHeight;
            mUnit = unit;
        }

        public void fieldMerging(FieldMergingArgs e)
        {
            throw new UnsupportedOperationException();
        }

        public void imageFieldMerging(ImageFieldMergingArgs args)
        {
            args.setImageFileName(args.getFieldValue().toString());
            args.setImageWidth(new MergeFieldImageDimension(mImageWidth, mUnit));
            args.setImageHeight(new MergeFieldImageDimension(mImageHeight, mUnit));

            Assert.assertEquals(mImageWidth, args.getImageWidth().getValue());
            Assert.assertEquals(mUnit, args.getImageWidth().getUnit());
            Assert.assertEquals(mImageHeight, args.getImageHeight().getValue());
            Assert.assertEquals(mUnit, args.getImageHeight().getUnit());
        }

        private /*final*/ double mImageWidth;
        private /*final*/ double mImageHeight;
        private /*final*/ /*MergeFieldImageDimensionUnit*/int mUnit;
    }
    //ExEnd

    private void testMergeFieldImageDimension(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals(3, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, shape);
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 2, true);

        TestUtil.verifyImageInShape(534, 534, ImageType.EMF, shape);
        Assert.assertEquals(200.0d, shape.getWidth());
        Assert.assertEquals(200.0d, shape.getHeight());
    }

    //ExStart
    //ExFor:ImageFieldMergingArgs.Image
    //ExSummary:Shows how to use a callback to customize image merging logic.
    @Test //ExSkip
    public void mergeFieldImages() throws Exception
    {
        Document doc = new Document();

        // Insert a MERGEFIELD that will accept images from a source during a mail merge. Use the field code to reference
        // a column in the data source which contains local system filenames of images we wish to use in the mail merge.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldMergeField field = (FieldMergeField)builder.insertField("MERGEFIELD Image:ImageColumn");

        // In this case, the field expects the data source to have such a column named "ImageColumn".
        Assert.assertEquals("Image:ImageColumn", field.getFieldName());

        // Filenames can be lengthy, and if we can find a way to avoid storing them in the data source,
        // we may considerably reduce its size.
        // Create a data source that refers to images using short names.
        DataTable dataTable = new DataTable("Images");
        dataTable.getColumns().add(new DataColumn("ImageColumn"));
        dataTable.getRows().add("Dark logo");
        dataTable.getRows().add("Transparent logo");

        // Assign a merging callback that contains all logic that processes those names,
        // and then execute the mail merge. 
        doc.getMailMerge().setFieldMergingCallback(new ImageFilenameCallback());
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "Field.MERGEFIELD.Images.docx");
        testMergeFieldImages(new Document(getArtifactsDir() + "Field.MERGEFIELD.Images.docx")); //ExSkip
    }

    /// <summary>
    /// Contains a dictionary that maps names of images to local system filenames that contain these images.
    /// If a mail merge data source uses one of the dictionary's names to refer to an image,
    /// this callback will pass the respective filename to the merge destination.
    /// </summary>
    private static class ImageFilenameCallback implements IFieldMergingCallback
    {
        public ImageFilenameCallback()
        {
            mImageFilenames = new HashMap<String, String>();
            msDictionary.add(mImageFilenames, "Dark logo", getImageDir() + "Logo.jpg");
            msDictionary.add(mImageFilenames, "Transparent logo", getImageDir() + "Transparent background logo.png");
        }

        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args)
        {
            throw new UnsupportedOperationException();
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            if (mImageFilenames.containsKey(args.getFieldValue().toString()))
            {
                                args.setImage(ImageIO.read(mImageFilenames.get(args.getFieldValue().toString())));
                                                }
            
            Assert.assertNotNull(args.getImage());
        }

        private /*final*/ HashMap<String, String> mImageFilenames;
    }
    //ExEnd

    private void testMergeFieldImages(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals(2, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        Shape shape = (Shape)doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.JPEG, shape);
        Assert.assertEquals(300.0d, shape.getWidth());
        Assert.assertEquals(300.0d, shape.getHeight());

        shape = (Shape)doc.getChild(NodeType.SHAPE, 1, true);

        TestUtil.verifyImageInShape(400, 400, ImageType.PNG, shape);
        Assert.assertEquals(300.0d, shape.getWidth(), 1.0);
        Assert.assertEquals(300.0d, shape.getHeight(), 1.0);
    }

    @Test (enabled = false, description = "WORDSNET-17524")
    public void fieldIndexFilter() throws Exception
    {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.BookmarkName
        //ExFor:FieldIndex.EntryType
        //ExFor:FieldXE
        //ExFor:FieldXE.EntryType
        //ExFor:FieldXE.Text
        //ExSummary:Shows how to create an INDEX field, and then use XE fields to populate it with entries.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side
        // and the page containing the XE field on the right.
        // If the XE fields have the same value in their "Text" property,
        // the INDEX field will group them into one entry.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);

        // Configure the INDEX field only to display XE fields that are within the bounds
        // of a bookmark named "MainBookmark", and whose "EntryType" properties have a value of "A".
        // For both INDEX and XE fields, the "EntryType" property only uses the first character of its string value.
        index.setBookmarkName("MainBookmark");
        index.setEntryType("A");

        Assert.assertEquals(" INDEX  \\b MainBookmark \\f A", index.getFieldCode());

        // On a new page, start the bookmark with a name that matches the value
        // of the INDEX field's "BookmarkName" property.
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark("MainBookmark");

        // The INDEX field will pick up this entry because it is inside the bookmark,
        // and its entry type also matches the INDEX field's entry type.
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 1");
        indexEntry.setEntryType("A");

        Assert.assertEquals(" XE  \"Index entry 1\" \\f A", indexEntry.getFieldCode());

        // Insert an XE field that will not appear in the INDEX because the entry types do not match.
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 2");
        indexEntry.setEntryType("B");

        // End the bookmark and insert an XE field afterwards.
        // It is of the same type as the INDEX field, but will not appear
        // since it is outside the bookmark's boundaries.
        builder.endBookmark("MainBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 3");
        indexEntry.setEntryType("A");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Filtering.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Filtering.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\b MainBookmark \\f A", "Index entry 1, 2\r", index);
        Assert.assertEquals("MainBookmark", index.getBookmarkName());
        Assert.assertEquals("A", index.getEntryType());

        indexEntry = (FieldXE)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 1\" \\f A", "", indexEntry);
        Assert.assertEquals("Index entry 1", indexEntry.getText());
        Assert.assertEquals("A", indexEntry.getEntryType());

        indexEntry = (FieldXE)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 2\" \\f B", "", indexEntry);
        Assert.assertEquals("Index entry 2", indexEntry.getText());
        Assert.assertEquals("B", indexEntry.getEntryType());

        indexEntry = (FieldXE)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Index entry 3\" \\f A", "", indexEntry);
        Assert.assertEquals("Index entry 3", indexEntry.getText());
        Assert.assertEquals("A", indexEntry.getEntryType());
    }

    @Test (enabled = false, description = "WORDSNET-17524")
    public void fieldIndexFormatting() throws Exception
    {
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
        //ExSummary:Shows how to populate an INDEX field with entries using XE fields, and also modify its appearance.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // If the XE fields have the same value in their "Text" property,
        // the INDEX field will group them into one entry.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);
        index.setLanguageId("1033");

        // Setting this property's value to "A" will group all the entries by their first letter,
        // and place that letter in uppercase above each group.
        index.setHeading("A");

        // Set the table created by the INDEX field to span over 2 columns.
        index.setNumberOfColumns("2");

        // Set any entries with starting letters outside the "a-c" character range to be omitted.
        index.setLetterRange("a-c");

        Assert.assertEquals(" INDEX  \\z 1033 \\h A \\c 2 \\p a-c", index.getFieldCode());

        // These next two XE fields will show up under the "A" heading,
        // with their respective text stylings also applied to their page numbers.
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Apple");
        indexEntry.isItalic(true);

        Assert.assertEquals(" XE  Apple \\i", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Apricot");
        indexEntry.isBold(true);

        Assert.assertEquals(" XE  Apricot \\b", indexEntry.getFieldCode());

        // Both the next two XE fields will be under a "B" and "C" heading in the INDEX fields table of contents.
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Banana");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Cherry");

        // INDEX fields sort all entries alphabetically, so this entry will show up under "A" with the other two.
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Avocado");

        // This entry will not appear because it starts with the letter "D",
        // which is outside the "a-c" character range that the INDEX field's LetterRange property defines.
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Durian");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Formatting.docx");
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Formatting.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

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

        indexEntry = (FieldXE)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Apple \\i", "", indexEntry);
        Assert.assertEquals("Apple", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertTrue(indexEntry.isItalic());

        indexEntry = (FieldXE)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Apricot \\b", "", indexEntry);
        Assert.assertEquals("Apricot", indexEntry.getText());
        Assert.assertTrue(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Banana", "", indexEntry);
        Assert.assertEquals("Banana", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Cherry", "", indexEntry);
        Assert.assertEquals("Cherry", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Avocado", "", indexEntry);
        Assert.assertEquals("Avocado", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());

        indexEntry = (FieldXE)doc.getRange().getFields().get(6);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Durian", "", indexEntry);
        Assert.assertEquals("Durian", indexEntry.getText());
        Assert.assertFalse(indexEntry.isBold());
        Assert.assertFalse(indexEntry.isItalic());
    }

    @Test (enabled = false, description = "WORDSNET-17524")
    public void fieldIndexSequence() throws Exception
    {
        //ExStart
        //ExFor:FieldIndex.HasSequenceName
        //ExFor:FieldIndex.SequenceName
        //ExFor:FieldIndex.SequenceSeparator
        //ExSummary:Shows how to split a document into portions by combining INDEX and SEQ fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // If the XE fields have the same value in their "Text" property,
        // the INDEX field will group them into one entry.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);

        // In the SequenceName property, name a SEQ field sequence. Each entry of this INDEX field will now also display
        // the number that the sequence count is on at the XE field location that created this entry.
        index.setSequenceName("MySequence");

        // Set text that will around the sequence and page numbers to explain their meaning to the user.
        // An entry created with this configuration will display something like "MySequence at 1 on page 1" at its page number.
        // PageNumberSeparator and SequenceSeparator cannot be longer than 15 characters.
        index.setPageNumberSeparator("\tMySequence at ");
        index.setSequenceSeparator(" on page ");
        Assert.assertTrue(index.hasSequenceName());

        Assert.assertEquals(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.getFieldCode());

        // SEQ fields display a count that increments at each SEQ field.
        // These fields also maintain separate counts for each unique named sequence
        // identified by the SEQ field's "SequenceIdentifier" property.
        // Insert a SEQ field which moves the "MySequence" sequence to 1.
        // This field no different from normal document text. It will not appear on an INDEX field's table of contents.
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldSeq sequenceField = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        sequenceField.setSequenceIdentifier("MySequence");

        Assert.assertEquals(" SEQ  MySequence", sequenceField.getFieldCode());

        // Insert an XE field which will create an entry in the INDEX field.
        // Since "MySequence" is at 1 and this XE field is on page 2, along with the custom separators we defined above,
        // this field's INDEX entry will display "Cat" on the left side, and "MySequence at 1 on page 2" on the right.
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Cat");

        Assert.assertEquals(" XE  Cat", indexEntry.getFieldCode());

        // Insert a page break and use SEQ fields to advance "MySequence" to 3.
        builder.insertBreak(BreakType.PAGE_BREAK);
        sequenceField = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        sequenceField.setSequenceIdentifier("MySequence");
        sequenceField = (FieldSeq)builder.insertField(FieldType.FIELD_SEQUENCE, true);
        sequenceField.setSequenceIdentifier("MySequence");

        // Insert an XE field with the same Text property as the one above.
        // The INDEX entry will group XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        // Since we are on page 2 with "MySequence" at 3, ", 3 on page 3" will be appended to the same INDEX entry as above.
        // The page number portion of that INDEX entry will now display "MySequence at 1 on page 2, 3 on page 3".
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Cat");

        // Insert an XE field with a new and unique Text property value.
        // This will add a new entry, with MySequence at 3 on page 4.
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Dog");
        
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Sequence.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Sequence.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

        Assert.assertEquals("MySequence", index.getSequenceName());
        Assert.assertEquals("\tMySequence at ", index.getPageNumberSeparator());
        Assert.assertEquals(" on page ", index.getSequenceSeparator());
        Assert.assertTrue(index.hasSequenceName());
        Assert.assertEquals(" INDEX  \\s MySequence \\e \"\tMySequence at \" \\d \" on page \"", index.getFieldCode());
        Assert.assertEquals("Cat\tMySequence at 1 on page 2, 3 on page 3\r" +
                        "Dog\tMySequence at 3 on page 4\r", index.getResult());

        Assert.AreEqual(3, doc.getRange().getFields().Where(f => f.Type == FieldType.FieldSequence).Count());
    }

    @Test (enabled = false, description = "WORDSNET-17524")
    public void fieldIndexPageNumberSeparator() throws Exception
    {
        //ExStart
        //ExFor:FieldIndex.HasPageNumberSeparator
        //ExFor:FieldIndex.PageNumberSeparator
        //ExFor:FieldIndex.PageNumberListSeparator
        //ExSummary:Shows how to edit the page number separator in an INDEX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will group XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);

        // If our INDEX field has an entry for a group of XE fields,
        // this entry will display the number of each page that contains an XE field that belongs to this group.
        // We can set custom separators to customize the appearance of these page numbers.
        index.setPageNumberSeparator(", on page(s) ");
        index.setPageNumberListSeparator(" & ");
        
        Assert.assertEquals(" INDEX  \\e \", on page(s) \" \\l \" & \"", index.getFieldCode());
        Assert.assertTrue(index.hasPageNumberSeparator());

        // After we insert these XE fields, the INDEX field will display "First entry, on page(s) 2 & 3 & 4".
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("First entry");

        Assert.assertEquals(" XE  \"First entry\"", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("First entry");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("First entry");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.PageNumberList.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.PageNumberList.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\e \", on page(s) \" \\l \" & \"", "First entry, on page(s) 2 & 3 & 4\r", index);
        Assert.assertEquals(", on page(s) ", index.getPageNumberSeparator());
        Assert.assertEquals(" & ", index.getPageNumberListSeparator());
        Assert.assertTrue(index.hasPageNumberSeparator());
    }

    @Test (enabled = false, description = "WORDSNET-17524")
    public void fieldIndexPageRangeBookmark() throws Exception
    {
        //ExStart
        //ExFor:FieldIndex.PageRangeSeparator
        //ExFor:FieldXE.HasPageRangeBookmarkName
        //ExFor:FieldXE.PageRangeBookmarkName
        //ExSummary:Shows how to specify a bookmark's spanned pages as a page range for an INDEX field entry.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);

        // For INDEX entries that display page ranges, we can specify a separator string
        // which will appear between the number of the first page, and the number of the last.
        index.setPageNumberSeparator(", on page(s) ");
        index.setPageRangeSeparator(" to ");

        Assert.assertEquals(" INDEX  \\e \", on page(s) \" \\g \" to \"", index.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("My entry");

        // If an XE field names a bookmark using the PageRangeBookmarkName property,
        // its INDEX entry will show the range of pages that the bookmark spans
        // instead of the number of the page that contains the XE field.
        indexEntry.setPageRangeBookmarkName("MyBookmark");

        Assert.assertEquals(" XE  \"My entry\" \\r MyBookmark", indexEntry.getFieldCode());
        Assert.assertTrue(indexEntry.hasPageRangeBookmarkName());

        // Insert a bookmark that starts on page 3 and ends on page 5.
        // The INDEX entry for the XE field that references this bookmark will display this page range.
        // In our table, the INDEX entry will display "My entry, on page(s) 3 to 5".
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
        index = (FieldIndex)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\e \", on page(s) \" \\g \" to \"", "My entry, on page(s) 3 to 5\r", index);
        Assert.assertEquals(", on page(s) ", index.getPageNumberSeparator());
        Assert.assertEquals(" to ", index.getPageRangeSeparator());

        indexEntry = (FieldXE)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"My entry\" \\r MyBookmark", "", indexEntry);
        Assert.assertEquals("My entry", indexEntry.getText());
        Assert.assertEquals("MyBookmark", indexEntry.getPageRangeBookmarkName());
        Assert.assertTrue(indexEntry.hasPageRangeBookmarkName());
    }

    @Test (enabled = false, description = "WORDSNET-17524")
    public void fieldIndexCrossReferenceSeparator() throws Exception
    {
        //ExStart
        //ExFor:FieldIndex.CrossReferenceSeparator
        //ExFor:FieldXE.PageNumberReplacement
        //ExSummary:Shows how to define cross references in an INDEX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);

        // We can configure an XE field to get its INDEX entry to display a string instead of a page number.
        // First, for entries that substitute a page number with a string,
        // specify a custom separator between the XE field's Text property value and the string.
        index.setCrossReferenceSeparator(", see: ");

        Assert.assertEquals(" INDEX  \\k \", see: \"", index.getFieldCode());

        // Insert an XE field, which creates a regular INDEX entry which displays this field's page number,
        // and does not invoke the CrossReferenceSeparator value.
        // The entry for this XE field will display "Apple, 2".
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Apple");

        Assert.assertEquals(" XE  Apple", indexEntry.getFieldCode());

        // Insert another XE field on page 3 and set a value for the PageNumberReplacement property.
        // This value will show up instead of the number of the page that this field is on,
        // and the INDEX field's CrossReferenceSeparator value will appear in front of it.
        // The entry for this XE field will display "Banana, see: Tropical fruit".
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Banana");
        indexEntry.setPageNumberReplacement("Tropical fruit");

        Assert.assertEquals(" XE  Banana \\t \"Tropical fruit\"", indexEntry.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.CrossReferenceSeparator.docx");
        //ExEnd
        
        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.CrossReferenceSeparator.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " INDEX  \\k \", see: \"",
            "Apple, 2\r" +
            "Banana, see: Tropical fruit\r", index);
        Assert.assertEquals(", see: ", index.getCrossReferenceSeparator());

        indexEntry = (FieldXE)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Apple", "", indexEntry);
        Assert.assertEquals("Apple", indexEntry.getText());
        Assert.assertNull(indexEntry.getPageNumberReplacement());

        indexEntry = (FieldXE)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  Banana \\t \"Tropical fruit\"", "", indexEntry);
        Assert.assertEquals("Banana", indexEntry.getText());
        Assert.assertEquals("Tropical fruit", indexEntry.getPageNumberReplacement());
    }

    @Test (enabled = false, description = "WORDSNET-17524", dataProvider = "fieldIndexSubheadingDataProvider")
    public void fieldIndexSubheading(boolean runSubentriesOnTheSameLine) throws Exception
    {
        //ExStart
        //ExFor:FieldIndex.RunSubentriesOnSameLine
        //ExSummary:Shows how to work with subentries in an INDEX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);
        index.setPageNumberSeparator(", see page ");
        index.setHeading("A");

        // XE fields that have a Text property whose value becomes the heading of the INDEX entry.
        // If this value contains two string segments split by a colon (the INDEX entry will treat :) delimiter,
        // the first segment is heading, and the second segment will become the subheading.
        // The INDEX field first groups entries alphabetically, then, if there are multiple XE fields with the same
        // headings, the INDEX field will further subgroup them by the values of these headings.
        // There can be multiple subgrouping layers, depending on how many times
        // the Text properties of XE fields get segmented like this.
        // By default, an INDEX field entry group will create a new line for every subheading within this group. 
        // We can set the RunSubentriesOnSameLine flag to true to keep the heading,
        // and every subheading for the group on one line instead, which will make the INDEX field more compact.
        index.setRunSubentriesOnSameLine(runSubentriesOnTheSameLine);
        
        if (runSubentriesOnTheSameLine)
            Assert.assertEquals(" INDEX  \\e \", see page \" \\h A \\r", index.getFieldCode());
        else
            Assert.assertEquals(" INDEX  \\e \", see page \" \\h A", index.getFieldCode());

        // Insert two XE fields, each on a new page, and with the same heading named "Heading 1",
        // which the INDEX field will use to group them.
        // If RunSubentriesOnSameLine is false, then the INDEX table will create three lines:
        // one line for the grouping heading "Heading 1", and one more line for each subheading.
        // If RunSubentriesOnSameLine is true, then the INDEX table will create a one-line
        // entry that encompasses the heading and every subheading.
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Heading 1:Subheading 1");

        Assert.assertEquals(" XE  \"Heading 1:Subheading 1\"", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Heading 1:Subheading 2");
        
        doc.updateFields();
        doc.save(getArtifactsDir() + $"Field.INDEX.XE.Subheading.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + $"Field.INDEX.XE.Subheading.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

        if (runSubentriesOnTheSameLine)
        {
            TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\r \\e \", see page \" \\h A",
                "H\r" +
                "Heading 1: Subheading 1, see page 2; Subheading 2, see page 3\r", index);
            Assert.assertTrue(index.getRunSubentriesOnSameLine());
        }
        else
        {
            TestUtil.verifyField(FieldType.FIELD_INDEX, " INDEX  \\e \", see page \" \\h A",
                "H\r" +
                "Heading 1\r" +
                "Subheading 1, see page 2\r" +
                "Subheading 2, see page 3\r", index);
            Assert.assertFalse(index.getRunSubentriesOnSameLine());
        }

        indexEntry = (FieldXE)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Heading 1:Subheading 1\"", "", indexEntry);
        Assert.assertEquals("Heading 1:Subheading 1", indexEntry.getText());

        indexEntry = (FieldXE)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  \"Heading 1:Subheading 2\"", "", indexEntry);
        Assert.assertEquals("Heading 1:Subheading 2", indexEntry.getText());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fieldIndexSubheadingDataProvider")
	public static Object[][] fieldIndexSubheadingDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (enabled = false, description = "WORDSNET-17524", dataProvider = "fieldIndexYomiDataProvider")
    public void fieldIndexYomi(boolean sortEntriesUsingYomi) throws Exception
    {
        //ExStart
        //ExFor:FieldIndex.UseYomi
        //ExFor:FieldXE.Yomi
        //ExSummary:Shows how to sort INDEX field entries phonetically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an INDEX field which will display an entry for each XE field found in the document.
        // Each entry will display the XE field's Text property value on the left side,
        // and the number of the page that contains the XE field on the right.
        // The INDEX entry will collect all XE fields with matching values in the "Text" property
        // into one entry as opposed to making an entry for each XE field.
        FieldIndex index = (FieldIndex)builder.insertField(FieldType.FIELD_INDEX, true);

        // The INDEX table automatically sorts its entries by the values of their Text properties in alphabetic order.
        // Set the INDEX table to sort entries phonetically using Hiragana instead.
        index.setUseYomi(sortEntriesUsingYomi);

        if (sortEntriesUsingYomi)
            Assert.assertEquals(" INDEX  \\y", index.getFieldCode());
        else
            Assert.assertEquals(" INDEX ", index.getFieldCode());

        // Insert 4 XE fields, which would show up as entries in the INDEX field's table of contents.
        // The "Text" property may contain a word's spelling in Kanji, whose pronunciation may be ambiguous,
        // while the "Yomi" version of the word will spell exactly how it is pronounced using Hiragana.
        // If we set our INDEX field to use Yomi, it will sort these entries
        // by the value of their Yomi properties, instead of their Text values.
        builder.insertBreak(BreakType.PAGE_BREAK);
        FieldXE indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("愛子");
        indexEntry.setYomi("あ");

        Assert.assertEquals(" XE  愛子 \\y あ", indexEntry.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("明美");
        indexEntry.setYomi("あ");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("恵美");
        indexEntry.setYomi("え");

        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE)builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("愛美");
        indexEntry.setYomi("え");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.INDEX.XE.Yomi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INDEX.XE.Yomi.docx");
        index = (FieldIndex)doc.getRange().getFields().get(0);

        if (sortEntriesUsingYomi)
        {
            Assert.assertTrue(index.getUseYomi());
            Assert.assertEquals(" INDEX  \\y", index.getFieldCode());
            Assert.assertEquals("愛子, 2\r" +
                            "明美, 3\r" +
                            "恵美, 4\r" +
                            "愛美, 5\r", index.getResult());
        }
        else
        {
            Assert.assertFalse(index.getUseYomi());
            Assert.assertEquals(" INDEX ", index.getFieldCode());
            Assert.assertEquals("恵美, 4\r" +
                            "愛子, 2\r" +
                            "愛美, 5\r" +
                            "明美, 3\r", index.getResult());
        }

        indexEntry = (FieldXE)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  愛子 \\y あ", "", indexEntry);
        Assert.assertEquals("愛子", indexEntry.getText());
        Assert.assertEquals("あ", indexEntry.getYomi());

        indexEntry = (FieldXE)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  明美 \\y あ", "", indexEntry);
        Assert.assertEquals("明美", indexEntry.getText());
        Assert.assertEquals("あ", indexEntry.getYomi());

        indexEntry = (FieldXE)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  恵美 \\y え", "", indexEntry);
        Assert.assertEquals("恵美", indexEntry.getText());
        Assert.assertEquals("え", indexEntry.getYomi());

        indexEntry = (FieldXE)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_INDEX_ENTRY, " XE  愛美 \\y え", "", indexEntry);
        Assert.assertEquals("愛美", indexEntry.getText());
        Assert.assertEquals("え", indexEntry.getYomi());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fieldIndexYomiDataProvider")
	public static Object[][] fieldIndexYomiDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void fieldBarcode() throws Exception
    {
        //ExStart
        //ExFor:FieldBarcode
        //ExFor:FieldBarcode.FacingIdentificationMark
        //ExFor:FieldBarcode.IsBookmark
        //ExFor:FieldBarcode.IsUSPostalAddress
        //ExFor:FieldBarcode.PostalAddress
        //ExSummary:Shows how to use the BARCODE field to display U.S. ZIP codes in the form of a barcode. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln();

        // Below are two ways of using BARCODE fields to display custom values as barcodes.
        // 1 -  Store the value that the barcode will display in the PostalAddress property:
        FieldBarcode field = (FieldBarcode)builder.insertField(FieldType.FIELD_BARCODE, true);

        // This value needs to be a valid ZIP code.
        field.setPostalAddress("96801");
        field.isUSPostalAddress(true);
        field.setFacingIdentificationMark("C");

        Assert.assertEquals(" BARCODE  96801 \\u \\f C", field.getFieldCode());

        builder.insertBreak(BreakType.LINE_BREAK);

        // 2 -  Reference a bookmark that stores the value that this barcode will display:
        field = (FieldBarcode)builder.insertField(FieldType.FIELD_BARCODE, true);
        field.setPostalAddress("BarcodeBookmark");
        field.isBookmark(true);

        Assert.assertEquals(" BARCODE  BarcodeBookmark \\b", field.getFieldCode());

        // The bookmark that the BARCODE field references in its PostalAddress property
        // need to contain nothing besides the valid ZIP code.
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark("BarcodeBookmark");
        builder.writeln("968877");
        builder.endBookmark("BarcodeBookmark");

        doc.save(getArtifactsDir() + "Field.BARCODE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.BARCODE.docx");

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        field = (FieldBarcode)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_BARCODE, " BARCODE  96801 \\u \\f C", "", field);
        Assert.assertEquals("C", field.getFacingIdentificationMark());
        Assert.assertEquals("96801", field.getPostalAddress());
        Assert.assertTrue(field.isUSPostalAddress());

        field = (FieldBarcode)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_BARCODE, " BARCODE  BarcodeBookmark \\b", "", field);
        Assert.assertEquals("BarcodeBookmark", field.getPostalAddress());
        Assert.assertTrue(field.isBookmark());
    }

    @Test
    public void fieldDisplayBarcode() throws Exception
    {
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
        //ExSummary:Shows how to insert a DISPLAYBARCODE field, and set its properties. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldDisplayBarcode field = (FieldDisplayBarcode)builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);

        // Below are four types of barcodes, decorated in various ways, that the DISPLAYBARCODE field can display.
        // 1 -  QR code with custom colors:
        field.setBarcodeType("QR");
        field.setBarcodeValue("ABC123");
        field.setBackgroundColor("0xF8BD69");
        field.setForegroundColor("0xB5413B");
        field.setErrorCorrectionLevel("3");
        field.setScalingFactor("250");
        field.setSymbolHeight("1000");
        field.setSymbolRotation("0");

        Assert.assertEquals(" DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", field.getFieldCode());
        builder.writeln();

        // 2 -  EAN13 barcode, with the digits displayed below the bars:
        field = (FieldDisplayBarcode)builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);
        field.setBarcodeType("EAN13");
        field.setBarcodeValue("501234567890");
        field.setDisplayText(true);
        field.setPosCodeStyle("CASE");
        field.setFixCheckDigit(true);

        Assert.assertEquals(" DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", field.getFieldCode());
        builder.writeln();

        // 3 -  CODE39 barcode:
        field = (FieldDisplayBarcode)builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);
        field.setBarcodeType("CODE39");
        field.setBarcodeValue("12345ABCDE");
        field.setAddStartStopChar(true);

        Assert.assertEquals(" DISPLAYBARCODE  12345ABCDE CODE39 \\d", field.getFieldCode());
        builder.writeln();

        // 4 -  ITF4 barcode, with a specified case code:
        field = (FieldDisplayBarcode)builder.insertField(FieldType.FIELD_DISPLAY_BARCODE, true);
        field.setBarcodeType("ITF14");
        field.setBarcodeValue("09312345678907");
        field.setCaseCodeStyle("STD");

        Assert.assertEquals(" DISPLAYBARCODE  09312345678907 ITF14 \\c STD", field.getFieldCode());

        doc.save(getArtifactsDir() + "Field.DISPLAYBARCODE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DISPLAYBARCODE.docx");

        Assert.assertEquals(0, doc.getChildNodes(NodeType.SHAPE, true).getCount());

        field = (FieldDisplayBarcode)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  ABC123 QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0", "", field);
        Assert.assertEquals("QR", field.getBarcodeType());
        Assert.assertEquals("ABC123", field.getBarcodeValue());
        Assert.assertEquals("0xF8BD69", field.getBackgroundColor());
        Assert.assertEquals("0xB5413B", field.getForegroundColor());
        Assert.assertEquals("3", field.getErrorCorrectionLevel());
        Assert.assertEquals("250", field.getScalingFactor());
        Assert.assertEquals("1000", field.getSymbolHeight());
        Assert.assertEquals("0", field.getSymbolRotation());

        field = (FieldDisplayBarcode)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  501234567890 EAN13 \\t \\p CASE \\x", "", field);
        Assert.assertEquals("EAN13", field.getBarcodeType());
        Assert.assertEquals("501234567890", field.getBarcodeValue());
        Assert.assertTrue(field.getDisplayText());
        Assert.assertEquals("CASE", field.getPosCodeStyle());
        Assert.assertTrue(field.getFixCheckDigit());

        field = (FieldDisplayBarcode)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  12345ABCDE CODE39 \\d", "", field);
        Assert.assertEquals("CODE39", field.getBarcodeType());
        Assert.assertEquals("12345ABCDE", field.getBarcodeValue());
        Assert.assertTrue(field.getAddStartStopChar());

        field = (FieldDisplayBarcode)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, " DISPLAYBARCODE  09312345678907 ITF14 \\c STD", "", field);
        Assert.assertEquals("ITF14", field.getBarcodeType());
        Assert.assertEquals("09312345678907", field.getBarcodeValue());
        Assert.assertEquals("STD", field.getCaseCodeStyle());
    }

    @Test
    public void fieldMergeBarcode_QR() throws Exception
    {
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

        // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
        // This field will convert all values in a merge data source's "MyQRCode" column into QR codes.
        FieldMergeBarcode field = (FieldMergeBarcode)builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("QR");
        field.setBarcodeValue("MyQRCode");

        // Apply custom colors and scaling.
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

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
        // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
        // which will display a QR code with the value from the merged row.
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyQRCode");
        table.getRows().add(new String[] { "ABC123" });
        table.getRows().add(new String[] { "DEF456" });

        doc.getMailMerge().execute(table);

        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", 
            doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B",
            doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.QR.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.QR.docx");

        Assert.AreEqual(0, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldMergeBarcode));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, 
            "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", "", barcode);
        Assert.assertEquals("ABC123", barcode.getBarcodeValue());
        Assert.assertEquals("QR", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, 
            "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B", "", barcode);
        Assert.assertEquals("DEF456", barcode.getBarcodeValue());
        Assert.assertEquals("QR", barcode.getBarcodeType());
    }

    @Test
    public void fieldMergeBarcode_EAN13() throws Exception
    {
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

        // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
        // This field will convert all values in a merge data source's "MyEAN13Barcode" column into EAN13 barcodes.
        FieldMergeBarcode field = (FieldMergeBarcode)builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("EAN13");
        field.setBarcodeValue("MyEAN13Barcode");

        // Display the numeric value of the barcode underneath the bars.
        field.setDisplayText(true);
        field.setPosCodeStyle("CASE");
        field.setFixCheckDigit(true);

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x", field.getFieldCode());
        builder.writeln();

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
        // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
        // which will display an EAN13 barcode with the value from the merged row.
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyEAN13Barcode");
        table.getRows().add(new String[] { "501234567890" });
        table.getRows().add(new String[] { "123456789012" });

        doc.getMailMerge().execute(table);

        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x",
            doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x",
            doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.EAN13.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.EAN13.docx");

        Assert.AreEqual(0, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldMergeBarcode));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x", "", barcode);
        Assert.assertEquals("501234567890", barcode.getBarcodeValue());
        Assert.assertEquals("EAN13", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x", "", barcode);
        Assert.assertEquals("123456789012", barcode.getBarcodeValue());
        Assert.assertEquals("EAN13", barcode.getBarcodeType());
    }

    @Test
    public void fieldMergeBarcode_CODE39() throws Exception
    {
        //ExStart
        //ExFor:FieldMergeBarcode
        //ExFor:FieldMergeBarcode.AddStartStopChar
        //ExFor:FieldMergeBarcode.BarcodeType
        //ExSummary:Shows how to perform a mail merge on CODE39 barcodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
        // This field will convert all values in a merge data source's "MyCODE39Barcode" column into CODE39 barcodes.
        FieldMergeBarcode field = (FieldMergeBarcode)builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("CODE39");
        field.setBarcodeValue("MyCODE39Barcode");

        // Edit its appearance to display start/stop characters.
        field.setAddStartStopChar(true);

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyCODE39Barcode CODE39 \\d", field.getFieldCode());
        builder.writeln();

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
        // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
        // which will display a CODE39 barcode with the value from the merged row.
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyCODE39Barcode");
        table.getRows().add(new String[] { "12345ABCDE" });
        table.getRows().add(new String[] { "67890FGHIJ" });

        doc.getMailMerge().execute(table);

        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d",
            doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d",
            doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.CODE39.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.CODE39.docx");

        Assert.AreEqual(0, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldMergeBarcode));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d", "", barcode);
        Assert.assertEquals("12345ABCDE", barcode.getBarcodeValue());
        Assert.assertEquals("CODE39", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d", "", barcode);
        Assert.assertEquals("67890FGHIJ", barcode.getBarcodeValue());
        Assert.assertEquals("CODE39", barcode.getBarcodeType());
    }

    @Test
    public void fieldMergeBarcode_ITF14() throws Exception
    {
        //ExStart
        //ExFor:FieldMergeBarcode
        //ExFor:FieldMergeBarcode.BarcodeType
        //ExFor:FieldMergeBarcode.CaseCodeStyle
        //ExSummary:Shows how to perform a mail merge on ITF14 barcodes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEBARCODE field, which will accept values from a data source during a mail merge.
        // This field will convert all values in a merge data source's "MyITF14Barcode" column into ITF14 barcodes.
        FieldMergeBarcode field = (FieldMergeBarcode)builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("ITF14");
        field.setBarcodeValue("MyITF14Barcode");
        field.setCaseCodeStyle("STD");

        Assert.assertEquals(FieldType.FIELD_MERGE_BARCODE, field.getType());
        Assert.assertEquals(" MERGEBARCODE  MyITF14Barcode ITF14 \\c STD", field.getFieldCode());

        // Create a DataTable with a column with the same name as our MERGEBARCODE field's BarcodeValue.
        // The mail merge will create a new page for each row. Each page will contain a DISPLAYBARCODE field,
        // which will display an ITF14 barcode with the value from the merged row.
        DataTable table = new DataTable("Barcodes");
        table.getColumns().add("MyITF14Barcode");
        table.getRows().add(new String[] { "09312345678907" });
        table.getRows().add(new String[] { "1234567891234" });

        doc.getMailMerge().execute(table);

        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(0).getType());
        Assert.assertEquals("DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD",
            doc.getRange().getFields().get(0).getFieldCode());
        Assert.assertEquals(FieldType.FIELD_DISPLAY_BARCODE, doc.getRange().getFields().get(1).getType());
        Assert.assertEquals("DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD",
            doc.getRange().getFields().get(1).getFieldCode());

        doc.save(getArtifactsDir() + "Field.MERGEBARCODE.ITF14.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MERGEBARCODE.ITF14.docx");

        Assert.AreEqual(0, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldMergeBarcode));

        FieldDisplayBarcode barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DISPLAY_BARCODE, "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD", "", barcode);
        Assert.assertEquals("09312345678907", barcode.getBarcodeValue());
        Assert.assertEquals("ITF14", barcode.getBarcodeType());

        barcode = (FieldDisplayBarcode)doc.getRange().getFields().get(1);

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
    //ExSummary:Shows how to use various field types to link to other documents in the local file system, and display their contents.
    @Test (enabled = false, description = "WORDSNET-16226", dataProvider = "fieldLinkedObjectsAsTextDataProvider") //ExSkip
    public void fieldLinkedObjectsAsText(/*InsertLinkedObjectAs*/int insertLinkedObjectAs) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are three types of fields we can use to display contents from a linked document in the form of text.
        // 1 -  A LINK field:
        builder.writeln("FieldLink:\n");
        insertFieldLink(builder, insertLinkedObjectAs, "Word.Document.8", getMyDir() + "Document.docx", null, true);

        // 2 -  A DDE field:
        builder.writeln("FieldDde:\n");
        insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
            "Sheet1!R1C1", true, true);

        // 3 -  A DDEAUTO field:
        builder.writeln("FieldDdeAuto:\n");
        insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
            "Sheet1!R1C1", true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LINK.DDE.DDEAUTO.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fieldLinkedObjectsAsTextDataProvider")
	public static Object[][] fieldLinkedObjectsAsTextDataProvider() throws Exception
	{
		return new Object[][]
		{
			{InsertLinkedObjectAs.TEXT},
			{InsertLinkedObjectAs.UNICODE},
			{InsertLinkedObjectAs.HTML},
			{InsertLinkedObjectAs.RTF},
		};
	}

    @Test (enabled = false, description = "WORDSNET-16226", dataProvider = "fieldLinkedObjectsAsImageDataProvider") //ExSkip
    public void fieldLinkedObjectsAsImage(/*InsertLinkedObjectAs*/int insertLinkedObjectAs) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are three types of fields we can use to display contents from a linked document in the form of an image.
        // 1 -  A LINK field:
        builder.writeln("FieldLink:\n");
        insertFieldLink(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "MySpreadsheet.xlsx",
            "Sheet1!R2C2", true);

        // 2 -  A DDE field:
        builder.writeln("FieldDde:\n");
        insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
            "Sheet1!R1C1", true, true);

        // 3 -  A DDEAUTO field:
        builder.writeln("FieldDdeAuto:\n");
        insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Spreadsheet.xlsx",
            "Sheet1!R1C1", true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LINK.DDE.DDEAUTO.AsImage.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fieldLinkedObjectsAsImageDataProvider")
	public static Object[][] fieldLinkedObjectsAsImageDataProvider() throws Exception
	{
		return new Object[][]
		{
			{InsertLinkedObjectAs.PICTURE},
			{InsertLinkedObjectAs.BITMAP},
		};
	}

    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters.
    /// </summary>
    private static void insertFieldLink(DocumentBuilder builder, /*InsertLinkedObjectAs*/int insertLinkedObjectAs,
        String progId, String sourceFullName, String sourceItem, boolean shouldAutoUpdate) throws Exception
    {
        FieldLink field = (FieldLink)builder.insertField(FieldType.FIELD_LINK, true);

        switch (insertLinkedObjectAs)
        {
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
    /// Use a document builder to insert a DDE field, and set its properties according to parameters.
    /// </summary>
    private static void insertFieldDde(DocumentBuilder builder, /*InsertLinkedObjectAs*/int insertLinkedObjectAs, String progId,
        String sourceFullName, String sourceItem, boolean isLinked, boolean shouldAutoUpdate) throws Exception
    {
        FieldDde field = (FieldDde)builder.insertField(FieldType.FIELD_DDE, true);

        switch (insertLinkedObjectAs)
        {
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
    /// Use a document builder to insert a DDEAUTO, field and set its properties according to parameters.
    /// </summary>
    private static void insertFieldDdeAuto(DocumentBuilder builder, /*InsertLinkedObjectAs*/int insertLinkedObjectAs,
        String progId, String sourceFullName, String sourceItem, boolean isLinked) throws Exception
    {
        FieldDdeAuto field = (FieldDdeAuto)builder.insertField(FieldType.FIELD_DDE_AUTO, true);

        switch (insertLinkedObjectAs)
        {
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

    public /*enum*/ final class InsertLinkedObjectAs
    {
        private InsertLinkedObjectAs(){}
        
        // LinkedObjectAsText
        public static final int TEXT = 0;
        public static final int UNICODE = 1;
        public static final int HTML = 2;
        public static final int RTF = 3;
        // LinkedObjectAsImage
        public static final int PICTURE = 4;
        public static final int BITMAP = 5;

        public static final int length = 6;
    }
    //ExEnd

    @Test
    public void fieldUserAddress() throws Exception
    {
        //ExStart
        //ExFor:FieldUserAddress
        //ExFor:FieldUserAddress.UserAddress
        //ExSummary:Shows how to use the USERADDRESS field.
        Document doc = new Document();

        // Create a UserInformation object and set it as the source of user information for any fields that we create.
        UserInformation userInformation = new UserInformation();
        userInformation.setAddress("123 Main Street");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Create a USERADDRESS field to display the current user's address,
        // taken from the UserInformation object we created above.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldUserAddress fieldUserAddress = (FieldUserAddress)builder.insertField(FieldType.FIELD_USER_ADDRESS, true);
        Assert.assertEquals(userInformation.getAddress(), fieldUserAddress.getResult()); //ExSkip

        Assert.assertEquals(" USERADDRESS ", fieldUserAddress.getFieldCode());
        Assert.assertEquals("123 Main Street", fieldUserAddress.getResult());

        // We can set this property to get our field to override the value currently stored in the UserInformation object. 
        fieldUserAddress.setUserAddress("456 North Road");
        fieldUserAddress.update();

        Assert.assertEquals(" USERADDRESS  \"456 North Road\"", fieldUserAddress.getFieldCode());
        Assert.assertEquals("456 North Road", fieldUserAddress.getResult());

        // This does not affect the value in the UserInformation object.
        Assert.assertEquals("123 Main Street", doc.getFieldOptions().getCurrentUser().getAddress());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USERADDRESS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.USERADDRESS.docx");

        fieldUserAddress = (FieldUserAddress)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_USER_ADDRESS, " USERADDRESS  \"456 North Road\"", "456 North Road", fieldUserAddress);
        Assert.assertEquals("456 North Road", fieldUserAddress.getUserAddress());
    }

    @Test
    public void fieldUserInitials() throws Exception
    {
        //ExStart
        //ExFor:FieldUserInitials
        //ExFor:FieldUserInitials.UserInitials
        //ExSummary:Shows how to use the USERINITIALS field.
        Document doc = new Document();

        // Create a UserInformation object and set it as the source of user information for any fields that we create.
        UserInformation userInformation = new UserInformation();
        userInformation.setInitials("J. D.");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Create a USERINITIALS field to display the current user's initials,
        // taken from the UserInformation object we created above.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldUserInitials fieldUserInitials = (FieldUserInitials)builder.insertField(FieldType.FIELD_USER_INITIALS, true);
        Assert.assertEquals(userInformation.getInitials(), fieldUserInitials.getResult());

        Assert.assertEquals(" USERINITIALS ", fieldUserInitials.getFieldCode());
        Assert.assertEquals("J. D.", fieldUserInitials.getResult());

        // We can set this property to get our field to override the value currently stored in the UserInformation object. 
        fieldUserInitials.setUserInitials("J. C.");
        fieldUserInitials.update();

        Assert.assertEquals(" USERINITIALS  \"J. C.\"", fieldUserInitials.getFieldCode());
        Assert.assertEquals("J. C.", fieldUserInitials.getResult());

        // This does not affect the value in the UserInformation object.
        Assert.assertEquals("J. D.", doc.getFieldOptions().getCurrentUser().getInitials());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USERINITIALS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.USERINITIALS.docx");

        fieldUserInitials = (FieldUserInitials)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_USER_INITIALS, " USERINITIALS  \"J. C.\"", "J. C.", fieldUserInitials);
        Assert.assertEquals("J. C.", fieldUserInitials.getUserInitials());
    }

    @Test
    public void fieldUserName() throws Exception
    {
        //ExStart
        //ExFor:FieldUserName
        //ExFor:FieldUserName.UserName
        //ExSummary:Shows how to use the USERNAME field.
        Document doc = new Document();

        // Create a UserInformation object and set it as the source of user information for any fields that we create.
        UserInformation userInformation = new UserInformation();
        userInformation.setName("John Doe");
        doc.getFieldOptions().setCurrentUser(userInformation);

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a USERNAME field to display the current user's name,
        // taken from the UserInformation object we created above.
        FieldUserName fieldUserName = (FieldUserName)builder.insertField(FieldType.FIELD_USER_NAME, true);
        Assert.assertEquals(userInformation.getName(), fieldUserName.getResult());

        Assert.assertEquals(" USERNAME ", fieldUserName.getFieldCode());
        Assert.assertEquals("John Doe", fieldUserName.getResult());

        // We can set this property to get our field to override the value currently stored in the UserInformation object. 
        fieldUserName.setUserName("Jane Doe");
        fieldUserName.update();

        Assert.assertEquals(" USERNAME  \"Jane Doe\"", fieldUserName.getFieldCode());
        Assert.assertEquals("Jane Doe", fieldUserName.getResult());

        // This does not affect the value in the UserInformation object.
        Assert.assertEquals("John Doe", doc.getFieldOptions().getCurrentUser().getName());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USERNAME.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.USERNAME.docx");

        fieldUserName = (FieldUserName)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_USER_NAME, " USERNAME  \"Jane Doe\"", "Jane Doe", fieldUserName);
        Assert.assertEquals("Jane Doe", fieldUserName.getUserName());
    }

    @Test (enabled = false, description = "WORDSNET-17657")
    public void fieldStyleRefParagraphNumbers() throws Exception
    {
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

        // Create a list based using a Microsoft Word list template.
        List list = doc.getLists().add(com.aspose.words.ListTemplate.NUMBER_DEFAULT);

        // This generated list will display "1.a )".
        // Space before the bracket is a non-delimiter character, which we can suppress. 
        list.getListLevels().get(0).setNumberFormat("\u0000.");
        list.getListLevels().get(1).setNumberFormat("\u0001 )");

        // Add text and apply paragraph styles that STYLEREF fields will reference.
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

        // Place a STYLEREF field in the header and display the first "List Paragraph"-styled text in the document.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        FieldStyleRef field = (FieldStyleRef)builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("List Paragraph");

        // Place a STYLEREF field in the footer, and have it display the last text.
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        field = (FieldStyleRef)builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("List Paragraph");
        field.setSearchFromBottom(true);

        builder.moveToDocumentEnd();

        // We can also use STYLEREF fields to reference the list numbers of lists.
        builder.write("\nParagraph number: ");
        field = (FieldStyleRef)builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumber(true);

        builder.write("\nParagraph number, relative context: ");
        field = (FieldStyleRef)builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumberInRelativeContext(true);

        builder.write("\nParagraph number, full context: ");
        field = (FieldStyleRef)builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumberInFullContext(true);

        builder.write("\nParagraph number, full context, non-delimiter chars suppressed: ");
        field = (FieldStyleRef)builder.insertField(FieldType.FIELD_STYLE_REF, true);
        field.setStyleName("Quote");
        field.setInsertParagraphNumberInFullContext(true);
        field.setSuppressNonDelimiters(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.STYLEREF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.STYLEREF.docx");

        field = (FieldStyleRef)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  \"List Paragraph\"", "Item 1", field);
        Assert.assertEquals("List Paragraph", field.getStyleName());

        field = (FieldStyleRef)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  \"List Paragraph\" \\l", "Item 3", field);
        Assert.assertEquals("List Paragraph", field.getStyleName());
        Assert.assertTrue(field.getSearchFromBottom());

        field = (FieldStyleRef)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\n", "b )", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumber());

        field = (FieldStyleRef)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\r", "b )", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumberInRelativeContext());

        field = (FieldStyleRef)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\w", "1.b )", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumberInFullContext());

        field = (FieldStyleRef)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_STYLE_REF, " STYLEREF  Quote \\w \\t", "1.b)", field);
        Assert.assertEquals("Quote", field.getStyleName());
        Assert.assertTrue(field.getInsertParagraphNumberInFullContext());
        Assert.assertTrue(field.getSuppressNonDelimiters());
    }

    @Test
    public void fieldDate() throws Exception
    {
        //ExStart
        //ExFor:FieldDate
        //ExFor:FieldDate.UseLunarCalendar
        //ExFor:FieldDate.UseSakaEraCalendar
        //ExFor:FieldDate.UseUmAlQuraCalendar
        //ExFor:FieldDate.UseLastFormat
        //ExSummary:Shows how to use DATE fields to display dates according to different kinds of calendars.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we want the text in the document always to display the correct date, we can use a DATE field.
        // Below are three types of cultural calendars that a DATE field can use to display a date.
        // 1 -  Islamic Lunar Calendar:
        FieldDate field = (FieldDate)builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseLunarCalendar(true);
        Assert.assertEquals(" DATE  \\h", field.getFieldCode());
        builder.writeln();

        // 2 -  Umm al-Qura calendar:
        field = (FieldDate)builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseUmAlQuraCalendar(true);
        Assert.assertEquals(" DATE  \\u", field.getFieldCode());
        builder.writeln();

        // 3 -  Indian National Calendar:
        field = (FieldDate)builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseSakaEraCalendar(true);
        Assert.assertEquals(" DATE  \\s", field.getFieldCode());
        builder.writeln();

        // Insert a DATE field and set its calendar type to the one last used by the host application.
        // In Microsoft Word, the type will be the most recently used in the Insert -> Text -> Date and Time dialog box.
        field = (FieldDate)builder.insertField(FieldType.FIELD_DATE, true);
        field.setUseLastFormat(true);
        Assert.assertEquals(" DATE  \\l", field.getFieldCode());
        builder.writeln();

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.DATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DATE.docx");

        field = (FieldDate)doc.getRange().getFields().get(0);

        Assert.assertEquals(FieldType.FIELD_DATE, field.getType());
        Assert.assertTrue(field.getUseLunarCalendar());
        Assert.assertEquals(" DATE  \\h", field.getFieldCode());
        Assert.assertTrue(Regex.match(doc.getRange().getFields().get(0).getResult(), "\\d{1,2}[/]\\d{1,2}[/]\\d{4}").getSuccess());

        field = (FieldDate)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\u", new Date().toShortDateString(), field);
        Assert.assertTrue(field.getUseUmAlQuraCalendar());

        field = (FieldDate)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\s", new Date().toShortDateString(), field);
        Assert.assertTrue(field.getUseSakaEraCalendar());

        field = (FieldDate)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_DATE, " DATE  \\l", new Date().toShortDateString(), field);
        Assert.assertTrue(field.getUseLastFormat());
    }

    @Test (enabled = false, description = "WORDSNET-17669")
    public void fieldCreateDate() throws Exception
    {
        //ExStart
        //ExFor:FieldCreateDate
        //ExFor:FieldCreateDate.UseLunarCalendar
        //ExFor:FieldCreateDate.UseSakaEraCalendar
        //ExFor:FieldCreateDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to use the CREATEDATE field to display the creation date/time of the document.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln(" Date this document was created:");

        // We can use the CREATEDATE field to display the date and time of the creation of the document.
        // Below are three different calendar types according to which the CREATEDATE field can display the date/time.
        // 1 -  Islamic Lunar Calendar:
        builder.write("According to the Lunar Calendar - ");
        FieldCreateDate field = (FieldCreateDate)builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        field.setUseLunarCalendar(true);

        Assert.assertEquals(" CREATEDATE  \\h", field.getFieldCode());

        // 2 -  Umm al-Qura calendar:
        builder.write("\nAccording to the Umm al-Qura Calendar - ");
        field = (FieldCreateDate)builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        field.setUseUmAlQuraCalendar(true);

        Assert.assertEquals(" CREATEDATE  \\u", field.getFieldCode());

        // 3 -  Indian National Calendar:
        builder.write("\nAccording to the Indian National Calendar - ");
        field = (FieldCreateDate)builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        field.setUseSakaEraCalendar(true);

        Assert.assertEquals(" CREATEDATE  \\s", field.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CREATEDATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.CREATEDATE.docx");

        Assert.assertEquals(new DateTime(2017, 12, 5, 9, 56, 0), doc.getBuiltInDocumentProperties().getCreatedTimeInternal());

        DateTime expectedDate = doc.getBuiltInDocumentProperties().getCreatedTimeInternal().addHours(TimeZoneInfo.Local.GetUtcOffset(DateTime.getUtcNow()).getHours());
        field = (FieldCreateDate)doc.getRange().getFields().get(0);
        Calendar umAlQuraCalendar = new UmAlQuraCalendar();

        TestUtil.verifyField(FieldType.FIELD_CREATE_DATE, " CREATEDATE  \\h",
            $"{umAlQuraCalendar.GetMonth(expectedDate)}/{umAlQuraCalendar.GetDayOfMonth(expectedDate)}/{umAlQuraCalendar.GetYear(expectedDate)} " +
            expectedDate.addHours(1.0).toString("hh:mm:ss tt"), field);
        Assert.assertEquals(FieldType.FIELD_CREATE_DATE, field.getType());
        Assert.assertTrue(field.getUseLunarCalendar());
        
        field = (FieldCreateDate)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_CREATE_DATE, " CREATEDATE  \\u",
            $"{umAlQuraCalendar.GetMonth(expectedDate)}/{umAlQuraCalendar.GetDayOfMonth(expectedDate)}/{umAlQuraCalendar.GetYear(expectedDate)} " +
            expectedDate.addHours(1.0).toString("hh:mm:ss tt"), field);
        Assert.assertEquals(FieldType.FIELD_CREATE_DATE, field.getType());
        Assert.assertTrue(field.getUseUmAlQuraCalendar());
    }

    @Test (enabled = false, description = "WORDSNET-17669")
    public void fieldSaveDate() throws Exception
    {
        //ExStart
        //ExFor:BuiltInDocumentProperties.LastSavedTime
        //ExFor:FieldSaveDate
        //ExFor:FieldSaveDate.UseLunarCalendar
        //ExFor:FieldSaveDate.UseSakaEraCalendar
        //ExFor:FieldSaveDate.UseUmAlQuraCalendar
        //ExSummary:Shows how to use the SAVEDATE field to display the date/time of the document's most recent save operation performed using Microsoft Word.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.writeln(" Date this document was last saved:");

        // We can use the SAVEDATE field to display the last save operation's date and time on the document.
        // The save operation that these fields refer to is the manual save in an application like Microsoft Word,
        // not the document's Save method.
        // Below are three different calendar types according to which the SAVEDATE field can display the date/time.
        // 1 -  Islamic Lunar Calendar:
        builder.write("According to the Lunar Calendar - ");
        FieldSaveDate field = (FieldSaveDate)builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        field.setUseLunarCalendar(true);

        Assert.assertEquals(" SAVEDATE  \\h", field.getFieldCode());

        // 2 -  Umm al-Qura calendar:
        builder.write("\nAccording to the Umm al-Qura calendar - ");
        field = (FieldSaveDate)builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        field.setUseUmAlQuraCalendar(true);

        Assert.assertEquals(" SAVEDATE  \\u", field.getFieldCode());

        // 3 -  Indian National calendar:
        builder.write("\nAccording to the Indian National calendar - ");
        field = (FieldSaveDate)builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        field.setUseSakaEraCalendar(true);

        Assert.assertEquals(" SAVEDATE  \\s", field.getFieldCode());

        // The SAVEDATE fields draw their date/time values from the LastSavedTime built-in property.
        // The document's Save method will not update this value, but we can still update it manually.
        doc.getBuiltInDocumentProperties().setLastSavedTimeInternal(new Date());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SAVEDATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SAVEDATE.docx");

        System.out.println(doc.getBuiltInDocumentProperties().getLastSavedTimeInternal());

        field = (FieldSaveDate)doc.getRange().getFields().get(0);

        Assert.assertEquals(FieldType.FIELD_SAVE_DATE, field.getType());
        Assert.assertTrue(field.getUseLunarCalendar());
        Assert.assertEquals(" SAVEDATE  \\h", field.getFieldCode());

        Assert.assertTrue(Regex.match(field.getResult(), "\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M").getSuccess());

        field = (FieldSaveDate)doc.getRange().getFields().get(1);

        Assert.assertEquals(FieldType.FIELD_SAVE_DATE, field.getType());
        Assert.assertTrue(field.getUseUmAlQuraCalendar());
        Assert.assertEquals(" SAVEDATE  \\u", field.getFieldCode());
        Assert.assertTrue(Regex.match(field.getResult(), "\\d{1,2}[/]\\d{1,2}[/]\\d{4} \\d{1,2}:\\d{1,2}:\\d{1,2} [A,P]M").getSuccess());
    }

    @Test
    public void fieldBuilder() throws Exception
    {
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
        //ExSummary:Shows how to construct fields using a field builder, and then insert them into the document.
        Document doc = new Document();

        // Below are three examples of field construction done using a field builder.
        // 1 -  Single field:
        // Use a field builder to add a SYMBOL field which displays the ƒ (Florin) symbol.
        FieldBuilder builder = new FieldBuilder(FieldType.FIELD_SYMBOL);
        builder.addArgument(402);
        builder.addSwitch("\\f", "Arial");
        builder.addSwitch("\\s", 25);
        builder.addSwitch("\\u");
        Field field = builder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph());

        Assert.assertEquals(" SYMBOL 402 \\f Arial \\s 25 \\u ", field.getFieldCode());

        // 2 -  Nested field:
        // Use a field builder to create a formula field used as an inner field by another field builder.
        FieldBuilder innerFormulaBuilder = new FieldBuilder(FieldType.FIELD_FORMULA);
        innerFormulaBuilder.addArgument(100);
        innerFormulaBuilder.addArgument("+");
        innerFormulaBuilder.addArgument(74);

        // Create another builder for another SYMBOL field, and insert the formula field
        // that we have created above into the SYMBOL field as its argument. 
        builder = new FieldBuilder(FieldType.FIELD_SYMBOL);
        builder.addArgument(innerFormulaBuilder);
        field = builder.buildAndInsert(doc.getFirstSection().getBody().appendParagraph(""));

        // The outer SYMBOL field will use the formula field result, 174, as its argument,
        // which will make the field display the ® (Registered Sign) symbol since its character number is 174.
        Assert.assertEquals(" SYMBOL \u0013 = 100 + 74 \u0014\u0015 ", field.getFieldCode());

        // 3 -  Multiple nested fields and arguments:
        // Now, we will use a builder to create an IF field, which displays one of two custom string values,
        // depending on the true/false value of its expression. To get a true/false value
        // that determines which string the IF field displays, the IF field will test two numeric expressions for equality.
        // We will provide the two expressions in the form of formula fields, which we will nest inside the IF field.
        FieldBuilder leftExpression = new FieldBuilder(FieldType.FIELD_FORMULA);
        leftExpression.addArgument(2);
        leftExpression.addArgument("+");
        leftExpression.addArgument(3);

        FieldBuilder rightExpression = new FieldBuilder(FieldType.FIELD_FORMULA);
        rightExpression.addArgument(2.5);
        rightExpression.addArgument("*");
        rightExpression.addArgument(5.2);

        // Next, we will build two field arguments, which will serve as the true/false output strings for the IF field.
        // These arguments will reuse the output values of our numeric expressions.
        FieldArgumentBuilder trueOutput = new FieldArgumentBuilder();
        trueOutput.addText("True, both expressions amount to ");
        trueOutput.addField(leftExpression);

        FieldArgumentBuilder falseOutput = new FieldArgumentBuilder();
        falseOutput.addNode(new Run(doc, "False, "));
        falseOutput.addField(leftExpression);
        falseOutput.addNode(new Run(doc, " does not equal "));
        falseOutput.addField(rightExpression);

        // Finally, we will create one more field builder for the IF field and combine all of the expressions. 
        builder = new FieldBuilder(FieldType.FIELD_IF);
        builder.addArgument(leftExpression);
        builder.addArgument("=");
        builder.addArgument(rightExpression);
        builder.addArgument(trueOutput);
        builder.addArgument(falseOutput);
        field = builder.buildAndInsert(doc.getFirstSection().getBody().appendParagraph(""));

        Assert.assertEquals(" IF \u0013 = 2 + 3 \u0014\u0015 = \u0013 = 2.5 * 5.2 \u0014\u0015 " +
                        "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
                        "\"False, \u0013 = 2 + 3 \u0014\u0015 does not equal \u0013 = 2.5 * 5.2 \u0014\u0015\" ", field.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SYMBOL.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SYMBOL.docx");

        FieldSymbol fieldSymbol = (FieldSymbol)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL 402 \\f Arial \\s 25 \\u ", "", fieldSymbol);
        Assert.assertEquals("ƒ", fieldSymbol.getDisplayResult());

        fieldSymbol = (FieldSymbol)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL \u0013 = 100 + 74 \u0014174\u0015 ", "", fieldSymbol);
        Assert.assertEquals("®", fieldSymbol.getDisplayResult());

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 100 + 74 ", "174", doc.getRange().getFields().get(2));

        TestUtil.verifyField(FieldType.FIELD_IF,
            " IF \u0013 = 2 + 3 \u00145\u0015 = \u0013 = 2.5 * 5.2 \u001413\u0015 " +
            "\"True, both expressions amount to \u0013 = 2 + 3 \u0014\u0015\" " +
            "\"False, \u0013 = 2 + 3 \u00145\u0015 does not equal \u0013 = 2.5 * 5.2 \u001413\u0015\" ",
            "False, 5 does not equal 13", doc.getRange().getFields().get(3));

        Assert.<AssertionError>Throws(() => TestUtil.fieldsAreNested(doc.getRange().getFields().get(2), doc.getRange().getFields().get(3)));

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
    public void fieldAuthor() throws Exception
    {
        //ExStart
        //ExFor:FieldAuthor
        //ExFor:FieldAuthor.AuthorName  
        //ExFor:FieldOptions.DefaultDocumentAuthor
        //ExSummary:Shows how to use an AUTHOR field to display a document creator's name.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // AUTHOR fields source their results from the built-in document property called "Author".
        // If we create and save a document in Microsoft Word,
        // it will have our username in that property.
        // However, if we create a document programmatically using Aspose.Words,
        // the "Author" property, by default, will be an empty string. 
        Assert.assertEquals("", doc.getBuiltInDocumentProperties().getAuthor());

        // Set a backup author name for AUTHOR fields to use
        // if the "Author" property contains an empty string.
        doc.getFieldOptions().setDefaultDocumentAuthor("Joe Bloggs");

        builder.write("This document was created by ");
        FieldAuthor field = (FieldAuthor)builder.insertField(FieldType.FIELD_AUTHOR, true);
        field.update();

        Assert.assertEquals(" AUTHOR ", field.getFieldCode());
        Assert.assertEquals("Joe Bloggs", field.getResult());

        // Updating an AUTHOR field that contains a value
        // will apply that value to the "Author" built-in property.
        Assert.assertEquals("Joe Bloggs", doc.getBuiltInDocumentProperties().getAuthor());

        // Changing this property, then updating the AUTHOR field will apply this value to the field.
        doc.getBuiltInDocumentProperties().setAuthor("John Doe");      
        field.update();

        Assert.assertEquals(" AUTHOR ", field.getFieldCode());
        Assert.assertEquals("John Doe", field.getResult());
        
        // If we update an AUTHOR field after changing its "Name" property,
        // then the field will display the new name and apply the new name to the built-in property.
        field.setAuthorName("Jane Doe");
        field.update();

        Assert.assertEquals(" AUTHOR  \"Jane Doe\"", field.getFieldCode());
        Assert.assertEquals("Jane Doe", field.getResult());

        // AUTHOR fields do not affect the DefaultDocumentAuthor property.
        Assert.assertEquals("Jane Doe", doc.getBuiltInDocumentProperties().getAuthor());
        Assert.assertEquals("Joe Bloggs", doc.getFieldOptions().getDefaultDocumentAuthor());

        doc.save(getArtifactsDir() + "Field.AUTHOR.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.AUTHOR.docx");

        Assert.assertNull(doc.getFieldOptions().getDefaultDocumentAuthor());
        Assert.assertEquals("Jane Doe", doc.getBuiltInDocumentProperties().getAuthor());

        field = (FieldAuthor)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_AUTHOR, " AUTHOR  \"Jane Doe\"", "Jane Doe", field);
        Assert.assertEquals("Jane Doe", field.getAuthorName());
    }

    @Test
    public void fieldDocVariable() throws Exception
    {
        //ExStart
        //ExFor:FieldDocProperty
        //ExFor:FieldDocVariable
        //ExFor:FieldDocVariable.VariableName
        //ExSummary:Shows how to use DOCPROPERTY fields to display document properties and variables.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two ways of using DOCPROPERTY fields.
        // 1 -  Display a built-in property:
        // Set a custom value for the "Category" built-in property, then insert a DOCPROPERTY field that references it.
        doc.getBuiltInDocumentProperties().setCategory("My category");

        FieldDocProperty fieldDocProperty = (FieldDocProperty)builder.insertField(" DOCPROPERTY Category ");
        fieldDocProperty.update();

        Assert.assertEquals(" DOCPROPERTY Category ", fieldDocProperty.getFieldCode());
        Assert.assertEquals("My category", fieldDocProperty.getResult());

        builder.insertParagraph();

        // 2 -  Display a custom document variable:
        // Define a custom variable, then reference that variable with a DOCPROPERTY field.
        Assert.That(doc.getVariables(), Is.Empty);
        doc.getVariables().add("My variable", "My variable's value");

        FieldDocVariable fieldDocVariable = (FieldDocVariable)builder.insertField(FieldType.FIELD_DOC_VARIABLE, true);
        fieldDocVariable.setVariableName("My Variable");
        fieldDocVariable.update();

        Assert.assertEquals(" DOCVARIABLE  \"My Variable\"", fieldDocVariable.getFieldCode());
        Assert.assertEquals("My variable's value", fieldDocVariable.getResult());

        doc.save(getArtifactsDir() + "Field.DOCPROPERTY.DOCVARIABLE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.DOCPROPERTY.DOCVARIABLE.docx");

        Assert.assertEquals("My category", doc.getBuiltInDocumentProperties().getCategory());

        fieldDocProperty = (FieldDocProperty)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_DOC_PROPERTY, " DOCPROPERTY Category ", "My category", fieldDocProperty);

        fieldDocVariable = (FieldDocVariable)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_DOC_VARIABLE, " DOCVARIABLE  \"My Variable\"", "My variable's value", fieldDocVariable);
        Assert.assertEquals("My Variable", fieldDocVariable.getVariableName());
    }

    @Test
    public void fieldSubject() throws Exception
    {
        //ExStart
        //ExFor:FieldSubject
        //ExFor:FieldSubject.Text
        //ExSummary:Shows how to use the SUBJECT field.
        Document doc = new Document();

        // Set a value for the document's "Subject" built-in property.
        doc.getBuiltInDocumentProperties().setSubject("My subject");

        // Create a SUBJECT field to display the value of that built-in property.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldSubject field = (FieldSubject)builder.insertField(FieldType.FIELD_SUBJECT, true);
        field.update();

        Assert.assertEquals(" SUBJECT ", field.getFieldCode());
        Assert.assertEquals("My subject", field.getResult());

        // If we give the SUBJECT field's Text property value and update it, the field will
        // overwrite the current value of the "Subject" built-in property with the value of its Text property,
        // and then display the new value.
        field.setText("My new subject");
        field.update();

        Assert.assertEquals(" SUBJECT  \"My new subject\"", field.getFieldCode());
        Assert.assertEquals("My new subject", field.getResult());

        Assert.assertEquals("My new subject", doc.getBuiltInDocumentProperties().getSubject());

        doc.save(getArtifactsDir() + "Field.SUBJECT.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SUBJECT.docx");

        Assert.assertEquals("My new subject", doc.getBuiltInDocumentProperties().getSubject());

        field = (FieldSubject)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SUBJECT, " SUBJECT  \"My new subject\"", "My new subject", field);
        Assert.assertEquals("My new subject", field.getText());
    }

    @Test
    public void fieldComments() throws Exception
    {
        //ExStart
        //ExFor:FieldComments
        //ExFor:FieldComments.Text
        //ExSummary:Shows how to use the COMMENTS field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a value for the document's "Comments" built-in property.
        doc.getBuiltInDocumentProperties().setComments("My comment.");

        // Create a COMMENTS field to display the value of that built-in property.
        FieldComments field = (FieldComments)builder.insertField(FieldType.FIELD_COMMENTS, true);
        field.update();

        Assert.assertEquals(" COMMENTS ", field.getFieldCode());
        Assert.assertEquals("My comment.", field.getResult());

        // If we give the COMMENTS field's Text property value and update it, the field will
        // overwrite the current value of the "Comments" built-in property with the value of its Text property,
        // and then display the new value.
        field.setText("My overriding comment.");
        field.update();

        Assert.assertEquals(" COMMENTS  \"My overriding comment.\"", field.getFieldCode());
        Assert.assertEquals("My overriding comment.", field.getResult());

        doc.save(getArtifactsDir() + "Field.COMMENTS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.COMMENTS.docx");

        Assert.assertEquals("My overriding comment.", doc.getBuiltInDocumentProperties().getComments());

        field = (FieldComments)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_COMMENTS, " COMMENTS  \"My overriding comment.\"", "My overriding comment.", field);
        Assert.assertEquals("My overriding comment.", field.getText());
    }
    
    @Test
    public void fieldFileSize() throws Exception
    {
        //ExStart
        //ExFor:FieldFileSize
        //ExFor:FieldFileSize.IsInKilobytes
        //ExFor:FieldFileSize.IsInMegabytes            
        //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
        Document doc = new Document(getMyDir() + "Document.docx");

        Assert.assertEquals(16222, doc.getBuiltInDocumentProperties().getBytes());

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();
        builder.insertParagraph();

        // Below are three different units of measure
        // with which FILESIZE fields can display the document's file size.
        // 1 -  Bytes:
        FieldFileSize field = (FieldFileSize)builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.update();

        Assert.assertEquals(" FILESIZE ", field.getFieldCode());
        Assert.assertEquals("16222", field.getResult());

        // 2 -  Kilobytes:
        builder.insertParagraph();
        field = (FieldFileSize)builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.isInKilobytes(true);
        field.update();

        Assert.assertEquals(" FILESIZE  \\k", field.getFieldCode());
        Assert.assertEquals("16", field.getResult());

        // 3 -  Megabytes:
        builder.insertParagraph();
        field = (FieldFileSize)builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.isInMegabytes(true);
        field.update();

        Assert.assertEquals(" FILESIZE  \\m", field.getFieldCode());
        Assert.assertEquals("0", field.getResult());

        // To update the values of these fields while editing in Microsoft Word,
        // we must first save the changes, and then manually update these fields.
        doc.save(getArtifactsDir() + "Field.FILESIZE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.FILESIZE.docx");

        field = (FieldFileSize)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_FILE_SIZE, " FILESIZE ", "16222", field);

        // These fields will need to be updated to produce an accurate result.
        doc.updateFields();

        field = (FieldFileSize)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_FILE_SIZE, " FILESIZE  \\k", "13", field);
        Assert.assertTrue(field.isInKilobytes());

        field = (FieldFileSize)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_FILE_SIZE, " FILESIZE  \\m", "0", field);
        Assert.assertTrue(field.isInMegabytes());
    }

    @Test
    public void fieldGoToButton() throws Exception
    {
        //ExStart
        //ExFor:FieldGoToButton
        //ExFor:FieldGoToButton.DisplayText
        //ExFor:FieldGoToButton.Location
        //ExSummary:Shows to insert a GOTOBUTTON field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a GOTOBUTTON field. When we double-click this field in Microsoft Word,
        // it will take the text cursor to the bookmark whose name the Location property references.
        FieldGoToButton field = (FieldGoToButton)builder.insertField(FieldType.FIELD_GO_TO_BUTTON, true);
        field.setDisplayText("My Button");
        field.setLocation("MyBookmark");

        Assert.assertEquals(" GOTOBUTTON  MyBookmark My Button", field.getFieldCode());

        // Insert a valid bookmark for the field to reference.
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark(field.getLocation());
        builder.writeln("Bookmark text contents.");
        builder.endBookmark(field.getLocation());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.GOTOBUTTON.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.GOTOBUTTON.docx");
        field = (FieldGoToButton)doc.getRange().getFields().get(0);

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
    public void fieldFillIn() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a FILLIN field. When we manually update this field in Microsoft Word,
        // it will prompt us to enter a response. The field will then display the response as text.
        FieldFillIn field = (FieldFillIn)builder.insertField(FieldType.FIELD_FILL_IN, true);
        field.setPromptText("Please enter a response:");
        field.setDefaultResponse("A default response.");

        // We can also use these fields to ask the user for a unique response for each page
        // created during a mail merge done using Microsoft Word.
        field.setPromptOnceOnMailMerge(true);

        Assert.assertEquals(" FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", field.getFieldCode());

        FieldMergeField mergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        mergeField.setFieldName("MergeField");
        
        // If we perform a mail merge programmatically, we can use a custom prompt respondent
        // to automatically edit responses for FILLIN fields that the mail merge encounters.
        doc.getFieldOptions().setUserPromptRespondent(new PromptRespondent());
        doc.getMailMerge().execute(new String[] { "MergeField" }, new Object[] { "" });
        
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FILLIN.docx");
        testFieldFillIn(new Document(getArtifactsDir() + "Field.FILLIN.docx")); //ExSKip
    }

    /// <summary>
    /// Prepends a line to the default response of every FILLIN field during a mail merge.
    /// </summary>
    private static class PromptRespondent implements IFieldUserPromptRespondent
    {
        public String respond(String promptText, String defaultResponse)
        {
            return "Response modified by PromptRespondent. " + defaultResponse;
        }
    }
    //ExEnd

    private void testFieldFillIn(Document doc) throws Exception
    {
        doc = DocumentHelper.saveOpen(doc);

        Assert.assertEquals(1, doc.getRange().getFields().getCount());

        FieldFillIn field = (FieldFillIn)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_FILL_IN, " FILLIN  \"Please enter a response:\" \\d \"A default response.\" \\o", 
            "Response modified by PromptRespondent. A default response.", field);
        Assert.assertEquals("Please enter a response:", field.getPromptText());
        Assert.assertEquals("A default response.", field.getDefaultResponse());
        Assert.assertTrue(field.getPromptOnceOnMailMerge());
    }

    @Test
    public void fieldInfo() throws Exception
    {
        //ExStart
        //ExFor:FieldInfo
        //ExFor:FieldInfo.InfoType
        //ExFor:FieldInfo.NewValue
        //ExSummary:Shows how to work with INFO fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a value for the "Comments" built-in property and then insert an INFO field to display that property's value.
        doc.getBuiltInDocumentProperties().setComments("My comment");
        FieldInfo field = (FieldInfo)builder.insertField(FieldType.FIELD_INFO, true);
        field.setInfoType("Comments");
        field.update();

        Assert.assertEquals(" INFO  Comments", field.getFieldCode());
        Assert.assertEquals("My comment", field.getResult());

        builder.writeln();

        // Setting a value for the field's NewValue property and updating
        // the field will also overwrite the corresponding built-in property with the new value.
        field = (FieldInfo)builder.insertField(FieldType.FIELD_INFO, true);
        field.setInfoType("Comments");
        field.setNewValue("New comment");
        field.update();

        Assert.assertEquals(" INFO  Comments \"New comment\"", field.getFieldCode());
        Assert.assertEquals("New comment", field.getResult());
        Assert.assertEquals("New comment", doc.getBuiltInDocumentProperties().getComments());

        doc.save(getArtifactsDir() + "Field.INFO.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.INFO.docx");

        Assert.assertEquals("New comment", doc.getBuiltInDocumentProperties().getComments());
        
        field = (FieldInfo)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_INFO, " INFO  Comments", "My comment", field);
        Assert.assertEquals("Comments", field.getInfoType());

        field = (FieldInfo)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_INFO, " INFO  Comments \"New comment\"", "New comment", field);
        Assert.assertEquals("Comments", field.getInfoType());
        Assert.assertEquals("New comment", field.getNewValue());
    }

    @Test
    public void fieldMacroButton() throws Exception
    {
        //ExStart
        //ExFor:Document.HasMacros
        //ExFor:FieldMacroButton
        //ExFor:FieldMacroButton.DisplayText
        //ExFor:FieldMacroButton.MacroName
        //ExSummary:Shows how to use MACROBUTTON fields to allow us to run a document's macros by clicking.
        Document doc = new Document(getMyDir() + "Macro.docm");
        DocumentBuilder builder = new DocumentBuilder(doc);

        Assert.assertTrue(doc.hasMacros());

        // Insert a MACROBUTTON field, and reference one of the document's macros by name in the MacroName property.
        FieldMacroButton field = (FieldMacroButton)builder.insertField(FieldType.FIELD_MACRO_BUTTON, true);
        field.setMacroName("MyMacro");
        field.setDisplayText("Double click to run macro: " + field.getMacroName());

        Assert.assertEquals(" MACROBUTTON  MyMacro Double click to run macro: MyMacro", field.getFieldCode());

        // Use the property to reference "ViewZoom200", a macro that ships with Microsoft Word.
        // We can find all other macros via View -> Macros (dropdown) -> View Macros.
        // In that menu, select "Word Commands" from the "Macros in:" drop down.
        // If our document contains a custom macro with the same name as a stock macro,
        // our macro will be the one that the MACROBUTTON field runs.
        builder.insertParagraph();
        field = (FieldMacroButton)builder.insertField(FieldType.FIELD_MACRO_BUTTON, true);
        field.setMacroName("ViewZoom200");
        field.setDisplayText("Run " + field.getMacroName());

        Assert.assertEquals(" MACROBUTTON  ViewZoom200 Run ViewZoom200", field.getFieldCode());

        // Save the document as a macro-enabled document type.
        doc.save(getArtifactsDir() + "Field.MACROBUTTON.docm");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.MACROBUTTON.docm");

        field = (FieldMacroButton)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_MACRO_BUTTON, " MACROBUTTON  MyMacro Double click to run macro: MyMacro", "", field);
        Assert.assertEquals("MyMacro", field.getMacroName());
        Assert.assertEquals("Double click to run macro: MyMacro", field.getDisplayText());

        field = (FieldMacroButton)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_MACRO_BUTTON, " MACROBUTTON  ViewZoom200 Run ViewZoom200", "", field);
        Assert.assertEquals("ViewZoom200", field.getMacroName());
        Assert.assertEquals("Run ViewZoom200", field.getDisplayText());
    }

    @Test
    public void fieldKeywords() throws Exception
    {
        //ExStart
        //ExFor:FieldKeywords
        //ExFor:FieldKeywords.Text
        //ExSummary:Shows to insert a KEYWORDS field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some keywords, also referred to as "tags" in File Explorer.
        doc.getBuiltInDocumentProperties().setKeywords("Keyword1, Keyword2");

        // The KEYWORDS field displays the value of this property.
        FieldKeywords field = (FieldKeywords)builder.insertField(FieldType.FIELD_KEYWORD, true);
        field.update();

        Assert.assertEquals(" KEYWORDS ", field.getFieldCode());
        Assert.assertEquals("Keyword1, Keyword2", field.getResult());

        // Setting a value for the field's Text property,
        // and then updating the field will also overwrite the corresponding built-in property with the new value.
        field.setText("OverridingKeyword");
        field.update();

        Assert.assertEquals(" KEYWORDS  OverridingKeyword", field.getFieldCode());
        Assert.assertEquals("OverridingKeyword", field.getResult());
        Assert.assertEquals("OverridingKeyword", doc.getBuiltInDocumentProperties().getKeywords());

        doc.save(getArtifactsDir() + "Field.KEYWORDS.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.KEYWORDS.docx");

        Assert.assertEquals("OverridingKeyword", doc.getBuiltInDocumentProperties().getKeywords());

        field = (FieldKeywords)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_KEYWORD, " KEYWORDS  OverridingKeyword", "OverridingKeyword", field);
        Assert.assertEquals("OverridingKeyword", field.getText());
    }

    @Test
    public void fieldNum() throws Exception
    {
        //ExStart
        //ExFor:FieldPage
        //ExFor:FieldNumChars
        //ExFor:FieldNumPages
        //ExFor:FieldNumWords
        //ExSummary:Shows how to use NUMCHARS, NUMWORDS, NUMPAGES and PAGE fields to track the size of our documents.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Below are three types of fields that we can use to track the size of our documents.
        // 1 -  Track the character count with a NUMCHARS field:
        FieldNumChars fieldNumChars = (FieldNumChars)builder.insertField(FieldType.FIELD_NUM_CHARS, true);       
        builder.writeln(" characters");

        // 2 -  Track the word count with a NUMWORDS field:
        FieldNumWords fieldNumWords = (FieldNumWords)builder.insertField(FieldType.FIELD_NUM_WORDS, true);
        builder.writeln(" words");

        // 3 -  Use both PAGE and NUMPAGES fields to display what page the field is on,
        // and the total number of pages in the document:
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Page ");
        FieldPage fieldPage = (FieldPage)builder.insertField(FieldType.FIELD_PAGE, true);
        builder.write(" of ");
        FieldNumPages fieldNumPages = (FieldNumPages)builder.insertField(FieldType.FIELD_NUM_PAGES, true);

        Assert.assertEquals(" NUMCHARS ", fieldNumChars.getFieldCode());
        Assert.assertEquals(" NUMWORDS ", fieldNumWords.getFieldCode());
        Assert.assertEquals(" NUMPAGES ", fieldNumPages.getFieldCode());
        Assert.assertEquals(" PAGE ", fieldPage.getFieldCode());

        // These fields will not maintain accurate values in real time
        // while we edit the document programmatically using Aspose.Words, or in Microsoft Word.
        // We need to update them every we need to see an up-to-date value. 
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
    public void fieldPrint() throws Exception
    {
        //ExStart
        //ExFor:FieldPrint
        //ExFor:FieldPrint.PostScriptGroup
        //ExFor:FieldPrint.PrinterInstructions
        //ExSummary:Shows to insert a PRINT field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("My paragraph");

        // The PRINT field can send instructions to the printer.
        FieldPrint field = (FieldPrint)builder.insertField(FieldType.FIELD_PRINT, true);

        // Set the area for the printer to perform instructions over.
        // In this case, it will be the paragraph that contains our PRINT field.
        field.setPostScriptGroup("para");

        // When we use a printer that supports PostScript to print our document,
        // this command will turn the entire area that we specified in "field.PostScriptGroup" white.
        field.setPrinterInstructions("erasepage");

        Assert.assertEquals(" PRINT  erasepage \\p para", field.getFieldCode());
        
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.PRINT.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.PRINT.docx");

        field = (FieldPrint)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_PRINT, " PRINT  erasepage \\p para", "", field);
        Assert.assertEquals("para", field.getPostScriptGroup());
        Assert.assertEquals("erasepage", field.getPrinterInstructions());
    }

    @Test
    public void fieldPrintDate() throws Exception
    {
        //ExStart
        //ExFor:FieldPrintDate
        //ExFor:FieldPrintDate.UseLunarCalendar
        //ExFor:FieldPrintDate.UseSakaEraCalendar
        //ExFor:FieldPrintDate.UseUmAlQuraCalendar
        //ExSummary:Shows read PRINTDATE fields.
        Document doc = new Document(getMyDir() + "Field sample - PRINTDATE.docx");

        // When a document is printed by a printer or printed as a PDF (but not exported to PDF),
        // PRINTDATE fields will display the print operation's date/time.
        // If no printing has taken place, these fields will display "0/0/0000".
        FieldPrintDate field = (FieldPrintDate)doc.getRange().getFields().get(0);

        Assert.assertEquals("3/25/2020 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE ", field.getFieldCode());

        // Below are three different calendar types according to which the PRINTDATE field
        // can display the date and time of the last printing operation.
        // 1 -  Islamic Lunar Calendar:
        field = (FieldPrintDate)doc.getRange().getFields().get(1);

        Assert.assertTrue(field.getUseLunarCalendar());
        Assert.assertEquals("8/1/1441 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE  \\h", field.getFieldCode());

        field = (FieldPrintDate)doc.getRange().getFields().get(2);

        // 2 -  Umm al-Qura calendar:
        Assert.assertTrue(field.getUseUmAlQuraCalendar());
        Assert.assertEquals("8/1/1441 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE  \\u", field.getFieldCode());

        field = (FieldPrintDate)doc.getRange().getFields().get(3);

        // 3 -  Indian National Calendar:
        Assert.assertTrue(field.getUseSakaEraCalendar());
        Assert.assertEquals("1/5/1942 12:00:00 AM", field.getResult());
        Assert.assertEquals(" PRINTDATE  \\s", field.getFieldCode());
        //ExEnd
    }

    @Test
    public void fieldQuote() throws Exception
    {
        //ExStart
        //ExFor:FieldQuote
        //ExFor:FieldQuote.Text
        //ExFor:Document.UpdateFields
        //ExSummary:Shows to use the QUOTE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QUOTE field, which will display the value of its Text property.
        FieldQuote field = (FieldQuote)builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("\"Quoted text\"");

        Assert.assertEquals(" QUOTE  \"\\\"Quoted text\\\"\"", field.getFieldCode());

        // Insert a QUOTE field and nest a DATE field inside it.
        // DATE fields update their value to the current date every time we open the document using Microsoft Word.
        // Nesting the DATE field inside the QUOTE field like this will freeze its value
        // to the date when we created the document.
        builder.write("\nDocument creation date: ");
        field = (FieldQuote)builder.insertField(FieldType.FIELD_QUOTE, true);
        builder.moveTo(field.getSeparator());
        builder.insertField(FieldType.FIELD_DATE, true);

        Assert.assertEquals(" QUOTE \u0013 DATE \u0014" + new Date().getDate().toShortDateString() + "\u0015", field.getFieldCode());

        // Update all the fields to display their correct results.
        doc.updateFields();

        Assert.assertEquals("\"Quoted text\"", doc.getRange().getFields().get(0).getResult());

        doc.save(getArtifactsDir() + "Field.QUOTE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.QUOTE.docx");

        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE  \"\\\"Quoted text\\\"\"", "\"Quoted text\"", doc.getRange().getFields().get(0));

        TestUtil.verifyField(FieldType.FIELD_QUOTE, " QUOTE \u0013 DATE \u0014" + new Date().getDate().toShortDateString() + "\u0015", 
            new Date().getDate().toShortDateString(), doc.getRange().getFields().get(1));

    }

    //ExStart
    //ExFor:FieldNext
    //ExFor:FieldNextIf
    //ExFor:FieldNextIf.ComparisonOperator
    //ExFor:FieldNextIf.LeftExpression
    //ExFor:FieldNextIf.RightExpression
    //ExSummary:Shows how to use NEXT/NEXTIF fields to merge multiple rows into one page during a mail merge.
    @Test //ExSkip
    public void fieldNext() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a data source for our mail merge with 3 rows.
        // A mail merge that uses this table would normally create a 3-page document.
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");
        table.getRows().add("Mr.", "Joe", "Bloggs");

        insertMergeFields(builder, "First row: ");

        // If we have multiple merge fields with the same FieldName,
        // they will receive data from the same row of the data source and display the same value after the merge.
        // A NEXT field tells the mail merge instantly to move down one row,
        // which means any MERGEFIELDs that follow the NEXT field will receive data from the next row.
        // Make sure never to try to skip to the next row while already on the last row.
        FieldNext fieldNext = (FieldNext)builder.insertField(FieldType.FIELD_NEXT, true);

        Assert.assertEquals(" NEXT ", fieldNext.getFieldCode());

        // After the merge, the data source values that these MERGEFIELDs accept
        // will end up on the same page as the MERGEFIELDs above. 
        insertMergeFields(builder, "Second row: ");

        // A NEXTIF field has the same function as a NEXT field,
        // but it skips to the next row only if a statement constructed by the following 3 properties is true.
        FieldNextIf fieldNextIf = (FieldNextIf)builder.insertField(FieldType.FIELD_NEXT_IF, true);
        fieldNextIf.setLeftExpression("5");
        fieldNextIf.setRightExpression("2 + 3");
        fieldNextIf.setComparisonOperator("=");

        Assert.assertEquals(" NEXTIF  5 = \"2 + 3\"", fieldNextIf.getFieldCode());

        // If the comparison asserted by the above field is correct,
        // the following 3 merge fields will take data from the third row.
        // Otherwise, these fields will take data from row 2 again.
        insertMergeFields(builder, "Third row: ");

        doc.getMailMerge().execute(table);

        // Our data source has 3 rows, and we skipped rows twice. 
        // Our output document will have 1 page with data from all 3 rows.
        doc.save(getArtifactsDir() + "Field.NEXT.NEXTIF.docx");
        testFieldNext(doc); //ExSKip
    }

    /// <summary>
    /// Uses a document builder to insert MERGEFIELDs for a data source that contains columns named "Courtesy Title", "First Name" and "Last Name".
    /// </summary>
    @Test (enabled = false)
    public void insertMergeFields(DocumentBuilder builder, String firstFieldTextBefore) throws Exception
    {
        insertMergeField(builder, "Courtesy Title", firstFieldTextBefore, " ");
        insertMergeField(builder, "First Name", null, " ");
        insertMergeField(builder, "Last Name", null, null);
        builder.insertParagraph();
    }

    /// <summary>
    /// Uses a document builder to insert a MERRGEFIELD with specified properties.
    /// </summary>
    @Test (enabled = false)
    public void insertMergeField(DocumentBuilder builder, String fieldName, String textBefore, String textAfter) throws Exception
    {
        FieldMergeField field = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        field.setFieldName(fieldName);
        field.setTextBefore(textBefore);
        field.setTextAfter(textAfter);
    }
    //ExEnd

    private void testFieldNext(Document doc) throws Exception
    {
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
    //ExSummary:Shows to insert NOTEREF fields, and modify their appearance.
    @Test (enabled = false, description = "WORDSNET-17845") //ExSkip
    public void fieldNoteRef() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bookmark with a footnote that the NOTEREF field will reference.
        insertBookmarkWithFootnote(builder, "MyBookmark1", "Contents of MyBookmark1", "Footnote from MyBookmark1");

        // This NOTEREF field will display the number of the footnote inside the referenced bookmark.
        // Setting the InsertHyperlink property lets us jump to the bookmark by Ctrl + clicking the field in Microsoft Word.
        Assert.assertEquals(" NOTEREF  MyBookmark2 \\h",
            insertFieldNoteRef(builder, "MyBookmark2", true, false, false, "Hyperlink to Bookmark2, with footnote number ").getFieldCode());

        // When using the \p flag, after the footnote number, the field also displays the bookmark's position relative to the field.
        // Bookmark1 is above this field and contains footnote number 1, so the result will be "1 above" on update.
        Assert.assertEquals(" NOTEREF  MyBookmark1 \\h \\p",
            insertFieldNoteRef(builder, "MyBookmark1", true, true, false, "Bookmark1, with footnote number ").getFieldCode());

        // Bookmark2 is below this field and contains footnote number 2, so the field will display "2 below".
        // The \f flag makes the number 2 appear in the same format as the footnote number label in the actual text.
        Assert.assertEquals(" NOTEREF  MyBookmark2 \\h \\p \\f",
            insertFieldNoteRef(builder, "MyBookmark2", true, true, true, "Bookmark2, with footnote number ").getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        insertBookmarkWithFootnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.NOTEREF.docx");
        testNoteRef(new Document(getArtifactsDir() + "Field.NOTEREF.docx")); //ExSkip
    }

    /// <summary>
    /// Uses a document builder to insert a NOTEREF field with specified properties.
    /// </summary>
    private static FieldNoteRef insertFieldNoteRef(DocumentBuilder builder, String bookmarkName, boolean insertHyperlink, boolean insertRelativePosition, boolean insertReferenceMark, String textBefore) throws Exception
    {
        builder.write(textBefore);

        FieldNoteRef field = (FieldNoteRef)builder.insertField(FieldType.FIELD_NOTE_REF, true);
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
    private static void insertBookmarkWithFootnote(DocumentBuilder builder, String bookmarkName, String bookmarkText, String footnoteText)
    {
        builder.startBookmark(bookmarkName);
        builder.write(bookmarkText);
        builder.insertFootnote(FootnoteType.FOOTNOTE, footnoteText);
        builder.endBookmark(bookmarkName);
        builder.writeln();
    }
    //ExEnd

    private void testNoteRef(Document doc)
    {
        FieldNoteRef field = (FieldNoteRef)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark2 \\h", "2", field);
        Assert.assertEquals("MyBookmark2", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertFalse(field.getInsertRelativePosition());
        Assert.assertFalse(field.getInsertReferenceMark());

        field = (FieldNoteRef)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark1 \\h \\p", "1 above", field);
        Assert.assertEquals("MyBookmark1", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());
        Assert.assertFalse(field.getInsertReferenceMark());

        field = (FieldNoteRef)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_NOTE_REF, " NOTEREF  MyBookmark2 \\h \\p \\f", "2 below", field);
        Assert.assertEquals("MyBookmark2", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());
        Assert.assertTrue(field.getInsertReferenceMark());
    }

    @Test (enabled = false, description = "WORDSNET-17845")
    public void footnoteRef() throws Exception
    {
        //ExStart
        //ExFor:FieldFootnoteRef
        //ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("CrossRefBookmark");
        builder.write("Hello world!");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Cross referenced footnote.");
        builder.endBookmark("CrossRefBookmark");
        builder.insertParagraph();

        // Insert a FOOTNOTEREF field, which lets us reference a footnote more than once while re-using the same footnote marker.
        builder.write("CrossReference: ");
        FieldFootnoteRef field = (FieldFootnoteRef) builder.insertField(FieldType.FIELD_FOOTNOTE_REF, true);

        // Reference the bookmark that we have created with the FOOTNOTEREF field. That bookmark contains a footnote marker
        // belonging to the footnote we inserted. The field will display that footnote marker.
        builder.moveTo(field.getSeparator());
        builder.write("CrossRefBookmark");

        Assert.assertEquals(" FOOTNOTEREF CrossRefBookmark", field.getFieldCode());

        doc.updateFields();

        // This field works only in older versions of Microsoft Word.
        doc.save(getArtifactsDir() + "Field.FOOTNOTEREF.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.FOOTNOTEREF.doc");
        field = (FieldFootnoteRef)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_FOOTNOTE_REF, " FOOTNOTEREF CrossRefBookmark", "1", field);
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "Cross referenced footnote.", 
            (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
    }

    //ExStart
    //ExFor:FieldPageRef
    //ExFor:FieldPageRef.BookmarkName
    //ExFor:FieldPageRef.InsertHyperlink
    //ExFor:FieldPageRef.InsertRelativePosition
    //ExSummary:Shows to insert PAGEREF fields to display the relative location of bookmarks.
    @Test (enabled = false, description = "WORDSNET-17836") //ExSkip
    public void fieldPageRef() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        insertAndNameBookmark(builder, "MyBookmark1");

        // Insert a PAGEREF field that displays what page a bookmark is on.
        // Set the InsertHyperlink flag to make the field also function as a clickable link to the bookmark.
        Assert.assertEquals(" PAGEREF  MyBookmark3 \\h", 
            insertFieldPageRef(builder, "MyBookmark3", true, false, "Hyperlink to Bookmark3, on page: ").getFieldCode());

        // We can use the \p flag to get the PAGEREF field to display
        // the bookmark's position relative to the position of the field.
        // Bookmark1 is on the same page and above this field, so this field's displayed result will be "above".
        Assert.assertEquals(" PAGEREF  MyBookmark1 \\h \\p", 
            insertFieldPageRef(builder, "MyBookmark1", true, true, "Bookmark1 is ").getFieldCode());

        // Bookmark2 will be on the same page and below this field, so this field's displayed result will be "below".
        Assert.assertEquals(" PAGEREF  MyBookmark2 \\h \\p", 
            insertFieldPageRef(builder, "MyBookmark2", true, true, "Bookmark2 is ").getFieldCode());

        // Bookmark3 will be on a different page, so the field will display "on page 2".
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
    /// Uses a document builder to insert a PAGEREF field and sets its properties.
    /// </summary>
    private static FieldPageRef insertFieldPageRef(DocumentBuilder builder, String bookmarkName, boolean insertHyperlink, boolean insertRelativePosition, String textBefore) throws Exception
    {
        builder.write(textBefore);

        FieldPageRef field = (FieldPageRef)builder.insertField(FieldType.FIELD_PAGE_REF, true);
        field.setBookmarkName(bookmarkName);
        field.setInsertHyperlink(insertHyperlink);
        field.setInsertRelativePosition(insertRelativePosition);
        builder.writeln();
      
        return field;
    }

    /// <summary>
    /// Uses a document builder to insert a named bookmark.
    /// </summary>
    private static void insertAndNameBookmark(DocumentBuilder builder, String bookmarkName)
    {
        builder.startBookmark(bookmarkName);
        builder.writeln($"Contents of bookmark \"{bookmarkName}\".");
        builder.endBookmark(bookmarkName);
    }
    //ExEnd

    private void testPageRef(Document doc)
    {
        FieldPageRef field = (FieldPageRef)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark3 \\h", "2", field);
        Assert.assertEquals("MyBookmark3", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertFalse(field.getInsertRelativePosition());

        field = (FieldPageRef)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark1 \\h \\p", "above", field);
        Assert.assertEquals("MyBookmark1", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());

        field = (FieldPageRef)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF  MyBookmark2 \\h \\p", "below", field);
        Assert.assertEquals("MyBookmark2", field.getBookmarkName());
        Assert.assertTrue(field.getInsertHyperlink());
        Assert.assertTrue(field.getInsertRelativePosition());

        field = (FieldPageRef)doc.getRange().getFields().get(3);

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
    //ExSummary:Shows how to insert REF fields to reference bookmarks.
    @Test (enabled = false, description = "WORDSNET-18067") //ExSkip
    public void fieldRef() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("MyBookmark");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "MyBookmark footnote #1");
        builder.write("Text that will appear in REF field");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "MyBookmark footnote #2");
        builder.endBookmark("MyBookmark");
        builder.moveToDocumentStart();

        // We will apply a custom list format, where the amount of angle brackets indicates the list level we are currently at.
        builder.getListFormat().applyNumberDefault();
        builder.getListFormat().getListLevel().setNumberFormat("> \u0000");

        // Insert a REF field that will contain the text within our bookmark, act as a hyperlink, and clone the bookmark's footnotes.
        FieldRef field = insertFieldRef(builder, "MyBookmark", "", "\n");
        field.setIncludeNoteOrComment(true);
        field.setInsertHyperlink(true);

        Assert.assertEquals(" REF  MyBookmark \\f \\h", field.getFieldCode());

        // Insert a REF field, and display whether the referenced bookmark is above or below it.
        field = insertFieldRef(builder, "MyBookmark", "The referenced paragraph is ", " this field.\n");
        field.setInsertRelativePosition(true);

        Assert.assertEquals(" REF  MyBookmark \\p", field.getFieldCode());

        // Display the list number of the bookmark as it appears in the document.
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number is ", "\n");
        field.setInsertParagraphNumber(true);

        Assert.assertEquals(" REF  MyBookmark \\n", field.getFieldCode());

        // Display the bookmark's list number, but with non-delimiter characters, such as the angle brackets, omitted.
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's paragraph number, non-delimiters suppressed, is ", "\n");
        field.setInsertParagraphNumber(true);
        field.setSuppressNonDelimiters(true);

        Assert.assertEquals(" REF  MyBookmark \\n \\t", field.getFieldCode());

        // Move down one list level.
        builder.getListFormat().setListLevelNumber(builder.getListFormat().getListLevelNumber() + 1)/*Property++*/;
        builder.getListFormat().getListLevel().setNumberFormat(">> \u0001");

        // Display the list number of the bookmark and the numbers of all the list levels above it.
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's full context paragraph number is ", "\n");
        field.setInsertParagraphNumberInFullContext(true);

        Assert.assertEquals(" REF  MyBookmark \\w", field.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);

        // Display the list level numbers between this REF field, and the bookmark that it is referencing.
        field = insertFieldRef(builder, "MyBookmark", "The bookmark's relative paragraph number is ", "\n");
        field.setInsertParagraphNumberInRelativeContext(true);

        Assert.assertEquals(" REF  MyBookmark \\r", field.getFieldCode());

        // At the end of the document, the bookmark will show up as a list item here.
        builder.writeln("List level above bookmark");
        builder.getListFormat().setListLevelNumber(builder.getListFormat().getListLevelNumber() + 1)/*Property++*/;
        builder.getListFormat().getListLevel().setNumberFormat(">>> \u0002");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.REF.docx");
        testFieldRef(new Document(getArtifactsDir() + "Field.REF.docx")); //ExSkip
    }

    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after it.
    /// </summary>
    private static FieldRef insertFieldRef(DocumentBuilder builder, String bookmarkName, String textBefore, String textAfter) throws Exception
    {
        builder.write(textBefore);
        FieldRef field = (FieldRef)builder.insertField(FieldType.FIELD_REF, true);
        field.setBookmarkName(bookmarkName);
        builder.write(textAfter);
        return field;
    }
    //ExEnd

    private void testFieldRef(Document doc) throws Exception
    {
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "MyBookmark footnote #1", 
            (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));
        TestUtil.verifyFootnote(FootnoteType.FOOTNOTE, true, "", "MyBookmark footnote #2", 
            (Footnote)doc.getChild(NodeType.FOOTNOTE, 0, true));

        FieldRef field = (FieldRef)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\f \\h", 
            "\u0002 MyBookmark footnote #1\r" +
            "Text that will appear in REF field\u0002 MyBookmark footnote #2\r", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getIncludeNoteOrComment());
        Assert.assertTrue(field.getInsertHyperlink());

        field = (FieldRef)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\p", "below", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertRelativePosition());

        field = (FieldRef)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\n", ">>> i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumber());
        Assert.assertEquals(" REF  MyBookmark \\n", field.getFieldCode());
        Assert.assertEquals(">>> i", field.getResult());

        field = (FieldRef)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\n \\t", "i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumber());
        Assert.assertTrue(field.getSuppressNonDelimiters());

        field = (FieldRef)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\w", "> 4>> c>>> i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumberInFullContext());

        field = (FieldRef)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark \\r", ">> c>>> i", field);
        Assert.assertEquals("MyBookmark", field.getBookmarkName());
        Assert.assertTrue(field.getInsertParagraphNumberInRelativeContext());
    }

    @Test (enabled = false, description = "WORDSNET-18068")
    public void fieldRD() throws Exception
    {
        //ExStart
        //ExFor:FieldRD
        //ExFor:FieldRD.FileName
        //ExFor:FieldRD.IsPathRelative
        //ExSummary:Shows to use the RD field to create a table of contents entries from headings in other documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a table of contents,
        // and then add one entry for the table of contents on the following page.
        builder.insertField(FieldType.FIELD_TOC, true);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.getCurrentParagraph().getParagraphFormat().setStyleName("Heading 1");
        builder.writeln("TOC entry from within this document");

        // Insert an RD field, which references another local file system document in its FileName property.
        // The TOC will also now accept all headings from the referenced document as entries for its table.
        FieldRD field = (FieldRD)builder.insertField(FieldType.FIELD_REF_DOC, true);
        field.setFileName("ReferencedDocument.docx");
        field.isPathRelative(true);

        Assert.assertEquals(" RD  ReferencedDocument.docx \\f", field.getFieldCode());

        // Create the document that the RD field is referencing and insert a heading. 
        // This heading will show up as an entry in the TOC field in our first document.
        Document referencedDoc = new Document();
        DocumentBuilder refDocBuilder = new DocumentBuilder(referencedDoc);
        refDocBuilder.getCurrentParagraph().getParagraphFormat().setStyleName("Heading 1");
        refDocBuilder.writeln("TOC entry from referenced document");
        referencedDoc.save(getArtifactsDir() + "ReferencedDocument.docx");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.RD.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.RD.docx");

        FieldToc fieldToc = (FieldToc)doc.getRange().getFields().get(0);

        Assert.assertEquals("TOC entry from within this document\t\u0013 PAGEREF _Toc36149519 \\h \u00142\u0015\r" +
                        "TOC entry from referenced document\t1\r", fieldToc.getResult());

        FieldPageRef fieldPageRef = (FieldPageRef)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_PAGE_REF, " PAGEREF _Toc36149519 \\h ", "2", fieldPageRef);

        field = (FieldRD)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_REF_DOC, " RD  ReferencedDocument.docx \\f", "", field);
        Assert.assertEquals("ReferencedDocument.docx", field.getFileName());
        Assert.assertTrue(field.isPathRelative());
    }

    @Test
    public void skipIf() throws Exception
    {
        //ExStart
        //ExFor:FieldSkipIf
        //ExFor:FieldSkipIf.ComparisonOperator
        //ExFor:FieldSkipIf.LeftExpression
        //ExFor:FieldSkipIf.RightExpression
        //ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // Insert a SKIPIF field. If the current row of a mail merge operation fulfills the condition
        // which the expressions of this field state, then the mail merge operation aborts the current row,
        // discards the current merge document, and then immediately moves to the next row to begin the next merge document.
        FieldSkipIf fieldSkipIf = (FieldSkipIf) builder.insertField(FieldType.FIELD_SKIP_IF, true);

        // Move the builder to the SKIPIF field's separator so we can place a MERGEFIELD inside the SKIPIF field.
        builder.moveTo(fieldSkipIf.getSeparator());
        FieldMergeField fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Department");

        // The MERGEFIELD refers to the "Department" column in our data table. If a row from that table
        // has a value of "HR" in its "Department" column, then this row will fulfill the condition.
        fieldSkipIf.setLeftExpression("=");
        fieldSkipIf.setRightExpression("HR");

        // Add content to our document, create the data source, and execute the mail merge.
        builder.moveToDocumentEnd();
        builder.write("Dear ");
        fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Name");
        builder.writeln(", ");
        
        // This table has three rows, and one of them fulfills the condition of our SKIPIF field. 
        // The mail merge will produce two pages.
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Name");
        table.getColumns().add("Department");
        table.getRows().add("John Doe", "Sales");
        table.getRows().add("Jane Doe", "Accounting");
        table.getRows().add("John Cardholder", "HR");

        doc.getMailMerge().execute(table);
        doc.save(getArtifactsDir() + "Field.SKIPIF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SKIPIF.docx");

        Assert.assertEquals(0, doc.getRange().getFields().getCount());
        Assert.assertEquals("Dear John Doe, \r" +
                        "\fDear Jane Doe, \r\f", doc.getText());
    }
  
    @Test
    public void fieldSetRef() throws Exception
    {
        //ExStart
        //ExFor:FieldRef
        //ExFor:FieldRef.BookmarkName
        //ExFor:FieldSet
        //ExFor:FieldSet.BookmarkName
        //ExFor:FieldSet.BookmarkText
        //ExSummary:Shows how to create bookmarked text with a SET field, and then display it in the document using a REF field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Name bookmarked text with a SET field. 
        // This field refers to the "bookmark" not a bookmark structure that appears within the text, but a named variable.
        FieldSet fieldSet = (FieldSet)builder.insertField(FieldType.FIELD_SET, false);
        fieldSet.setBookmarkName("MyBookmark");
        fieldSet.setBookmarkText("Hello world!");
        fieldSet.update();

        Assert.assertEquals(" SET  MyBookmark \"Hello world!\"", fieldSet.getFieldCode());

        // Refer to the bookmark by name in a REF field and display its contents.
        FieldRef fieldRef = (FieldRef)builder.insertField(FieldType.FIELD_REF, true);
        fieldRef.setBookmarkName("MyBookmark");
        fieldRef.update();

        Assert.assertEquals(" REF  MyBookmark", fieldRef.getFieldCode());
        Assert.assertEquals("Hello world!", fieldRef.getResult());

        doc.save(getArtifactsDir() + "Field.SET.REF.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SET.REF.docx");

        Assert.assertEquals("Hello world!", doc.getRange().getBookmarks().get(0).getText());

        fieldSet = (FieldSet)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SET, " SET  MyBookmark \"Hello world!\"", "Hello world!", fieldSet);
        Assert.assertEquals("MyBookmark", fieldSet.getBookmarkName());
        Assert.assertEquals("Hello world!", fieldSet.getBookmarkText());

        TestUtil.verifyField(FieldType.FIELD_REF, " REF  MyBookmark", "Hello world!", fieldRef);
        Assert.assertEquals("Hello world!", fieldRef.getResult());
    }

    @Test (enabled = false, description = "WORDSNET-18137")
    public void fieldTemplate() throws Exception
    {
        //ExStart
        //ExFor:FieldTemplate
        //ExFor:FieldTemplate.IncludeFullPath
        //ExSummary:Shows how to use a TEMPLATE field to display the local file system location of a document's template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldTemplate field = (FieldTemplate)builder.insertField(FieldType.FIELD_TEMPLATE, false);
        Assert.assertEquals(" TEMPLATE ", field.getFieldCode());

        builder.writeln();
        field = (FieldTemplate)builder.insertField(FieldType.FIELD_TEMPLATE, false);
        field.setIncludeFullPath(true);

        Assert.assertEquals(" TEMPLATE  \\p", field.getFieldCode());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TEMPLATE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.TEMPLATE.docx");

        field = (FieldTemplate)doc.getRange().getFields().get(0);
        Assert.assertEquals(" TEMPLATE ", field.getFieldCode());
        Assert.assertEquals("Normal.dotm", field.getResult());

        field = (FieldTemplate)doc.getRange().getFields().get(1);
        Assert.assertEquals(" TEMPLATE  \\p", field.getFieldCode());
        Assert.assertTrue(field.getResult().endsWith("\\Microsoft\\Templates\\Normal.dotm"));

    }

    @Test
    public void fieldSymbol() throws Exception
    {
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

        // Below are three ways to use a SYMBOL field to display a single character.
        // 1 -  Add a SYMBOL field which displays the © (Copyright) symbol, specified by an ANSI character code:
        FieldSymbol field = (FieldSymbol)builder.insertField(FieldType.FIELD_SYMBOL, true);

        // The ANSI character code "U+00A9", or "169" in integer form, is reserved for the copyright symbol.
        field.setCharacterCode(Integer.toString(0x00a9));
        field.isAnsi(true);

        Assert.assertEquals(" SYMBOL  169 \\a", field.getFieldCode());

        builder.writeln(" Line 1");

        // 2 -  Add a SYMBOL field which displays the ∞ (Infinity) symbol, and modify its appearance:
        field = (FieldSymbol)builder.insertField(FieldType.FIELD_SYMBOL, true);

        // In Unicode, the infinity symbol occupies the "221E" code.
        field.setCharacterCode(Integer.toString(0x221E));
        field.isUnicode(true);

        // Change the font of our symbol after using the Windows Character Map
        // to ensure that the font can represent that symbol.
        field.setFontName("Calibri");
        field.setFontSize("24");

        // We can set this flag for tall symbols to make them not push down the rest of the text on their line.
        field.setDontAffectsLineSpacing(true);

        Assert.assertEquals(" SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", field.getFieldCode());

        builder.writeln("Line 2");

        // 3 -  Add a SYMBOL field which displays the あ character,
        // with a font that supports Shift-JIS (Windows-932) codepage:
        field = (FieldSymbol)builder.insertField(FieldType.FIELD_SYMBOL, true);
        field.setFontName("MS Gothic");
        field.setCharacterCode(Integer.toString(0x82A0));
        field.isShiftJis(true);

        Assert.assertEquals(" SYMBOL  33440 \\f \"MS Gothic\" \\j", field.getFieldCode());

        builder.write("Line 3");

        doc.save(getArtifactsDir() + "Field.SYMBOL.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.SYMBOL.docx");

        field = (FieldSymbol)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL  169 \\a", "", field);
        Assert.assertEquals(Integer.toString(0x00a9), field.getCharacterCode());
        Assert.assertTrue(field.isAnsi());
        Assert.assertEquals("©", field.getDisplayResult());
            
        field = (FieldSymbol)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL  8734 \\u \\f Calibri \\s 24 \\h", "", field);
        Assert.assertEquals(Integer.toString(0x221E), field.getCharacterCode());
        Assert.assertEquals("Calibri", field.getFontName());
        Assert.assertEquals("24", field.getFontSize());
        Assert.assertTrue(field.isUnicode());
        Assert.assertTrue(field.getDontAffectsLineSpacing());
        Assert.assertEquals("∞", field.getDisplayResult());

        field = (FieldSymbol)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_SYMBOL, " SYMBOL  33440 \\f \"MS Gothic\" \\j", "", field);
        Assert.assertEquals(Integer.toString(0x82A0), field.getCharacterCode());
        Assert.assertEquals("MS Gothic", field.getFontName());
        Assert.assertTrue(field.isShiftJis());
    }

    @Test
    public void fieldTitle() throws Exception
    {
        //ExStart
        //ExFor:FieldTitle
        //ExFor:FieldTitle.Text
        //ExSummary:Shows how to use the TITLE field.
        Document doc = new Document();

        // Set a value for the "Title" built-in document property. 
        doc.getBuiltInDocumentProperties().setTitle("My Title");

        // We can use the TITLE field to display the value of this property in the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldTitle field = (FieldTitle)builder.insertField(FieldType.FIELD_TITLE, false);
        field.update();

        Assert.assertEquals(" TITLE ", field.getFieldCode());
        Assert.assertEquals("My Title", field.getResult());

        // Setting a value for the field's Text property,
        // and then updating the field will also overwrite the corresponding built-in property with the new value.
        builder.writeln();
        field = (FieldTitle)builder.insertField(FieldType.FIELD_TITLE, false);
        field.setText("My New Title");
        field.update();

        Assert.assertEquals(" TITLE  \"My New Title\"", field.getFieldCode());
        Assert.assertEquals("My New Title", field.getResult());
        Assert.assertEquals("My New Title", doc.getBuiltInDocumentProperties().getTitle());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TITLE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.TITLE.docx");

        Assert.assertEquals("My New Title", doc.getBuiltInDocumentProperties().getTitle());

        field = (FieldTitle)doc.getRange().getFields().get(0);

        TestUtil.verifyField(FieldType.FIELD_TITLE, " TITLE ", "My New Title", field);

        field = (FieldTitle)doc.getRange().getFields().get(1);

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
    public void fieldTOA() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOA field, which will create an entry for each TA field in the document,
        // displaying long citations and page numbers for each entry.
        FieldToa fieldToa = (FieldToa)builder.insertField(FieldType.FIELD_TOA, false);

        // Set the entry category for our table. This TOA will now only include TA fields
        // that have a matching value in their EntryCategory property.
        fieldToa.setEntryCategory("1");

        // Moreover, the Table of Authorities category at index 1 is "Cases",
        // which will show up as our table's title if we set this variable to true.
        fieldToa.setUseHeading(true);

        // We can further filter TA fields by naming a bookmark that they will need to be within the TOA bounds.
        fieldToa.setBookmarkName("MyBookmark");

        // By default, a dotted line page-wide tab appears between the TA field's citation
        // and its page number. We can replace it with any text we put on this property.
        // Inserting a tab character will preserve the original tab.
        fieldToa.setEntrySeparator(" \t p.");

        // If we have multiple TA entries that share the same long citation,
        // all their respective page numbers will show up on one row.
        // We can use this property to specify a string that will separate their page numbers.
        fieldToa.setPageNumberListSeparator(" & p. ");

        // We can set this to true to get our table to display the word "passim"
        // if there are five or more page numbers in one row.
        fieldToa.setUsePassim(true);

        // One TA field can refer to a range of pages.
        // We can specify a string here to appear between the start and end page numbers for such ranges.
        fieldToa.setPageRangeSeparator(" to ");

        // The format from the TA fields will carry over into our table.
        // We can disable this by setting the RemoveEntryFormatting flag.
        fieldToa.setRemoveEntryFormatting(true);
        builder.getFont().setColor(msColor.getGreen());
        builder.getFont().setName("Arial Black");

        Assert.assertEquals(" TOA  \\c 1 \\h \\b MyBookmark \\e \" \t p.\" \\l \" & p. \" \\p \\g \" to \" \\f", fieldToa.getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);

        // This TA field will not appear as an entry in the TOA since it is outside
        // the bookmark's bounds that the TOA's BookmarkName property specifies.
        FieldTA fieldTA = insertToaEntry(builder, "1", "Source 1");

        Assert.assertEquals(" TA  \\c 1 \\l \"Source 1\"", fieldTA.getFieldCode());

        // This TA field is inside the bookmark,
        // but the entry category does not match that of the table, so the TA field will not include it.
        builder.startBookmark("MyBookmark");
        fieldTA = insertToaEntry(builder, "2", "Source 2");

        // This entry will appear in the table.
        fieldTA = insertToaEntry(builder, "1", "Source 3");

        // A TOA table does not display short citations,
        // but we can use them as a shorthand to refer to bulky source names that multiple TA fields reference.
        fieldTA.setShortCitation("S.3");

        Assert.assertEquals(" TA  \\c 1 \\l \"Source 3\" \\s S.3", fieldTA.getFieldCode());

        // We can format the page number to make it bold/italic using the following properties.
        // We will still see these effects if we set our table to ignore formatting.
        fieldTA = insertToaEntry(builder, "1", "Source 2");
        fieldTA.isBold(true);
        fieldTA.isItalic(true);

        Assert.assertEquals(" TA  \\c 1 \\l \"Source 2\" \\b \\i", fieldTA.getFieldCode());

        // We can configure TA fields to get their TOA entries to refer to a range of pages that a bookmark spans across.
        // Note that this entry refers to the same source as the one above to share one row in our table.
        // This row will have the page number of the entry above and the page range of this entry,
        // with the table's page list and page number range separators between page numbers.
        fieldTA = insertToaEntry(builder, "1", "Source 3");
        fieldTA.setPageRangeBookmarkName("MyMultiPageBookmark");

        builder.startBookmark("MyMultiPageBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.endBookmark("MyMultiPageBookmark");

        Assert.assertEquals(" TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", fieldTA.getFieldCode());

        // If we have enabled the "Passim" feature of our table, having 5 or more TA entries with the same source will invoke it.
        for (int i = 0; i < 5; i++)
        {
            insertToaEntry(builder, "1", "Source 4");
        }

        builder.endBookmark("MyBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOA.TA.docx");
        testFieldTOA(new Document(getArtifactsDir() + "Field.TOA.TA.docx")); //ExSKip
    }

    private static FieldTA insertToaEntry(DocumentBuilder builder, String entryCategory, String longCitation) throws Exception
    {
        FieldTA field = (FieldTA)builder.insertField(FieldType.FIELD_TOA_ENTRY, false);
        field.setEntryCategory(entryCategory);
        field.setLongCitation(longCitation);

        builder.insertBreak(BreakType.PAGE_BREAK);

        return field;
    }
    //ExEnd

    private void testFieldTOA(Document doc)
    {
        FieldToa fieldTOA = (FieldToa)doc.getRange().getFields().get(0);

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

        FieldTA fieldTA = (FieldTA)doc.getRange().getFields().get(1);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 1\"", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 1", fieldTA.getLongCitation());

        fieldTA = (FieldTA)doc.getRange().getFields().get(2);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 2 \\l \"Source 2\"", "", fieldTA);
        Assert.assertEquals("2", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 2", fieldTA.getLongCitation());

        fieldTA = (FieldTA)doc.getRange().getFields().get(3);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 3\" \\s S.3", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 3", fieldTA.getLongCitation());
        Assert.assertEquals("S.3", fieldTA.getShortCitation());

        fieldTA = (FieldTA)doc.getRange().getFields().get(4);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 2\" \\b \\i", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 2", fieldTA.getLongCitation());
        Assert.assertTrue(fieldTA.isBold());
        Assert.assertTrue(fieldTA.isItalic());

        fieldTA = (FieldTA)doc.getRange().getFields().get(5);

        TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 3\" \\r MyMultiPageBookmark", "", fieldTA);
        Assert.assertEquals("1", fieldTA.getEntryCategory());
        Assert.assertEquals("Source 3", fieldTA.getLongCitation());
        Assert.assertEquals("MyMultiPageBookmark", fieldTA.getPageRangeBookmarkName());

        for (int i = 6; i < 11; i++)
        {
            fieldTA = (FieldTA)doc.getRange().getFields().get(i);

            TestUtil.verifyField(FieldType.FIELD_TOA_ENTRY, " TA  \\c 1 \\l \"Source 4\"", "", fieldTA);
            Assert.assertEquals("1", fieldTA.getEntryCategory());
            Assert.assertEquals("Source 4", fieldTA.getLongCitation());
        }
    }

    @Test
    public void fieldAddIn() throws Exception
    {
        //ExStart
        //ExFor:FieldAddIn
        //ExSummary:Shows how to process an ADDIN field.
        Document doc = new Document(getMyDir() + "Field sample - ADDIN.docx");

        // Aspose.Words does not support inserting ADDIN fields, but we can still load and read them.
        FieldAddIn field = (FieldAddIn)doc.getRange().getFields().get(0);

        Assert.assertEquals(" ADDIN \"My value\" ", field.getFieldCode());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.verifyField(FieldType.FIELD_ADDIN, " ADDIN \"My value\" ", "", doc.getRange().getFields().get(0));
    }

    @Test
    public void fieldEditTime() throws Exception
    {
        //ExStart
        //ExFor:FieldEditTime
        //ExSummary:Shows how to use the EDITTIME field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The EDITTIME field will show, in minutes,
        // the time spent with the document open in a Microsoft Word window.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("You've been editing this document for ");
        FieldEditTime field = (FieldEditTime)builder.insertField(FieldType.FIELD_EDIT_TIME, true);
        builder.writeln(" minutes.");
        
        // This built in document property tracks the minutes. Microsoft Word uses this property
        // to track the time spent with the document open. We can also edit it ourselves.
        doc.getBuiltInDocumentProperties().setTotalEditingTime(10);
        field.update();

        Assert.assertEquals(" EDITTIME ", field.getFieldCode());
        Assert.assertEquals("10", field.getResult());

        // The field does not update itself in real-time, and will also have to be
        // manually updated in Microsoft Word anytime we need an accurate value.
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
    public void fieldEQ() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // An EQ field displays a mathematical equation consisting of one or many elements.
        // Each element takes the following form: [switch][options][arguments].
        // There may be one switch, and several possible options.
        // The arguments are a set of coma-separated values enclosed by round braces.

        // Here we use a document builder to insert an EQ field, with an "\f" switch, which corresponds to "Fraction".
        // We will pass values 1 and 4 as arguments, and we will not use any options.
        // This field will display a fraction with 1 as the numerator and 4 as the denominator.
        FieldEQ field = insertFieldEQ(builder, "\\f(1,4)");

        Assert.assertEquals(" EQ \\f(1,4)", field.getFieldCode());

        // One EQ field may contain multiple elements placed sequentially.
        // We can also nest elements inside one another by placing the inner elements
        // inside the argument brackets of outer elements.
        // We can find the full list of switches, along with their uses here:
        // https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/

        // Below are applications of nine different EQ field switches that we can use to create different kinds of objects. 
        // 1 -  Array switch "\a", aligned left, 2 columns, 3 points of horizontal and vertical spacing:
        insertFieldEQ(builder, "\\a \\al \\co2 \\vs3 \\hs3(4x,- 4y,-4x,+ y)");

        // 2 -  Bracket switch "\b", bracket character "[", to enclose the contents in a set of square braces:
        // Note that we are nesting an array inside the brackets, which will altogether look like a matrix in the output.
        insertFieldEQ(builder, "\\b \\bc\\[ (\\a \\al \\co3 \\vs3 \\hs3(1,0,0,0,1,0,0,0,1))");

        // 3 -  Displacement switch "\d", displacing text "B" 30 spaces to the right of "A", displaying the gap as an underline:
        insertFieldEQ(builder, "A \\d \\fo30 \\li() B");

        // 4 -  Formula consisting of multiple fractions:
        insertFieldEQ(builder, "\\f(d,dx)(u + v) = \\f(du,dx) + \\f(dv,dx)");

        // 5 -  Integral switch "\i", with a summation symbol:
        insertFieldEQ(builder, "\\i \\su(n=1,5,n)");

        // 6 -  List switch "\l":
        insertFieldEQ(builder, "\\l(1,1,2,3,n,8,13)");

        // 7 -  Radical switch "\r", displaying a cubed root of x:
        insertFieldEQ(builder, "\\r (3,x)");

        // 8 -  Subscript/superscript switch "/s", first as a superscript and then as a subscript:
        insertFieldEQ(builder, "\\s \\up8(Superscript) Text \\s \\do8(Subscript)");

        // 9 -  Box switch "\x", with lines at the top, bottom, left and right of the input:
        insertFieldEQ(builder, "\\x \\to \\bo \\le \\ri(5)");

        // Some more complex combinations.
        insertFieldEQ(builder, "\\a \\ac \\vs1 \\co1(lim,n→∞) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))");
        insertFieldEQ(builder, "\\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)");
        insertFieldEQ(builder, "\\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)");

        doc.save(getArtifactsDir() + "Field.EQ.docx");
        testFieldEQ(new Document(getArtifactsDir() + "Field.EQ.docx")); //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph.
    /// </summary>
    private static FieldEQ insertFieldEQ(DocumentBuilder builder, String args) throws Exception
    {
        FieldEQ field = (FieldEQ)builder.insertField(FieldType.FIELD_EQUATION, true);
        builder.moveTo(field.getSeparator());
        builder.write(args);
        builder.moveTo(field.getStart().getParentNode());
        
        builder.insertParagraph();
        return field;
    }
    //ExEnd

    private void testFieldEQ(Document doc)
    {
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
        TestUtil.verifyWebResponseStatusCode(HttpStatusCode.OK, "https://blogs.msdn.microsoft.com/murrays/2018/01/23/microsoft-word-eq-field/");
    }

    @Test
    public void fieldForms() throws Exception
    {
        //ExStart
        //ExFor:FieldFormCheckBox
        //ExFor:FieldFormDropDown
        //ExFor:FieldFormText
        //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
        // These fields are legacy equivalents of the FormField. We can read, but not create these fields using Aspose.Words.
        // In Microsoft Word, we can insert these fields via the Legacy Tools menu in the Developer tab.
        Document doc = new Document(getMyDir() + "Form fields.docx");

        FieldFormCheckBox fieldFormCheckBox = (FieldFormCheckBox)doc.getRange().getFields().get(1);
        Assert.assertEquals(" FORMCHECKBOX \u0001", fieldFormCheckBox.getFieldCode());

        FieldFormDropDown fieldFormDropDown = (FieldFormDropDown)doc.getRange().getFields().get(2);
        Assert.assertEquals(" FORMDROPDOWN \u0001", fieldFormDropDown.getFieldCode());

        FieldFormText fieldFormText = (FieldFormText)doc.getRange().getFields().get(0);
        Assert.assertEquals(" FORMTEXT \u0001", fieldFormText.getFieldCode());
        //ExEnd
    }

    @Test
    public void fieldFormula() throws Exception
    {
        //ExStart
        //ExFor:FieldFormula
        //ExSummary:Shows how to use the formula field to display the result of an equation.
        Document doc = new Document();

        // Use a field builder to construct a mathematical equation,
        // then create a formula field to display the equation's result in the document.
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_FORMULA);
        fieldBuilder.addArgument(2);
        fieldBuilder.addArgument("*");
        fieldBuilder.addArgument(5);

        FieldFormula field = (FieldFormula)fieldBuilder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph());
        field.update();

        Assert.assertEquals(" = 2 * 5 ", field.getFieldCode());
        Assert.assertEquals("10", field.getResult());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FORMULA.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.FORMULA.docx");

        TestUtil.verifyField(FieldType.FIELD_FORMULA, " = 2 * 5 ", "10", doc.getRange().getFields().get(0));
    }

    @Test
    public void fieldLastSavedBy() throws Exception
    {
        //ExStart
        //ExFor:FieldLastSavedBy
        //ExSummary:Shows how to use the LASTSAVEDBY field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we create a document in Microsoft Word, it will have the user's name in the "Last saved by" built-in property.
        // If we make a document programmatically, this property will be null, and we will need to assign a value. 
        doc.getBuiltInDocumentProperties().setLastSavedBy("John Doe");

        // We can use the LASTSAVEDBY field to display the value of this property in the document.
        FieldLastSavedBy field = (FieldLastSavedBy)builder.insertField(FieldType.FIELD_LAST_SAVED_BY, true);

        Assert.assertEquals(" LASTSAVEDBY ", field.getFieldCode());
        Assert.assertEquals("John Doe", field.getResult());

        doc.save(getArtifactsDir() + "Field.LASTSAVEDBY.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.LASTSAVEDBY.docx");

        Assert.assertEquals("John Doe", doc.getBuiltInDocumentProperties().getLastSavedBy());
        TestUtil.verifyField(FieldType.FIELD_LAST_SAVED_BY, " LASTSAVEDBY ", "John Doe", doc.getRange().getFields().get(0));
    }

    @Test (enabled = false, description = "WORDSNET-18173")
    public void fieldMergeRec() throws Exception
    {
        //ExStart
        //ExFor:FieldMergeRec
        //ExFor:FieldMergeSeq
        //ExFor:FieldSkipIf
        //ExFor:FieldSkipIf.ComparisonOperator
        //ExFor:FieldSkipIf.LeftExpression
        //ExFor:FieldSkipIf.RightExpression
        //ExSummary:Shows how to use MERGEREC and MERGESEQ fields to the number and count mail merge records in a mail merge's output documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        FieldMergeField fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Name");
        builder.writeln(",");

        // A MERGEREC field will print the row number of the data being merged in every merge output document.
        builder.write("\nRow number of record in data source: ");
        FieldMergeRec fieldMergeRec = (FieldMergeRec)builder.insertField(FieldType.FIELD_MERGE_REC, true);

        Assert.assertEquals(" MERGEREC ", fieldMergeRec.getFieldCode());

        // A MERGESEQ field will count the number of successful merges and print the current value on each respective page.
        // If a mail merge skips no rows and invokes no SKIP/SKIPIF/NEXT/NEXTIF fields, then all merges are successful.
        // The MERGESEQ and MERGEREC fields will display the same results of their mail merge was successful.
        builder.write("\nSuccessful merge number: ");
        FieldMergeSeq fieldMergeSeq = (FieldMergeSeq)builder.insertField(FieldType.FIELD_MERGE_SEQ, true);

        Assert.assertEquals(" MERGESEQ ", fieldMergeSeq.getFieldCode());

        // Insert a SKIPIF field, which will skip a merge if the name is "John Doe".
        FieldSkipIf fieldSkipIf = (FieldSkipIf)builder.insertField(FieldType.FIELD_SKIP_IF, true);
        builder.moveTo(fieldSkipIf.getSeparator());
        fieldMergeField = (FieldMergeField)builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Name");
        fieldSkipIf.setLeftExpression("=");
        fieldSkipIf.setRightExpression("John Doe");

        // Create a data source with 3 rows, one of them having "John Doe" as a value for the "Name" column.
        // Since a SKIPIF field will be triggered once by that value, the output of our mail merge will have 2 pages instead of 3.
        // On page 1, the MERGESEQ and MERGEREC fields will both display "1".
        // On page 2, the MERGEREC field will display "3" and the MERGESEQ field will display "2".
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Name");
        table.getRows().add(new String[] { "Jane Doe" });
        table.getRows().add(new String[] { "John Doe" });
        table.getRows().add(new String[] { "Joe Bloggs" });

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
    public void fieldOcx() throws Exception
    {
        //ExStart
        //ExFor:FieldOcx
        //ExSummary:Shows how to insert an OCX field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldOcx field = (FieldOcx)builder.insertField(FieldType.FIELD_OCX, true);

        Assert.assertEquals(" OCX ", field.getFieldCode());
        //ExEnd

        TestUtil.verifyField(FieldType.FIELD_OCX, " OCX ", "", field);
    }

    //ExStart
    //ExFor:Field.Remove
    //ExFor:FieldPrivate
    //ExSummary:Shows how to process PRIVATE fields.
    @Test //ExSkip
    public void fieldPrivate() throws Exception
    {
        // Open a Corel WordPerfect document which we have converted to .docx format.
        Document doc = new Document(getMyDir() + "Field sample - PRIVATE.docx");

        // WordPerfect 5.x/6.x documents like the one we have loaded may contain PRIVATE fields.
        // Microsoft Word preserves PRIVATE fields during load/save operations,
        // but provides no functionality for them.
        FieldPrivate field = (FieldPrivate)doc.getRange().getFields().get(0);

        Assert.assertEquals(" PRIVATE \"My value\" ", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_PRIVATE, field.getType());

        // We can also insert PRIVATE fields using a document builder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField(FieldType.FIELD_PRIVATE, true);

        // These fields are not a viable way of protecting sensitive information.
        // Unless backward compatibility with older versions of WordPerfect is essential,
        // we can safely remove these fields. We can do this using a DocumentVisiitor implementation.
        Assert.assertEquals(2, doc.getRange().getFields().getCount());

        FieldPrivateRemover remover = new FieldPrivateRemover();
        doc.accept(remover);

        Assert.assertEquals(2, remover.getFieldsRemovedCount());
        Assert.assertEquals(0, doc.getRange().getFields().getCount());
    }

    /// <summary>
    /// Removes all encountered PRIVATE fields.
    /// </summary>
    public static class FieldPrivateRemover extends DocumentVisitor
    {
        public FieldPrivateRemover()
        {
            mFieldsRemovedCount = 0;
        }

        public int getFieldsRemovedCount()
        {
            return mFieldsRemovedCount;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// If the node belongs to a PRIVATE field, the entire field is removed.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldEnd(FieldEnd fieldEnd) throws Exception
        {
            if (fieldEnd.getFieldType() == FieldType.FIELD_PRIVATE)
            {
                fieldEnd.getField().remove();
                mFieldsRemovedCount++;
            }

            return VisitorAction.CONTINUE;
        }

        private int mFieldsRemovedCount;
    }
    //ExEnd

    @Test
    public void fieldSection() throws Exception
    {
        //ExStart
        //ExFor:FieldSection
        //ExFor:FieldSectionPages
        //ExSummary:Shows how to use SECTION and SECTIONPAGES fields to number pages by sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        // A SECTION field displays the number of the section it is in.
        builder.write("Section ");
        FieldSection fieldSection = (FieldSection)builder.insertField(FieldType.FIELD_SECTION, true);

        Assert.assertEquals(" SECTION ", fieldSection.getFieldCode());

        // A PAGE field displays the number of the page it is in.
        builder.write("\nPage ");
        FieldPage fieldPage = (FieldPage)builder.insertField(FieldType.FIELD_PAGE, true);

        Assert.assertEquals(" PAGE ", fieldPage.getFieldCode());

        // A SECTIONPAGES field displays the number of pages that the section it is in spans across.
        builder.write(" of ");
        FieldSectionPages fieldSectionPages = (FieldSectionPages)builder.insertField(FieldType.FIELD_SECTION_PAGES, true);

        Assert.assertEquals(" SECTIONPAGES ", fieldSectionPages.getFieldCode());

        // Move out of the header back into the main document and insert two pages.
        // All these pages will be in the first section. Our fields, which appear once every header,
        // will number the current/total pages of this section.
        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // We can insert a new section with the document builder like this.
        // This will affect the values displayed in the SECTION and SECTIONPAGES fields in all upcoming headers.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        // The PAGE field will keep counting pages across the whole document.
        // We can manually reset its count at each section to keep track of pages section-by-section.
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
    public void fieldTime() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default, time is displayed in the "h:mm am/pm" format.
        FieldTime field = insertFieldTime(builder, "");

        Assert.assertEquals(" TIME ", field.getFieldCode());

        // We can use the \@ flag to change the format of our displayed time.
        field = insertFieldTime(builder, "\\@ HHmm");

        Assert.assertEquals(" TIME \\@ HHmm", field.getFieldCode());

        // We can adjust the format to get TIME field to also display the date, according to the Gregorian calendar.
        field = insertFieldTime(builder, "\\@ \"M/d/yyyy h mm:ss am/pm\"");

        Assert.assertEquals(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.getFieldCode());

        doc.save(getArtifactsDir() + "Field.TIME.docx");
        testFieldTime(new Document(getArtifactsDir() + "Field.TIME.docx")); //ExSkip
    }

    /// <summary>
    /// Use a document builder to insert a TIME field, insert a new paragraph and return the field.
    /// </summary>
    private static FieldTime insertFieldTime(DocumentBuilder builder, String format) throws Exception
    {
        FieldTime field = (FieldTime)builder.insertField(FieldType.FIELD_TIME, true);
        builder.moveTo(field.getSeparator());
        builder.write(format);
        builder.moveTo(field.getStart().getParentNode());

        builder.insertParagraph();
        return field;
    }
    //ExEnd

    private void testFieldTime(Document doc) throws Exception
    {
        DateTime docLoadingTime = new Date();
        doc = DocumentHelper.saveOpen(doc);

        FieldTime field = (FieldTime)doc.getRange().getFields().get(0);

        Assert.assertEquals(" TIME ", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_TIME, field.getType());
        Assert.assertEquals(DateTime.parse(field.getResult()), DateTime.getToday().addHours(docLoadingTime.getHour()).addMinutes(docLoadingTime.getMinute()));

        field = (FieldTime)doc.getRange().getFields().get(1);

        Assert.assertEquals(" TIME \\@ HHmm", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_TIME, field.getType());
        Assert.assertEquals(DateTime.parse(field.getResult()), DateTime.getToday().addHours(docLoadingTime.getHour()).addMinutes(docLoadingTime.getMinute()));

        field = (FieldTime)doc.getRange().getFields().get(2);

        Assert.assertEquals(" TIME \\@ \"M/d/yyyy h mm:ss am/pm\"", field.getFieldCode());
        Assert.assertEquals(FieldType.FIELD_TIME, field.getType());
        Assert.assertEquals(DateTime.parse(field.getResult()), DateTime.getToday().addHours(docLoadingTime.getHour()).addMinutes(docLoadingTime.getMinute()));
    }

    @Test
    public void bidiOutline() throws Exception
    {
        //ExStart
        //ExFor:FieldBidiOutline
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExFor:ParagraphFormat.Bidi
        //ExSummary:Shows how to create right-to-left language-compatible lists with BIDIOUTLINE fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The BIDIOUTLINE field numbers paragraphs like the AUTONUM/LISTNUM fields,
        // but is only visible when a right-to-left editing language is enabled, such as Hebrew or Arabic.
        // The following field will display ".1", the RTL equivalent of list number "1.".
        FieldBidiOutline field = (FieldBidiOutline)builder.insertField(FieldType.FIELD_BIDI_OUTLINE, true);
        builder.writeln("שלום");

        Assert.assertEquals(" BIDIOUTLINE ", field.getFieldCode());

        // Add two more BIDIOUTLINE fields, which will display ".2" and ".3".
        builder.insertField(FieldType.FIELD_BIDI_OUTLINE, true);
        builder.writeln("שלום");
        builder.insertField(FieldType.FIELD_BIDI_OUTLINE, true);
        builder.writeln("שלום");

        // Set the horizontal text alignment for every paragraph in the document to RTL.
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().setBidi(true);
        }

        // If we enable a right-to-left editing language in Microsoft Word, our fields will display numbers.
        // Otherwise, they will display "###".
        doc.save(getArtifactsDir() + "Field.BIDIOUTLINE.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Field.BIDIOUTLINE.docx");

        for (Field fieldBidiOutline : doc.getRange().getFields())
            TestUtil.verifyField(FieldType.FIELD_BIDI_OUTLINE, " BIDIOUTLINE ", "", fieldBidiOutline);
    }

    @Test
    public void legacy() throws Exception
    {
        //ExStart
        //ExFor:FieldEmbed
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled during loading.
        // Open a document that was created in Microsoft Word 2003.
        Document doc = new Document(getMyDir() + "Legacy fields.doc");

        // If we open the Word document and press Alt+F9, we will see a SHAPE and an EMBED field.
        // A SHAPE field is the anchor/canvas for an AutoShape object with the "In line with text" wrapping style enabled.
        // An EMBED field has the same function, but for an embedded object,
        // such as a spreadsheet from an external Excel document.
        // However, these fields will not appear in the document's Fields collection.
        Assert.assertEquals(0, doc.getRange().getFields().getCount());

        // These fields are supported only by old versions of Microsoft Word.
        // The document loading process will convert these fields into Shape objects,
        // which we can access in the document's node collection.
        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        Assert.assertEquals(3, shapes.getCount());

        // The first Shape node corresponds to the SHAPE field in the input document,
        // which is the inline canvas for the AutoShape.
        Shape shape = (Shape)shapes.get(0);
        Assert.assertEquals(ShapeType.IMAGE, shape.getShapeType());

        // The second Shape node is the AutoShape itself.
        shape = (Shape)shapes.get(1);
        Assert.assertEquals(ShapeType.CAN, shape.getShapeType());

        // The third Shape is what was the EMBED field that contained the external spreadsheet.
        shape = (Shape)shapes.get(2);
        Assert.assertEquals(ShapeType.OLE_OBJECT, shape.getShapeType());
        //ExEnd
    }

    @Test
    public void setFieldIndexFormat() throws Exception
    {
        //ExStart
        //ExFor:FieldOptions.FieldIndexFormat
        //ExSummary:Shows how to formatting FieldIndex fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("A");
        builder.insertBreak(BreakType.LINE_BREAK);
        builder.insertField("XE \"A\"");
        builder.write("B");

        builder.insertField(" INDEX \\e \" · \" \\h \"A\" \\c \"2\" \\z \"1033\"", null);

        doc.getFieldOptions().setFieldIndexFormat(FieldIndexFormat.FANCY);
        doc.updateFields();

        doc.save(getArtifactsDir() + "Field.SetFieldIndexFormat.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:ComparisonEvaluationResult.#ctor(bool)
    //ExFor:ComparisonEvaluationResult.#ctor(string)
    //ExFor:ComparisonEvaluationResult
    //ExFor:ComparisonExpression
    //ExFor:ComparisonExpression.LeftExpression
    //ExFor:ComparisonExpression.ComparisonOperator
    //ExFor:ComparisonExpression.RightExpression
    //ExFor:FieldOptions.ComparisonExpressionEvaluator
    //ExSummary:Shows how to implement custom evaluation for the IF and COMPARE fields.
    @Test (dataProvider = "conditionEvaluationExtensionPointDataProvider") //ExSkip
    public void conditionEvaluationExtensionPoint(String fieldCode, byte comparisonResult, String comparisonError,
        String expectedResult) throws Exception
    {
        final String LEFT = "\"left expression\"";
        final String _OPERATOR = "<>";
        final String RIGHT = "\"right expression\"";

        DocumentBuilder builder = new DocumentBuilder();

        // Field codes that we use in this example:
        // 1.   " IF {0} {1} {2} \"true argument\" \"false argument\" ".
        // 2.   " COMPARE {0} {1} {2} ".
        Field field = builder.insertField(msString.format(fieldCode, LEFT, _OPERATOR, RIGHT), null);

        // If the "comparisonResult" is undefined, we create "ComparisonEvaluationResult" with string, instead of bool.
        ComparisonEvaluationResult result = comparisonResult != -1
            ? new ComparisonEvaluationResult(comparisonResult == 1)
            : comparisonError != null ? new ComparisonEvaluationResult(comparisonError) : null;

        ComparisonExpressionEvaluator evaluator = new ComparisonExpressionEvaluator(result);
        builder.getDocument().getFieldOptions().setComparisonExpressionEvaluator(evaluator);

        builder.getDocument().updateFields();

        Assert.assertEquals(expectedResult, field.getResult());
        evaluator.assertInvocationsCount(1).assertInvocationArguments(0, LEFT, _OPERATOR, RIGHT);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "conditionEvaluationExtensionPointDataProvider")
	public static Object[][] conditionEvaluationExtensionPointDataProvider() throws Exception
	{
		return new Object[][]
		{
			{" IF {0} {1} {2} \"true argument\" \"false argument\" ",  1,  null,  "true argument"},
			{" IF {0} {1} {2} \"true argument\" \"false argument\" ",  0,  null,  "false argument"},
			{" IF {0} {1} {2} \"true argument\" \"false argument\" ",  -1,  "Custom Error",  "Custom Error"},
			{" IF {0} {1} {2} \"true argument\" \"false argument\" ",  -1,  null,  "true argument"},
			{" COMPARE {0} {1} {2} ",  1,  null,  "1"},
			{" COMPARE {0} {1} {2} ",  0,  null,  "0"},
			{" COMPARE {0} {1} {2} ",  -1,  "Custom Error",  "Custom Error"},
			{" COMPARE {0} {1} {2} ",  -1,  null,  "1"},
		};
	}

    /// <summary>
    /// Comparison expressions evaluation for the FieldIf and FieldCompare.
    /// </summary>
    private static class ComparisonExpressionEvaluator implements IComparisonExpressionEvaluator
    {
        public ComparisonExpressionEvaluator(ComparisonEvaluationResult result)
        {
            mResult = result;
        }

        public ComparisonEvaluationResult evaluate(Field field, ComparisonExpression expression)
        {
            mInvocations.add(new String[]
            {
                expression.getLeftExpression(),
                expression.getComparisonOperator(),
                expression.getRightExpression()
            });

            return mResult;
        }

        public ComparisonExpressionEvaluator assertInvocationsCount(int expected)
        {
            Assert.assertEquals(expected, mInvocations.size());
            return this;
        }

        public ComparisonExpressionEvaluator assertInvocationArguments(
            int invocationIndex,
            String expectedLeftExpression,
            String expectedComparisonOperator,
            String expectedRightExpression)
        {
            String[] arguments = mInvocations.get(invocationIndex);

            Assert.assertEquals(expectedLeftExpression, arguments[0]);
            Assert.assertEquals(expectedComparisonOperator, arguments[1]);
            Assert.assertEquals(expectedRightExpression, arguments[2]);

            return this;
        }

        private /*final*/ ComparisonEvaluationResult mResult;
        private /*final*/ ArrayList<String[]> mInvocations = new ArrayList<String[]>();
    } 
    //ExEnd

    @Test
    public void comparisonExpressionEvaluatorNestedFields() throws Exception
    {
        Document document = new Document();

        new FieldBuilder(FieldType.FIELD_IF)
            .addArgument(
                new FieldBuilder(FieldType.FIELD_IF)
                    .addArgument(123)
                    .addArgument(">")
                    .addArgument(666)
                    .addArgument("left greater than right")
                    .addArgument("left less than right"))
            .addArgument("<>")
            .addArgument(new FieldBuilder(FieldType.FIELD_IF)
                .addArgument("left expression")
                .addArgument("=")
                .addArgument("right expression")
                .addArgument("expression are equal")
                .addArgument("expression are not equal"))
            .addArgument(new FieldBuilder(FieldType.FIELD_IF)
                    .addArgument(new FieldArgumentBuilder()
                        .addText("#")
                        .addField(new FieldBuilder(FieldType.FIELD_PAGE)))
                    .addArgument("=")
                    .addArgument(new FieldArgumentBuilder()
                        .addText("#")
                        .addField(new FieldBuilder(FieldType.FIELD_NUM_PAGES)))
                    .addArgument("the last page")
                    .addArgument("not the last page"))
            .addArgument(new FieldBuilder(FieldType.FIELD_IF)
                    .addArgument("unexpected")
                    .addArgument("=")
                    .addArgument("unexpected")
                    .addArgument("unexpected")
                    .addArgument("unexpected"))
            .buildAndInsert(document.getFirstSection().getBody().getFirstParagraph());

        ComparisonExpressionEvaluator evaluator = new ComparisonExpressionEvaluator(null);
        document.getFieldOptions().setComparisonExpressionEvaluator(evaluator);

        document.updateFields();

        evaluator
            .assertInvocationsCount(4)
            .assertInvocationArguments(0, "123", ">", "666")
            .assertInvocationArguments(1, "\"left expression\"", "=", "\"right expression\"")
            .assertInvocationArguments(2, "left less than right", "<>", "expression are not equal")
            .assertInvocationArguments(3, "\"#1\"", "=", "\"#1\"");
    }

    @Test
    public void comparisonExpressionEvaluatorHeaderFooterFields() throws Exception
    {
        Document document = new Document();
        DocumentBuilder builder = new DocumentBuilder(document);

        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        new FieldBuilder(FieldType.FIELD_IF)
            .addArgument(new FieldBuilder(FieldType.FIELD_PAGE))
            .addArgument("=")
            .addArgument(new FieldBuilder(FieldType.FIELD_NUM_PAGES))
            .addArgument(new FieldArgumentBuilder()
                .addField(new FieldBuilder(FieldType.FIELD_PAGE))
                .addText(" / ")
                .addField(new FieldBuilder(FieldType.FIELD_NUM_PAGES)))
            .addArgument(new FieldArgumentBuilder()
                .addField(new FieldBuilder(FieldType.FIELD_PAGE))
                .addText(" / ")
                .addField(new FieldBuilder(FieldType.FIELD_NUM_PAGES)))
            .buildAndInsert(builder.getCurrentParagraph());

        ComparisonExpressionEvaluator evaluator = new ComparisonExpressionEvaluator(null);
        document.getFieldOptions().setComparisonExpressionEvaluator(evaluator);

        document.updateFields();

        evaluator
            .assertInvocationsCount(3)
            .assertInvocationArguments(0, "1", "=", "3")
            .assertInvocationArguments(1, "2", "=", "3")
            .assertInvocationArguments(2, "3", "=", "3");
    }
}
