package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.email.system.DateTime;
import com.aspose.words.*;
import com.aspose.words.net.System.Globalization.CultureInfo;
import com.aspose.words.net.System.Globalization.DateTimeFormatInfo;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.time.LocalDate;
import java.time.Month;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Test
public class WorkingWithFields extends DocsExamplesBase
{
    @Test
    public void fieldCode() throws Exception
    {
        //ExStart:FieldCode
        //GistId:7c2b7b650a88375b1d438746f78f0d64
        Document doc = new Document(getMyDir() + "Hyperlinks.docx");

        for (Field field : doc.getRange().getFields())
        {
            String fieldCode = field.getFieldCode();
            String fieldResult = field.getResult();
        }
        //ExEnd:FieldCode
    }

    @Test
    public void changeFieldUpdateCultureSource() throws Exception
    {
        //ExStart:ChangeFieldUpdateCultureSource
        //GistId:9e90defe4a7bcafb004f73a2ef236986
        //ExStart:DocumentBuilderInsertField
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert content with German locale.
        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
        //ExEnd:DocumentBuilderInsertField

        // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
        // set the culture used during field update to the culture used by the field.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { LocalDate.of(2011, Month.JANUARY, 1) });
        
        doc.save(getArtifactsDir() + "WorkingWithFields.ChangeFieldUpdateCultureSource.docx");
        //ExEnd:ChangeFieldUpdateCultureSource
    }

    @Test
    public void specifyLocaleAtFieldLevel() throws Exception
    {
        //ExStart:SpecifyLocaleAtFieldLevel
        //GistId:1cf07762df56f15067d6aef90b14b3db
        DocumentBuilder builder = new DocumentBuilder();

        Field field = builder.insertField(FieldType.FIELD_DATE, true);
        field.setLocaleId(1049);
        
        builder.getDocument().save(getArtifactsDir() + "WorkingWithFields.SpecifyLocaleAtFieldLevel.docx");
        //ExEnd:SpecifyLocaleAtFieldLevel
    }

    @Test
    public void replaceHyperlinks() throws Exception
    {
        //ExStart:ReplaceHyperlinks
        //GistId:0213851d47551e83af42233f4d075cf6
        Document doc = new Document(getMyDir() + "Hyperlinks.docx");

        for (Field field : doc.getRange().getFields())
        {
            if (field.getType() == FieldType.FIELD_HYPERLINK)
            {
                FieldHyperlink hyperlink = (FieldHyperlink) field;

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if (hyperlink.getSubAddress() != null)
                    continue;

                hyperlink.setAddress("http://www.aspose.com");
                hyperlink.setResult("Aspose - The .NET & Java Component Publisher");
            }
        }

        doc.save(getArtifactsDir() + "WorkingWithFields.ReplaceHyperlinks.docx");
        //ExEnd:ReplaceHyperlinks
    }

    @Test
    public void renameMergeFields() throws Exception
    {
        //ExStart:RenameMergeFields
        //GistId:bf0f8a6b40b69a5274ab3553315e147f
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder.insertField("MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        for (Field f : doc.getRange().getFields())
        {
            if (f.getType() == FieldType.FIELD_MERGE_FIELD)
            {
                FieldMergeField mergeField = (FieldMergeField)f;
                mergeField.setFieldName(mergeField.getFieldName() + "_Renamed");
                mergeField.update();
            }
        }

        doc.save(getArtifactsDir() + "WorkingWithFields.RenameMergeFields.docx");
        //ExEnd:RenameMergeFields
    }

    @Test
    public void removeField() throws Exception
    {
        //ExStart:RemoveField
        //GistId:8c604665c1b97795df7a1e665f6b44ce
        Document doc = new Document(getMyDir() + "Various fields.docx");
        
        Field field = doc.getRange().getFields().get(0);
        field.remove();
        //ExEnd:RemoveField
    }

    @Test
    public void unlinkFields() throws Exception
    {
        //ExStart:UnlinkFields
        //GistId:f3592014d179ecb43905e37b2a68bc92
        Document doc = new Document(getMyDir() + "Various fields.docx");
        doc.unlinkFields();
        //ExEnd:UnlinkFields
    }

    @Test
    public void insertToaFieldWithoutDocumentBuilder() throws Exception
    {
        //ExStart:InsertToaFieldWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();
        Paragraph para = new Paragraph(doc);

        // We want to insert TA and TOA fields like this:
        // { TA  \c 1 \l "Value 0" }
        // { TOA  \c 1 }

        FieldTA fieldTA = (FieldTA) para.appendField(FieldType.FIELD_TOA_ENTRY, false);
        fieldTA.setEntryCategory("1");
        fieldTA.setLongCitation("Value 0");

        doc.getFirstSection().getBody().appendChild(para);

        para = new Paragraph(doc);

        FieldToa fieldToa = (FieldToa) para.appendField(FieldType.FIELD_TOA, false);
        fieldToa.setEntryCategory("1");
        doc.getFirstSection().getBody().appendChild(para);

        fieldToa.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertToaFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertToaFieldWithoutDocumentBuilder
    }

    @Test
    public void insertNestedFields() throws Exception
    {
        //ExStart:InsertNestedFields
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 0; i < 5; i++)
            builder.insertBreak(BreakType.PAGE_BREAK);

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);

        // We want to insert a field like this:
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        Field field = builder.insertField("IF ");
        builder.moveTo(field.getSeparator());
        builder.insertField("PAGE");
        builder.write(" <> ");
        builder.insertField("NUMPAGES");
        builder.write(" \"See Next Page\" \"Last Page\" ");

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertNestedFields.docx");
        //ExEnd:InsertNestedFields
    }

    @Test
    public void insertMergeFieldUsingDom() throws Exception
    {
        //ExStart:InsertMergeFieldUsingDom
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);
        builder.moveTo(para);

        // We want to insert a merge field like this:
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        FieldMergeField field = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, false);
        // { " MERGEFIELD Test1" }
        field.setFieldName("Test1");
        // { " MERGEFIELD Test1 \\b Test2" }
        field.setTextBefore("Test2");
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field.setTextAfter("Test3");
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field.isMapped(true);
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field.isVerticalFormatting(true);

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertMergeFieldUsingDom.docx");
        //ExEnd:InsertMergeFieldUsingDom
    }

    @Test
    public void insertAddressBlockFieldUsingDom() throws Exception
    {
        //ExStart:InsertAddressBlockFieldUsingDom
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);
        builder.moveTo(para);

        // We want to insert a mail merge address block like this:
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        FieldAddressBlock field = (FieldAddressBlock) builder.insertField(FieldType.FIELD_ADDRESS_BLOCK, false);
        // { ADDRESSBLOCK \\c 1" }
        field.setIncludeCountryOrRegionName("1");
        // { ADDRESSBLOCK \\c 1 \\d" }
        field.setFormatAddressOnCountryOrRegion(true);
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field.setExcludedCountryOrRegionName("Test2");
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field.setNameAndAddressFormat("Test3");
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field.setLanguageId("Test 4");

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertAddressBlockFieldUsingDom.docx");
        //ExEnd:InsertAddressBlockFieldUsingDom
    }

    @Test
    public void insertFieldIncludeTextWithoutDocumentBuilder() throws Exception
    {
        //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();

        Paragraph para = new Paragraph(doc);

        // We want to insert an INCLUDETEXT field like this:
        // { INCLUDETEXT  "file path" }
        FieldIncludeText fieldIncludeText = (FieldIncludeText) para.appendField(FieldType.FIELD_INCLUDE_TEXT, false);
        fieldIncludeText.setBookmarkName("bookmark");
        fieldIncludeText.setSourceFullName(getMyDir() + "IncludeText.docx");

        doc.getFirstSection().getBody().appendChild(para);

        fieldIncludeText.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertIncludeFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
    }

    @Test
    public void insertFieldNone() throws Exception
    {
        //ExStart:InsertFieldNone
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldUnknown field = (FieldUnknown) builder.insertField(FieldType.FIELD_NONE, false);

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertFieldNone.docx");
        //ExEnd:InsertFieldNone
    }

    @Test
    public void insertField() throws Exception
    {
        //ExStart:InsertField
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertField("MERGEFIELD MyFieldName \\* MERGEFORMAT");
        
        doc.save(getArtifactsDir() + "WorkingWithFields.InsertField.docx");
        //ExEnd:InsertField
    }

    @Test
    public void insertFieldUsingFieldBuilder() throws Exception
    {
        //ExStart:InsertFieldUsingFieldBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();

        // Prepare IF field with two nested MERGEFIELD fields: { IF "left expression" = "right expression" "Firstname: { MERGEFIELD firstname }" "Lastname: { MERGEFIELD lastname }"}
        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_IF)
            .addArgument("left expression")
            .addArgument("=")
            .addArgument("right expression")
            .addArgument(
                new FieldArgumentBuilder()
                    .addText("Firstname: ")
                    .addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD).addArgument("firstname")))
            .addArgument(
                new FieldArgumentBuilder()
                    .addText("Lastname: ")
                    .addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD).addArgument("lastname")));

        // Insert IF field in exact location
        Field field = fieldBuilder.buildAndInsert(doc.getFirstSection().getBody().getFirstParagraph());
        field.update();

        doc.save(getArtifactsDir() + "Field.InsertFieldUsingFieldBuilder.docx");
        //ExEnd:InsertFieldUsingFieldBuilder
    }

    @Test
    public void insertAuthorField() throws Exception
    {
        //ExStart:InsertAuthorField
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();

        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // We want to insert an AUTHOR field like this:
        // { AUTHOR Test1 }
        FieldAuthor field = (FieldAuthor) para.appendField(FieldType.FIELD_AUTHOR, false);            
        field.setAuthorName("Test1");

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertAuthorField.docx");
        //ExEnd:InsertAuthorField
    }

    @Test
    public void insertAskFieldWithoutDocumentBuilder() throws Exception
    {
        //ExStart:InsertAskFieldWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();

        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);
        // We want to insert an Ask field like this:
        // { ASK \"Test 1\" Test2 \\d Test3 \\o }
        FieldAsk field = (FieldAsk) para.appendField(FieldType.FIELD_ASK, false);
        // { ASK \"Test 1\" " }
        field.setBookmarkName("Test 1");
        // { ASK \"Test 1\" Test2 }
        field.setPromptText("Test2");
        // { ASK \"Test 1\" Test2 \\d Test3 }
        field.setDefaultResponse("Test3");
        // { ASK \"Test 1\" Test2 \\d Test3 \\o }
        field.setPromptOnceOnMailMerge(true);

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertAskFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertAskFieldWithoutDocumentBuilder
    }

    @Test
    public void insertAdvanceFieldWithoutDocumentBuilder() throws Exception
    {
        //ExStart:InsertAdvanceFieldWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        Document doc = new Document();

        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);
        // We want to insert an Advance field like this:
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        FieldAdvance field = (FieldAdvance) para.appendField(FieldType.FIELD_ADVANCE, false);
        // { ADVANCE \\d 10 " }
        field.setDownOffset("10");
        // { ADVANCE \\d 10 \\l 10 }
        field.setLeftOffset("10");
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
        field.setRightOffset("-3.3");
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
        field.setUpOffset("0");
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
        field.setHorizontalPosition("100");
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        field.setVerticalPosition("100");

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertAdvanceFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertAdvanceFieldWithoutDocumentBuilder
    }

    @Test
    public void getMailMergeFieldNames() throws Exception
    {
        //ExStart:GetFieldNames
        //GistId:b4bab1bf22437a86d8062e91cf154494
        Document doc = new Document();

        String[] fieldNames = doc.getMailMerge().getFieldNames();
        //ExEnd:GetFieldNames
        System.out.println("\nDocument have " + fieldNames.length + " fields.");
    }

    @Test
    public void mappedDataFields() throws Exception
    {
        //ExStart:MappedDataFields
        //GistId:b4bab1bf22437a86d8062e91cf154494
        Document doc = new Document();

        doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
        //ExEnd:MappedDataFields
    }

    @Test
    public void deleteFields() throws Exception
    {
        //ExStart:DeleteFields
        //GistId:f39874821cb317d245a769c9ce346fea
        Document doc = new Document();

        doc.getMailMerge().deleteFields();
        //ExEnd:DeleteFields
    }

    @Test
    public void fieldUpdateCulture() throws Exception
    {
        //ExStart:FieldUpdateCulture
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(FieldType.FIELD_TIME, true);

        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getFieldOptions().setFieldUpdateCultureProvider(new FieldUpdateCultureProvider());

        doc.save(getArtifactsDir() + "WorkingWithFields.FieldUpdateCulture.pdf");
        //ExEnd:FieldUpdateCulture
    }

    //ExStart:FieldUpdateCultureProvider
    //GistId:79b46682fbfd7f02f64783b163ed95fc
    private static class FieldUpdateCultureProvider implements IFieldUpdateCultureProvider
    {
        public CultureInfo getCulture(String name, Field field)
        {
            switch (name)
            {
                case "ru-RU":
                    CultureInfo culture = new CultureInfo(new Locale(name));
                    DateTimeFormatInfo format = culture.getDateTimeFormat();

                    format.setMonthNames(new String[]
                    {
                        "месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8",
                        "месяц 9", "месяц 10", "месяц 11", "месяц 12", ""
                    });
                    format.setMonthGenitiveNames(format.getMonthNames());
                    format.setAbbreviatedMonthNames(new String[]
                    {
                        "мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10",
                        "мес 11", "мес 12", ""
                    });
                    format.setAbbreviatedMonthGenitiveNames(format.getAbbreviatedMonthNames());

                    format.setDayNames(new String[]
                    {
                        "день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4",
                        "день недели 5", "день недели 6"
                    });
                    format.setAbbreviatedDayNames(new String[]
                        { "день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6" });
                    format.setShortestDayNames(new String[] { "д7", "д1", "д2", "д3", "д4", "д5", "д6" });

                    format.setAMDesignator("До полудня");
                    format.setPMDesignator("После полудня");

                    final String PATTERN = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                    format.setLongDatePattern(PATTERN);
                    format.setLongTimePattern(PATTERN);
                    format.setShortDatePattern(PATTERN);
                    format.setShortTimePattern(PATTERN);

                    return culture;
                case "en-US":
                    return new CultureInfo(new Locale(name));
                default:
                    return null;
            }
        }
    }
    //ExEnd:FieldUpdateCultureProvider

    @Test
    public void fieldDisplayResults() throws Exception
    {
        //ExStart:FieldDisplayResults
        //GistId:bf0f8a6b40b69a5274ab3553315e147f
        //ExStart:UpdateDocFields
        //GistId:08db64c4d86842c4afd1ecb925ed07c4
        Document document = new Document(getMyDir() + "Various fields.docx");

        document.updateFields();
        //ExEnd:UpdateDocFields

        for (Field field : document.getRange().getFields())
            System.out.println(field.getDisplayResult());
        //ExEnd:FieldDisplayResults
    }

    @Test
    public void evaluateIfCondition() throws Exception
    {
        //ExStart:EvaluateIfCondition
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        DocumentBuilder builder = new DocumentBuilder();

        FieldIf field = (FieldIf) builder.insertField("IF 1 = 1", null);
        int actualResult = field.evaluateCondition();

        System.out.println(actualResult);
        //ExEnd:EvaluateIfCondition
    }

    @Test
    public void unlinkFieldsInParagraph() throws Exception
    {
        //ExStart:UnlinkFieldsInParagraph
        //GistId:f3592014d179ecb43905e37b2a68bc92
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last 
        // paragraph of the document.
        for (Field field : doc.getFirstSection().getBody().getLastParagraph().getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_IF) {
                field.unlink();
            }
        }

        doc.save(getArtifactsDir() + "WorkingWithFields.UnlinkFieldsInParagraph.docx");
        //ExEnd:UnlinkFieldsInParagraph
    }

    @Test
    public void unlinkFieldsInDocument() throws Exception
    {
        //ExStart:UnlinkFieldsInDocument
        //GistId:f3592014d179ecb43905e37b2a68bc92
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_IF) {
                field.unlink();
            }
        }

        // Save the document with fields transformed to disk
        doc.save(getArtifactsDir() + "WorkingWithFields.UnlinkFieldsInDocument.docx");
        //ExEnd:UnlinkFieldsInDocument
    }

    @Test
    public void unlinkFieldsInBody() throws Exception
    {
        //ExStart:UnlinkFieldsInBody
        //GistId:f3592014d179ecb43905e37b2a68bc92
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_PAGE) {
                field.unlink();
            }
        }

        doc.save(getArtifactsDir() + "WorkingWithFields.UnlinkFieldsInBody.docx");
        //ExEnd:UnlinkFieldsInBody
    }

    @Test
    public void changeLocale() throws Exception
    {
        //ExStart:ChangeLocale
        //GistId:9e90defe4a7bcafb004f73a2ef236986
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete.
        CultureInfo currentCulture = new CultureInfo("");
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        Locale.setDefault(new Locale("de-DE"));

        doc.getMailMerge().execute(new String[] { "Date" }, new Object[] { new Date() });

        Locale.setDefault(currentCulture.getLocale(currentCulture));
        
        doc.save(getArtifactsDir() + "WorkingWithFields.ChangeLocale.docx");
        //ExEnd:ChangeLocale
    }

    //ExStart:ConvertFieldsToStaticText
    //GistId:f3592014d179ecb43905e37b2a68bc92
    /// <summary>
    /// Converts any fields of the specified type found in the descendants of the node into static text.
    /// </summary>
    /// <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
    /// <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
    public void convertFieldsToStaticText(CompositeNode compositeNode, int targetFieldType) throws Exception {
        // Get all fields in the composite node's range.
        FieldCollection fields = compositeNode.getRange().getFields();

        // Iterate through the fields and unlink those that match the target field type.
        for (Field field : fields) {
            if (field.getType() == targetFieldType) {
                field.unlink();
            }
        }
    }
    //ExEnd:ConvertFieldsToStaticText

    @Test
    public void fieldResultFormatting() throws Exception
    {
        //ExStart:FieldResultFormatting
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        FieldResultFormatter formatter = new FieldResultFormatter("${0}", "Date: {0}", "Item # {0}:");
        doc.getFieldOptions().setResultFormatter(formatter);

        // Our field result formatter applies a custom format to newly created fields of three types of formats.
        // Field result formatters apply new formatting to fields as they are updated,
        // which happens as soon as we create them using this InsertField method overload.
        // 1 -  Numeric:
        builder.insertField(" = 2 + 3 \\# $###");

        Assert.assertEquals("$5", doc.getRange().getFields().get(0).getResult());
        Assert.assertEquals(1, formatter.countFormatInvocations(FieldResultFormatter.FormatInvocationType.NUMERIC));

        // 2 -  Date/time:
        builder.insertField("DATE \\@ \"d MMMM yyyy\"");

        Assert.assertTrue(doc.getRange().getFields().get(1).getResult().startsWith("Date: "));
        Assert.assertEquals(1, formatter.countFormatInvocations(FieldResultFormatter.FormatInvocationType.DATE_TIME));

        // 3 -  General:
        builder.insertField("QUOTE \"2\" \\* Ordinal");

        Assert.assertEquals("Item # 2:", doc.getRange().getFields().get(2).getResult());
        Assert.assertEquals(1, formatter.countFormatInvocations(FieldResultFormatter.FormatInvocationType.GENERAL));

        formatter.printFormatInvocations();
        //ExEnd:FieldResultFormatting
    }

    //ExStart:FieldResultFormatter
    //GistId:79b46682fbfd7f02f64783b163ed95fc
    /// <summary>
    /// When fields with formatting are updated, this formatter will override their formatting
    /// with a custom format, while tracking every invocation.
    /// </summary>
    public static class FieldResultFormatter implements IFieldResultFormatter {

        private final String mNumberFormat;
        private final String mDateFormat;
        private final String mGeneralFormat;
        private final ArrayList<FormatInvocation> formatInvocations = new ArrayList<>();

        public FieldResultFormatter(String numberFormat, String dateFormat, String generalFormat) {
            this.mNumberFormat = numberFormat;
            this.mDateFormat = dateFormat;
            this.mGeneralFormat = generalFormat;
        }

        @Override
        public String formatNumeric(double value, String format) {
            if (mNumberFormat == null || mNumberFormat.isEmpty()) {
                return null;
            }

            String newValue = String.format(mNumberFormat, value);
            formatInvocations.add(new FormatInvocation(FormatInvocationType.NUMERIC, value, format, newValue));
            return newValue;
        }

        @Override
        public String formatDateTime(Date value, String format, int calendarType) {
            if (mDateFormat == null || mDateFormat.isEmpty()) {
                return null;
            }

            String newValue = String.format(mDateFormat, value);
            formatInvocations.add(new FormatInvocation(FormatInvocationType.DATE_TIME, value + " (" + calendarType + ")", format, newValue));
            return newValue;
        }

        @Override
        public String format(String value, int format) {
            return format((Object) value, format);
        }

        @Override
        public String format(double value, int format) {
            return format((Object) value, format);
        }

        private String format(Object value, int format) {
            if (mGeneralFormat == null || mGeneralFormat.isEmpty()) {
                return null;
            }

            String newValue = String.format(mGeneralFormat, value);
            formatInvocations.add(new FormatInvocation(FormatInvocationType.GENERAL, value, String.valueOf(format), newValue));
            return newValue;
        }

        public int countFormatInvocations(FormatInvocationType formatInvocationType) {
            if (formatInvocationType == FormatInvocationType.ALL) {
                return formatInvocations.size();
            }
            return (int) formatInvocations.stream()
                    .filter(f -> f.getFormatInvocationType() == formatInvocationType)
                    .count();
        }

        public void printFormatInvocations() {
            for (FormatInvocation f : formatInvocations) {
                System.out.println("Invocation type:\t" + f.getFormatInvocationType() + "\n" +
                        "\tOriginal value:\t\t" + f.getValue() + "\n" +
                        "\tOriginal format:\t" + f.getOriginalFormat() + "\n" +
                        "\tNew value:\t\t\t" + f.getNewValue() + "\n");
            }
        }

        public static class FormatInvocation {
            private final FormatInvocationType formatInvocationType;
            private final Object value;
            private final String originalFormat;
            private final String newValue;

            public FormatInvocation(FormatInvocationType formatInvocationType, Object value, String originalFormat, String newValue) {
                this.formatInvocationType = formatInvocationType;
                this.value = value;
                this.originalFormat = originalFormat;
                this.newValue = newValue;
            }

            public FormatInvocationType getFormatInvocationType() {
                return formatInvocationType;
            }

            public Object getValue() {
                return value;
            }

            public String getOriginalFormat() {
                return originalFormat;
            }

            public String getNewValue() {
                return newValue;
            }
        }

        public enum FormatInvocationType {
            NUMERIC, DATE_TIME, GENERAL, ALL
        }
    }
    //ExEnd:FieldResultFormatter
}
