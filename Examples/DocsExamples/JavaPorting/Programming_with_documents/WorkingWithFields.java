package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldUpdateCultureSource;
import com.aspose.ms.System.DateTime;
import com.aspose.words.Field;
import com.aspose.words.FieldType;
import com.aspose.words.FieldHyperlink;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.FieldStart;
import com.aspose.ms.ms;
import com.aspose.words.Run;
import com.aspose.ms.System.Text.RegularExpressions.Match;
import java.text.MessageFormat;
import com.aspose.words.Node;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.Paragraph;
import com.aspose.words.FieldTA;
import com.aspose.words.FieldToa;
import com.aspose.words.BreakType;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.FieldMergeField;
import com.aspose.words.FieldAddressBlock;
import com.aspose.words.FieldIncludeText;
import com.aspose.words.FieldUnknown;
import com.aspose.words.FieldAuthor;
import com.aspose.words.FieldAsk;
import com.aspose.words.FieldAdvance;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IFieldUpdateCultureProvider;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.Globalization.msDateTimeFormatInfo;
import com.aspose.words.FieldIf;
import com.aspose.words.FieldIfComparisonResult;
import com.aspose.ms.System.Threading.CurrentThread;
import java.util.Date;


class WorkingWithFields extends DocsExamplesBase
{
    @Test
    public void changeFieldUpdateCultureSource() throws Exception
    {
        //ExStart:ChangeFieldUpdateCultureSource
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
        doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { new DateTime(2011, 1, 1) });
        
        doc.save(getArtifactsDir() + "WorkingWithFields.ChangeFieldUpdateCultureSource.docx");
        //ExEnd:ChangeFieldUpdateCultureSource
    }

    @Test
    public void specifyLocaleAtFieldLevel() throws Exception
    {
        //ExStart:SpecifylocaleAtFieldlevel
        DocumentBuilder builder = new DocumentBuilder();

        Field field = builder.insertField(FieldType.FIELD_DATE, true);
        field.setLocaleId(1049);
        
        builder.getDocument().save(getArtifactsDir() + "WorkingWithFields.SpecifylocaleAtFieldlevel.docx");
        //ExEnd:SpecifylocaleAtFieldlevel
    }

    @Test
    public void replaceHyperlinks() throws Exception
    {
        //ExStart:ReplaceHyperlinks
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
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder.insertField("MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        // Select all field start nodes so we can find the merge fields.
        NodeCollection fieldStarts = doc.getChildNodes(NodeType.FIELD_START, true);
        for (FieldStart fieldStart : (Iterable<FieldStart>) fieldStarts)
        {
            if (fieldStart.getFieldType() == FieldType.FIELD_MERGE_FIELD)
            {
                MergeField mergeField = new MergeField(fieldStart);
                mergeField.(mergeField.getName() + "_Renamed");
            }
        }

        doc.save(getArtifactsDir() + "WorkingWithFields.RenameMergeFields.doc");
        //ExEnd:RenameMergeFields
    }

    //ExStart:MergeField
    /// <summary>
    /// Represents a facade object for a merge field in a Microsoft Word document.
    /// </summary>
    static class MergeField
    {
        MergeField(FieldStart fieldStart)
        {
            if (fieldStart == null)
                throw new NullPointerException(ms.nameof("fieldStart"));
            if (fieldStart.getFieldType() != FieldType.FIELD_MERGE_FIELD)
                throw new IllegalArgumentException("Field start type must be FieldMergeField.");

            mFieldStart = fieldStart;

            // Find the field separator node.
            mFieldSeparator = fieldStart.getField().getSeparator();
            if (mFieldSeparator == null)
                throw new IllegalStateException("Cannot find field separator.");

            mFieldEnd = fieldStart.getField().getEnd();
        }

        /// <summary>
        /// Gets or sets the name of the merge field.
        /// </summary>
        String getName() { return mName; }

        private String mName; => ((FieldStart) mFieldStart).GetField().Result.Replace("«", "").Replace("»", "");
            set
            {
                // Merge field name is stored in the field result which is a Run
                // node between field separator and field end.
                Run fieldResult = (Run) mFieldSeparator.NextSibling;
                fieldResult.Text = String.Format("«{0}»", value);

                // But sometimes the field result can consist of more than one run, delete these runs.
                RemoveSameParent(fieldResult.NextSibling, mFieldEnd);

                UpdateFieldCode(value);
            }
        }

        private void updateFieldCode(String fieldName)
        {
            // Field code is stored in a Run node between field start and field separator.
            Run fieldCode = (Run) mFieldStart.getNextSibling();

            Match match = gRegex.match(((FieldStart) mFieldStart).getField().getFieldCode());

            String newFieldCode = MessageFormat.format(" {0}{1} ", match.getGroups().get("start").getValue(), fieldName);
            fieldCode.setText(newFieldCode);

            // But sometimes the field code can consist of more than one run, delete these runs.
            removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
        }

        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Start and end are assumed to have the same parent.
        /// </summary>
        private void removeSameParent(Node startNode, Node endNode)
        {
            if (endNode != null && startNode.getParentNode() != endNode.getParentNode())
                throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

            Node curChild = startNode;
            while (curChild != null && curChild != endNode)
            {
                Node nextChild = curChild.getNextSibling();
                curChild.remove();
                curChild = nextChild;
            }
        }

        private /*final*/ Node mFieldStart;
        private /*final*/ Node mFieldSeparator;
        private /*final*/ Node mFieldEnd;

        private /*final*/ Regex gRegex = new Regex("\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+");
    }
    //ExEnd:MergeField

    @Test
    public void removeField() throws Exception
    {
        //ExStart:RemoveField
        Document doc = new Document(getMyDir() + "Various fields.docx");
        
        Field field = doc.getRange().getFields().get(0);
        field.remove();
        //ExEnd:RemoveField
    }

    @Test
    public void insertTOAFieldWithoutDocumentBuilder() throws Exception
    {
        //ExStart:InsertTOAFieldWithoutDocumentBuilder
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

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertTOAFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertTOAFieldWithoutDocumentBuilder
    }

    @Test
    public void insertNestedFields() throws Exception
    {
        //ExStart:InsertNestedFields
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
    public void insertMergeFieldUsingDOM() throws Exception
    {
        //ExStart:InsertMergeFieldUsingDOM
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(0);

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

        // Finally update this merge field
        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertMergeFieldUsingDOM.docx");
        //ExEnd:InsertMergeFieldUsingDOM
    }

    @Test
    public void insertMailMergeAddressBlockFieldUsingDOM() throws Exception
    {
        //ExStart:InsertMailMergeAddressBlockFieldUsingDOM
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(0);

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

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertMailMergeAddressBlockFieldUsingDOM.docx");
        //ExEnd:InsertMailMergeAddressBlockFieldUsingDOM
    }

    @Test
    public void insertFieldIncludeTextWithoutDocumentBuilder() throws Exception
    {
        //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
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
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertField("MERGEFIELD MyFieldName \\* MERGEFORMAT");
        
        doc.save(getArtifactsDir() + "WorkingWithFields.InsertField.docx");
        //ExEnd:InsertField
    }

    @Test
    public void insertAuthorField() throws Exception
    {
        //ExStart:InsertAuthorField
        Document doc = new Document();

        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(0);

        // We want to insert an AUTHOR field like this:
        // { AUTHOR Test1 }

        FieldAuthor field = (FieldAuthor) para.appendField(FieldType.FIELD_AUTHOR, false);            
        field.setAuthorName("Test1"); // { AUTHOR Test1 }

        field.update();

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertAuthorField.docx");
        //ExEnd:InsertAuthorField
    }

    @Test
    public void insertASKFieldWithOutDocumentBuilder() throws Exception
    {
        //ExStart:InsertASKFieldWithOutDocumentBuilder
        Document doc = new Document();

        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(0);

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

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertASKFieldWithOutDocumentBuilder.docx");
        //ExEnd:InsertASKFieldWithOutDocumentBuilder
    }

    @Test
    public void insertAdvanceFieldWithOutDocumentBuilder() throws Exception
    {
        //ExStart:InsertAdvanceFieldWithOutDocumentBuilder
        Document doc = new Document();

        Paragraph para = (Paragraph) doc.getChildNodes(NodeType.PARAGRAPH, true).get(0);

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

        doc.save(getArtifactsDir() + "WorkingWithFields.InsertAdvanceFieldWithOutDocumentBuilder.docx");
        //ExEnd:InsertAdvanceFieldWithOutDocumentBuilder
    }

    @Test
    public void getMailMergeFieldNames() throws Exception
    {
        //ExStart:GetFieldNames
        Document doc = new Document();

        String[] fieldNames = doc.getMailMerge().getFieldNames();
        //ExEnd:GetFieldNames
        System.out.println("\nDocument have " + fieldNames.length + " fields.");
    }

    @Test
    public void mappedDataFields() throws Exception
    {
        //ExStart:MappedDataFields
        Document doc = new Document();

        doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
        //ExEnd:MappedDataFields
    }

    @Test
    public void deleteFields() throws Exception
    {
        //ExStart:DeleteFields
        Document doc = new Document();

        doc.getMailMerge().deleteFields();
        //ExEnd:DeleteFields
    }

    @Test
    public void fieldUpdateCulture() throws Exception
    {
        //ExStart:FieldUpdateCultureProvider
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(FieldType.FIELD_TIME, true);

        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getFieldOptions().setFieldUpdateCultureProvider(new FieldUpdateCultureProvider());

        doc.save(getArtifactsDir() + "WorkingWithFields.FieldUpdateCulture.pdf");
        //ExEnd:FieldUpdateCultureProvider
    }

    //ExStart:FieldUpdateCultureProviderGetCulture
    private static class FieldUpdateCultureProvider implements IFieldUpdateCultureProvider
    {
        public msCultureInfo getCulture(String name, Field field)
        {
            switch (gStringSwitchMap.of(name))
            {
                case /*"ru-RU"*/0:
                    msCultureInfo culture = new msCultureInfo(name, false);
                    msDateTimeFormatInfo format = culture.getDateTimeFormat();

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
                case /*"en-US"*/1:
                    return new msCultureInfo(name, false);
                default:
                    return null;
            }
        }
    }
    //ExEnd:FieldUpdateCultureProviderGetCulture

    @Test
    public void fieldDisplayResults() throws Exception
    {
        //ExStart:FieldDisplayResults
        //ExStart:UpdateDocFields
        Document document = new Document(getMyDir() + "Various fields.docx");

        document.updateFields();
        //ExEnd:UpdateDocFields

        for (Field field : document.getRange().getFields())
            System.out.println(field.getDisplayResult());
        //ExEnd:FieldDisplayResults
    }

    @Test
    public void evaluateIFCondition() throws Exception
    {
        //ExStart:EvaluateIFCondition
        DocumentBuilder builder = new DocumentBuilder();

        FieldIf field = (FieldIf) builder.insertField("IF 1 = 1", null);
        /*FieldIfComparisonResult*/int actualResult = field.evaluateCondition();

        System.out.println(actualResult);
        //ExEnd:EvaluateIFCondition
    }

    @Test
    public void convertFieldsInParagraph() throws Exception
    {
        //ExStart:ConvertFieldsInParagraph
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last 
        // paragraph of the document.
        doc.getFirstSection().getBody().getLastParagraph().getRange().getFields().Where(f => f.Type == FieldType.FieldIf).ToList()
            .ForEach(f => f.Unlink());

        doc.save(getArtifactsDir() + "WorkingWithFields.TestFile.docx");
        //ExEnd:ConvertFieldsInParagraph
    }

    @Test
    public void convertFieldsInDocument() throws Exception
    {
        //ExStart:ConvertFieldsInDocument
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        doc.getRange().getFields().Where(f => f.Type == FieldType.FieldIf).ToList().ForEach(f => f.Unlink());

        // Save the document with fields transformed to disk
        doc.save(getArtifactsDir() + "WorkingWithFields.ConvertFieldsInDocument.docx");
        //ExEnd:ConvertFieldsInDocument
    }

    @Test
    public void convertFieldsInBody() throws Exception
    {
        //ExStart:ConvertFieldsInBody
        Document doc = new Document(getMyDir() + "Linked fields.docx");

        // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        doc.getFirstSection().getBody().getRange().getFields().Where(f => f.Type == FieldType.FieldPage).ToList().ForEach(f => f.Unlink());

        doc.save(getArtifactsDir() + "WorkingWithFields.ConvertFieldsInBody.docx");
        //ExEnd:ConvertFieldsInBody
    }

    @Test
    public void changeLocale() throws Exception
    {
        //ExStart:ChangeLocale
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete.
        msCultureInfo currentCulture = CurrentThread.getCurrentCulture();
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        CurrentThread.setCurrentCulture(new msCultureInfo("de-DE"));

        doc.getMailMerge().execute(new String[] { "Date" }, new Object[] { new Date() });
        
        CurrentThread.setCurrentCulture(currentCulture);
        
        doc.save(getArtifactsDir() + "WorkingWithFields.ChangeLocale.docx");
        //ExEnd:ChangeLocale
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"ru-RU",
		"en-US"
	);

}
