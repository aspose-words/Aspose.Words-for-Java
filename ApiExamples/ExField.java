//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////


package Examples;

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Date;
import java.util.Locale;
import java.util.ArrayList;
import java.util.regex.Pattern;


public class ExField extends ApiExampleBase
{
    @Test
    public void updateTOC() throws Exception
    {
        Document doc = new Document();

        //ExStart
        //ExId:UpdateTOC
        //ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
        doc.updateFields();
        //ExEnd
    }

    @Test
    public void getFieldType() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        //ExStart
        //ExFor:FieldType
        //ExFor:FieldChar
        //ExFor:FieldChar.FieldType
        //ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
        FieldChar fieldStart = (FieldChar) doc.getChild(NodeType.FIELD_START, 0, true);
        int type = fieldStart.getFieldType();
        //ExEnd
    }

    @Test
    public void getFieldFromDocument() throws Exception
    {
        //ExStart
        //ExFor:FieldChar.GetField
        //ExId:GetField
        //ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        FieldStart fieldStart = (FieldStart) doc.getChild(NodeType.FIELD_START, 0, true);

        // Retrieve the facade object which represents the field in the document.
        Field field = fieldStart.getField();

        System.out.println("Field code:" + field.getFieldCode());
        System.out.println("Field result: " + field.getResult());
        System.out.println("Is locked: " + field.isLocked());

        // This updates only this field in the document.
        field.update();
        //ExEnd
    }

    @Test
    public void createRevNumFieldWithFieldBuilder() throws Exception
    {
        //ExStart
        //ExFor:FieldBuilder.#ctor(FieldType)
        //ExFor:FieldBuilder.BuildAndInsert(Inline)
        //ExSummary:Builds and inserts a field into the document before the specified inline node
        Document doc = new Document();
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_REVISION_NUM);
        fieldBuilder.buildAndInsert(run);

        doc.updateFields();
        //ExEnd
        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        FieldRevNum revNum = (FieldRevNum) doc.getRange().getFields().get(0);
        Assert.assertNotNull(revNum);
    }

    @Test
    public void createRevNumFieldByDocumentBuilder() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("REVNUM MERGEFORMAT");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        FieldRevNum revNum = (FieldRevNum) doc.getRange().getFields().get(0);
        Assert.assertNotNull(revNum);
    }

    @Test
    public void createInfoFieldWithFieldBuilder() throws Exception
    {
        Document doc = new Document();
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_INFO);
        fieldBuilder.buildAndInsert(run);

        doc.updateFields();

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        FieldInfo info = (FieldInfo) doc.getRange().getFields().get(0);
        Assert.assertNotNull(info);
    }

    @Test
    public void createInfoFieldWithDocumentBuilder() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("INFO MERGEFORMAT");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        FieldInfo info = (FieldInfo) doc.getRange().getFields().get(0);
        Assert.assertNotNull(info);
    }

    @Test
    public void getFieldFromFieldCollection() throws Exception
    {
        //ExStart
        //ExId:GetFieldFromFieldCollection
        //ExSummary:Demonstrates how to retrieve a field using the range of a node.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        Field field = doc.getRange().getFields().get(0);

        // This should be the first field in the document - a TOC field.
        System.out.println(field.getType());
        //ExEnd
    }

    @Test
    public void insertTCField() throws Exception
    {
        //ExStart
        //ExId:InsertTCField
        //ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
        // Create a blank document.
        Document doc = new Document();

        // Create a document builder to insert content with.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TC field at the current document builder position.
        builder.insertField("TC \"Entry Text\" \\f t");
        //ExEnd
    }

    @Test
    public void changeLocale() throws Exception
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Date");

        //ExStart
        //ExId:ChangeCurrentCulture
        //ExSummary:Shows how to change the culture used in formatting fields during update.
        // Store the current culture so it can be set back once mail merge is complete.
        Locale currentCulture = Locale.getDefault();
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        Locale.setDefault(new Locale("de", "DE"));

        // Execute mail merge
        doc.getMailMerge().execute(new String[]{"Date"}, new Object[]{new Date()});

        // Restore the original culture.
        Locale.setDefault(currentCulture);
        //ExEnd

        doc.save(getMyDir() + "\\Artifacts\\Field.ChangeLocale.doc");
    }

    @Test
    public void removeTOCFromDocumentCaller() throws Exception
    {
        removeTOCFromDocument();
    }

    //ExStart
    //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
    //ExId:RemoveTableOfContents
    //ExSummary:Demonstrates how to remove a specified TOC from a document.
    public void removeTOCFromDocument() throws Exception
    {
        // Open a document which contains a TOC.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        // Remove the first TOC from the document.
        Field tocField = doc.getRange().getFields().get(0);
        tocField.remove();

        // Save the output.
        doc.save(getMyDir() + "\\Artifacts\\Document.TableOfContentsRemoveTOC.doc");
    }
    //ExEnd

    @Test
    //ExStart
    //ExId:TCFieldsRangeReplace
    //ExSummary:Shows how to find and insert a TC field at text in a document.
    public void insertTCFieldsAtText() throws Exception
    {
        Document doc = new Document();

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new InsertTCFieldHandler("Chapter 1", "\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc.getRange().replace(Pattern.compile("The Beginning"), "", options);
    }

    public class InsertTCFieldHandler implements IReplacingCallback
    {
        // Store the text and switches to be used for the TC fields.
        private String mFieldText;
        private String mFieldSwitches;

        /**
         * The switches to use for each TC field. Can be an empty string or null.
         */
        public InsertTCFieldHandler(String switches) throws Exception
        {
            this(null, switches);
        }

        /**
         * The display text and switches to use for each TC field. Display name can be an empty String or null.
         */
        public InsertTCFieldHandler(String text, String switches)
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        public int replacing(ReplacingArgs args) throws Exception
        {
            // Create a builder to insert the field.
            DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
            // Move to the first node of the match.
            builder.moveTo(args.getMatchNode());

            // If the user specified text to be used in the field as display text then use that, otherwise use the
            // match string as the display text.
            String insertText;

            if (!(mFieldText == null || "".equals(mFieldText))) insertText = mFieldText;
            else insertText = args.getMatch().group();

            // Insert the TC field before this node using the specified string as the display text and user defined switches.
            builder.insertField(java.text.MessageFormat.format("TC \"{0}\" {1}", insertText, mFieldSwitches));

            // We have done what we want so skip replacement.
            return ReplaceAction.SKIP;
        }
    }
    //ExEnd

    @Test(enabled = false, description = "WORDSNET-16037")
    public void insertAndUpdateDirtyField() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field fieldToc = builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        fieldToc.isDirty(true);

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);
        //Assert that field model is correct
        Assert.assertTrue(doc.getRange().getFields().get(0).isDirty());

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setUpdateDirtyFields(false);

        ByteArrayInputStream dstInputStream = new ByteArrayInputStream(dstStream.toByteArray());
        doc = new Document(dstInputStream);
        Field tocField = doc.getRange().getFields().get(0);
        //Assert that isDirty saves
        Assert.assertTrue(tocField.isDirty());
    }

    @Test
    public void insertFieldWithFieldBuilder() throws Exception
    {
        //ExStart
        //ExFor:FieldArgumentBuilder
        //ExFor:FieldArgumentBuilder.AddField(FieldBuilder)
        //ExFor:FieldArgumentBuilder.AddText(String)
        //ExFor:FieldBuilder.AddArgument(FieldArgumentBuilder)
        //ExFor:FieldBuilder.AddArgument(String)
        //ExFor:FieldBuilder.AddArgument(Int32)
        //ExFor:FieldBuilder.AddArgument(Double)
        //ExFor:FieldBuilder.AddSwitch(String, String)
        //ExSummary:Inserts a field into a document using field builder constructor
        Document doc = new Document();

        //Add text into the paragraph
        Paragraph para = doc.getFirstSection().getBody().getParagraphs().get(0);
        Run run = new Run(doc);
        {
            run.setText(" Hello World!");
        }
        para.appendChild(run);

        FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
        argumentBuilder.addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD));
        argumentBuilder.addText("BestField");

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_IF);
        fieldBuilder.addArgument(argumentBuilder).addArgument("=").addArgument("BestField").addArgument(10).addArgument(20.0).addSwitch("12", "13").buildAndInsert(run);

        doc.updateFields();
        //ExEnd
    }

    @Test
    public void insertFieldWithFieldBuilderException() throws Exception
    {
        Document doc = new Document();

        //Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
        argumentBuilder.addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD));
        argumentBuilder.addNode(run);
        argumentBuilder.addText("Text argument builder");

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_INCLUDE_TEXT);

        try
        {
            fieldBuilder.addArgument(argumentBuilder).addArgument("=").addArgument("BestField").addArgument(10).addArgument(20.0).buildAndInsert(run);
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    //For assert result of the test you need to open document and check that image are added correct and without truncated inside frame
    @Test
    public void updateFieldIgnoringMergeFormat() throws Exception
    {
        //ExStart
        //ExFor:Field.Update(bool)
        //ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setPreserveIncludePictureField(true);

        Document doc = new Document(getMyDir() + "Field.UpdateFieldIgnoringMergeFormat.docx", loadOptions);

        for (Field field : doc.getRange().getFields())
        {
            if (((field.getType()) == (FieldType.FIELD_INCLUDE_PICTURE)))
            {
                FieldIncludePicture includePicture = (FieldIncludePicture) field;

                includePicture.setSourceFullName(getMyDir() + "\\Images\\dotnet-logo.png");
                includePicture.update(true);
            }
        }

        doc.updateFields();
        doc.save(getMyDir() + "\\Artifacts\\Field.UpdateFieldIgnoringMergeFormat.docx");
        //ExEnd
    }

    @Test
    public void fieldFormat() throws Exception
    {
        //ExStart
        //ExFor:Field.Format
        //ExFor:FieldFormat
        //ExFor:FieldFormat.DateTimeFormat
        //ExFor:FieldFormat.NumericFormat
        //ExFor:FieldFormat.GeneralFormats
        //ExFor:GeneralFormat
        //ExFor:GeneralFormatCollection.Add(GeneralFormat)
        //ExSummary:Shows how to formatting fields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("MERGEFIELD Date");

        FieldFormat format = field.getFormat();

        format.setDateTimeFormat("dddd, MMMM dd, yyyy");
        format.setNumericFormat("0.#");
        format.getGeneralFormats().add(GeneralFormat.CHAR_FORMAT);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        field = doc.getRange().getFields().get(0);
        format = field.getFormat();

        Assert.assertEquals(format.getNumericFormat(), "0.#");
        Assert.assertEquals(format.getDateTimeFormat(), "dddd, MMMM dd, yyyy");
        Assert.assertEquals(format.getGeneralFormats().get(0), GeneralFormat.CHAR_FORMAT);
    }

    @Test
    public void unlinkAllFieldsInDocument() throws Exception
    {
        //ExStart
        //ExFor:Document.UnlinkFields
        //ExSummary:Shows how to unlink all fields in the document
        Document doc = new Document(getMyDir() + "Field.UnlinkFields.docx");

        doc.unlinkFields();
        //ExEnd

        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);
        Assert.assertEquals(paraWithFields, "Fields.Docx   Элементы указателя не найдены.     1.\r");
    }

    @Test
    public void unlinkAllFieldsInRange() throws Exception
    {
        //ExStart
        //ExFor:Range.UnlinkFields
        //ExSummary:Shows how to unlink all fields in range
        Document doc = new Document(getMyDir() + "Field.UnlinkFields.docx");

        Section newSection = doc.getSections().get(0).deepClone();
        doc.getSections().add(newSection);

        doc.getSections().get(1).getRange().unlinkFields();
        //ExEnd

        String secWithFields = DocumentHelper.getSectionText(doc, 1);
        Assert.assertEquals("Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4.\r\r\r\r\r\f", secWithFields);
    }

    @Test
    public void unlinkSingleField() throws Exception
    {
        //ExStart
        //ExFor:Field.Unlink
        //ExSummary:Shows how to unlink specific field
        Document doc = new Document(getMyDir() + "Field.UnlinkFields.docx");

        doc.getRange().getFields().get(1).unlink();
        //ExEnd

        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);
        Assert.assertEquals("\u0013 FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015\r", paraWithFields);
    }

    @Test
    public void updatePageNumbersInToc() throws Exception
    {
        Document doc = new Document(getMyDir() + "Field.UpdateTocPages.docx");

        Node startNode = DocumentHelper.getParagraph(doc, 2);
        Node endNode = null;

        NodeCollection paragraphCollection = doc.getChildNodes(NodeType.PARAGRAPH, true);

        for (Paragraph para : (Iterable<Paragraph>) paragraphCollection)
        {
            // Check all runs in the paragraph for the first page breaks.
            for (Run run : para.getRuns())
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

        for (FieldStart field : (Iterable<FieldStart>) fStart)
        {
            int fType = field.getFieldType();
            if (fType == FieldType.FIELD_TOC)
            {
                Paragraph para = (Paragraph) field.getAncestor(NodeType.PARAGRAPH);
                para.getRange().updateFields();
                break;
            }
        }

        doc.save(getMyDir() + "\\Artifacts\\Field.UpdateTocPages.docx");
    }

    private void removeSequence(Node start, Node end)
    {
        Node curNode = start.nextPreOrder(start.getDocument());
        while (curNode != null && !curNode.equals(end))
        {
            //Move to next node
            Node nextNode = curNode.nextPreOrder(start.getDocument());

            //Check whether current contains end node
            if (curNode.isComposite())
            {
                CompositeNode curComposite = (CompositeNode) curNode;
                if (!curComposite.getChildNodes(NodeType.ANY, true).contains(end) && !curComposite.getChildNodes(NodeType.ANY, true).contains(start))
                {
                    nextNode = curNode.getNextSibling();
                    curNode.remove();
                }
            } else
            {
                curNode.remove();
            }

            curNode = nextNode;
        }
    }
}
