package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.List;
import com.aspose.words.Shape;
import com.aspose.words.net.System.Data.DataColumn;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataTable;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.text.MessageFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashMap;
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
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        FieldChar fieldStart = (FieldChar) doc.getChild(NodeType.FIELD_START, 0, true);
        Assert.assertEquals(fieldStart.getFieldType(), FieldType.FIELD_TOC);
        Assert.assertEquals(fieldStart.isDirty(), true);
        Assert.assertEquals(fieldStart.isLocked(), false);

        // Retrieve the facade object which represents the field in the document.
        Field field = fieldStart.getField();

        Assert.assertEquals(false, field.isLocked());
        Assert.assertEquals(" TOC \\o \"1-3\" \\h \\z \\u ", field.getFieldCode());

        // This updates only this field in the document.
        field.update();
        //ExEnd
    }

    @Test
    public void createRevNumFieldWithFieldBuilder() throws Exception {
        //ExStart
        //ExFor:FieldBuilder.#ctor(FieldType)
        //ExFor:FieldBuilder.BuildAndInsert(Inline)
        //ExFor:FieldRevNum
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
    public void createRevNumFieldByDocumentBuilder() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("REVNUM MERGEFORMAT");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        FieldRevNum revNum = (FieldRevNum) doc.getRange().getFields().get(0);
        Assert.assertNotNull(revNum);
    }

    @Test
    public void createInfoFieldWithFieldBuilder() throws Exception {
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
    public void createInfoFieldWithDocumentBuilder() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("INFO MERGEFORMAT");

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        FieldInfo info = (FieldInfo) doc.getRange().getFields().get(0);
        Assert.assertNotNull(info);
    }

    @Test
    public void getFieldFromFieldCollection() throws Exception {
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        Field field = doc.getRange().getFields().get(0);

        // This should be the first field in the document - a TOC field.
        System.out.println(field.getType());
    }

    @Test
    public void insertFieldNone() throws Exception {
        //ExStart
        //ExFor:FieldUnknown
        //ExSummary:Shows how to work with 'FieldNone' field in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(FieldType.FIELD_NONE, false);

        ByteArrayOutputStream stream = new ByteArrayOutputStream();
        doc.save(stream, SaveFormat.DOCX);

        FieldCollection fieldCollection = doc.getRange().getFields();

        for (Field field : fieldCollection) {
            if (field.getType() != FieldType.FIELD_NONE) {
                Assert.fail("FieldUnknown doesn't exist");
            }
        }
        //ExEnd
    }

    @Test
    public void insertTCField() throws Exception {
        // Create a blank document.
        Document doc = new Document();

        // Create a document builder to insert content with.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TC field at the current document builder position.
        builder.insertField("TC \"Entry Text\" \\f t");
    }

    @Test
    public void changeLocale() throws Exception {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete.
        Locale currentCulture = Locale.getDefault();
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        Locale.setDefault(new Locale("de", "DE"));

        // Execute mail merge
        doc.getMailMerge().execute(new String[]{"Date"}, new Object[]{new Date()});

        // Restore the original culture.
        Locale.setDefault(currentCulture);

        doc.save(getArtifactsDir() + "Field.ChangeLocale.doc");
    }

    //ExStart
    //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
    //ExSummary:Demonstrates how to remove a specified TOC from a document.
    @Test //ExSkip
    public void removeTOCFromDocument() throws Exception {
        // Open a document which contains a TOC.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        // Remove the first TOC from the document.
        Field tocField = doc.getRange().getFields().get(0);
        tocField.remove();

        // Save the output.
        doc.save(getArtifactsDir() + "Document.TableOfContentsRemoveTOC.doc");
    }
    //ExEnd

    @Test //ExSkip
    public void insertTCFieldsAtText() throws Exception {
        Document doc = new Document();

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new InsertTCFieldHandler("Chapter 1", "\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc.getRange().replace(Pattern.compile("The Beginning"), "", options);
    }

    public class InsertTCFieldHandler implements IReplacingCallback {
        // Store the text and switches to be used for the TC fields.
        private String mFieldText;
        private String mFieldSwitches;

        /**
         * The switches to use for each TC field. Can be an empty string or null.
         */
        public InsertTCFieldHandler(final String switches) throws Exception {
            this(null, switches);
        }

        /**
         * The display text and switches to use for each TC field. Display name can be an empty String or null.
         */
        public InsertTCFieldHandler(final String text, final String switches) {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        public int replacing(final ReplacingArgs args) throws Exception {
            // Create a builder to insert the field.
            DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
            // Move to the first node of the match.
            builder.moveTo(args.getMatchNode());

            // If the user specified text to be used in the field as display text then use that, otherwise use the
            // match string as the display text.
            String insertText;

            if (!(mFieldText == null || "".equals(mFieldText))) {
                insertText = mFieldText;
            } else {
                insertText = args.getMatch().group();
            }

            // Insert the TC field before this node using the specified string as the display text and user defined switches.
            builder.insertField(MessageFormat.format("TC \"{0}\" {1}", insertText, mFieldSwitches));

            // We have done what we want so skip replacement.
            return ReplaceAction.SKIP;
        }
    }

    @Test(enabled = false, description = "WORDSNET-16037")
    public void insertAndUpdateDirtyField() throws Exception {
        //ExStart
        //ExFor:Field.IsDirty
        //ExFor:LoadOptions.UpdateDirtyFields
        //ExSummary:Shows how to use special property for updating field result
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field fieldToc = builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        fieldToc.isDirty(true);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);
        // Assert that field model is correct
        Assert.assertTrue(doc.getRange().getFields().get(0).isDirty());

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setUpdateDirtyFields(false);

        ByteArrayInputStream dstInputStream = new ByteArrayInputStream(dstStream.toByteArray());
        doc = new Document(dstInputStream);
        Field tocField = doc.getRange().getFields().get(0);
        // Assert that isDirty saves
        Assert.assertTrue(tocField.isDirty());
    }

    @Test
    public void insertFieldWithFieldBuilder() throws Exception {
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

        // Add text into the paragraph
        Paragraph para = doc.getFirstSection().getBody().getParagraphs().get(0);
        Run run = new Run(doc);
        run.setText(" Hello World!");

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
    public void insertFieldWithFieldBuilderException() throws Exception {
        Document doc = new Document();

        // Add some text into the paragraph
        Run run = DocumentHelper.insertNewRun(doc, " Hello World!", 0);

        FieldArgumentBuilder argumentBuilder = new FieldArgumentBuilder();
        argumentBuilder.addField(new FieldBuilder(FieldType.FIELD_MERGE_FIELD));
        argumentBuilder.addNode(run);
        argumentBuilder.addText("Text argument builder");

        FieldBuilder fieldBuilder = new FieldBuilder(FieldType.FIELD_INCLUDE_TEXT);

        try {
            fieldBuilder.addArgument(argumentBuilder).addArgument("=").addArgument("BestField").addArgument(10).addArgument(20.0).buildAndInsert(run);
        } catch (Exception e) {
            Assert.assertTrue(e instanceof IllegalArgumentException);
        }
    }

    //For assert result of the test you need to open document and check that image are added correct and without truncated inside frame
    @Test
    public void updateFieldIgnoringMergeFormat() throws Exception {
        //ExStart
        //ExFor:Field.Update(bool)
        //ExFor:LoadOptions.PreserveIncludePictureField
        //ExSummary:Shows a way to update a field ignoring the MERGEFORMAT switch
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setPreserveIncludePictureField(true);

        Document doc = new Document(getMyDir() + "Field.UpdateFieldIgnoringMergeFormat.docx", loadOptions);

        for (Field field : doc.getRange().getFields()) {
            if (((field.getType()) == (FieldType.FIELD_INCLUDE_PICTURE))) {
                FieldIncludePicture includePicture = (FieldIncludePicture) field;
                includePicture.setSourceFullName(getImageDir() + "dotnet-logo.png");
                includePicture.update(true);
            }
        }

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.UpdateFieldIgnoringMergeFormat.docx");
        //ExEnd
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
        //ExSummary:Shows how to formatting fields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert field with no format
        Field field = builder.insertField("= 2 + 3");

        // We can format our field here instead of in the field code
        FieldFormat format = field.getFormat();
        format.setNumericFormat("$###.00");
        field.update();

        // Apply a date/time format
        field = builder.insertField("DATE");
        format = field.getFormat();
        format.setDateTimeFormat("dddd, MMMM dd, yyyy");
        field.update();

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
    }

    @Test
    public void unlinkAllFieldsInDocument() throws Exception {
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
    public void unlinkAllFieldsInRange() throws Exception {
        //ExStart
        //ExFor:Range.UnlinkFields
        //ExSummary:Shows how to unlink all fields in range
        Document doc = new Document(getMyDir() + "Field.UnlinkFields.docx");

        Section newSection = doc.getSections().get(0).deepClone();
        doc.getSections().add(newSection);

        doc.getSections().get(1).getRange().unlinkFields();
        //ExEnd

        String secWithFields = DocumentHelper.getSectionText(doc, 1);
        Assert.assertEquals(secWithFields, "Fields.Docx   Элементы указателя не найдены.     3.\rОшибка! Не указана последовательность.    Fields.Docx   Элементы указателя не найдены.     4.\r\r\r\r\r\f");
    }

    @Test
    public void unlinkSingleField() throws Exception {
        //ExStart
        //ExFor:Field.Unlink
        //ExSummary:Shows how to unlink specific field
        Document doc = new Document(getMyDir() + "Field.UnlinkFields.docx");

        doc.getRange().getFields().get(1).unlink();
        //ExEnd

        String paraWithFields = DocumentHelper.getParagraphText(doc, 0);
        Assert.assertEquals(paraWithFields, "\u0013 FILENAME  \\* Caps  \\* MERGEFORMAT \u0014Fields.Docx\u0015   Элементы указателя не найдены.     \u0013 LISTNUM  LegalDefault \u0015\r");
    }

    @Test
    public void updatePageNumbersInToc() throws Exception {
        Document doc = new Document(getMyDir() + "Field.UpdateTocPages.docx");

        Node startNode = DocumentHelper.getParagraph(doc, 2);
        Node endNode = null;

        NodeCollection paragraphCollection = doc.getChildNodes(NodeType.PARAGRAPH, true);

        for (Paragraph para : (Iterable<Paragraph>) paragraphCollection) {
            // Check all runs in the paragraph for the first page breaks.
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

        doc.save(getArtifactsDir() + "Field.UpdateTocPages.docx");
    }

    private void removeSequence(final Node start, final Node end) {
        Node curNode = start.nextPreOrder(start.getDocument());
        while (curNode != null && !curNode.equals(end)) {
            //Move to next node
            Node nextNode = curNode.nextPreOrder(start.getDocument());

            //Check whether current contains end node
            if (curNode.isComposite()) {
                CompositeNode curComposite = (CompositeNode) curNode;
                if (!curComposite.getChildNodes(NodeType.ANY, true).contains(end) && !curComposite.getChildNodes(NodeType.ANY, true).contains(start)) {
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

        doc.save(getArtifactsDir() + "Fields.DropDownItems.docx");
        //ExEnd

        // Empty the collection
        dropDownItems.clear();
        Assert.assertEquals(dropDownItems.getCount(), 0);
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
        doc.save(getArtifactsDir() + "Fields.AskField.docx");

        Assert.assertEquals(fieldAsk.getFieldCode(),
                " ASK  MyAskField \"Please provide a response for this ASK field\" \\d \"Response from within the field.\" \\o");

        Assert.assertEquals(fieldAsk.getBookmarkName(), "MyAskField"); //ExSkip
        Assert.assertEquals(fieldAsk.getPromptText(), "Please provide a response for this ASK field"); //ExSkip
        Assert.assertEquals(fieldAsk.getDefaultResponse(), "Response from within the field."); //ExSkip
        Assert.assertEquals(fieldAsk.getPromptOnceOnMailMerge(), true); //ExSkip
    }

    /// <summary>
    /// IFieldUserPromptRespondent implementation that appends a line to the default response of an ASK field during a mail merge
    /// </summary>
    private static class MyPromptRespondent implements IFieldUserPromptRespondent {
        public String respond(final String promptText, final String defaultResponse) {
            return "Response from MyPromptRespondent. " + defaultResponse;
        }
    }
    //ExEnd

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

        Assert.assertEquals(field.getType(), FieldType.FIELD_ADVANCE);
        Assert.assertEquals(field.getFieldCode(), " ADVANCE ");
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

        doc.save(getArtifactsDir() + "Field.Advance.docx");
        //ExEnd
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
        Assert.assertEquals(field.getFieldCode(),
                " ADDRESSBLOCK  \\c 2 \\d \\e \"United States\" \\f \"<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>\" \\l 1033");

        //ExEnd
        Assert.assertEquals(field.getIncludeCountryOrRegionName(), "2");
        Assert.assertEquals(field.getFormatAddressOnCountryOrRegion(), true);
        Assert.assertEquals(field.getExcludedCountryOrRegionName(), "United States");
        Assert.assertEquals(field.getNameAndAddressFormat(), "<Title> <Forename> <Surname> <Address Line 1> <Region> <Postcode> <Country>");
        Assert.assertEquals(field.getLanguageId(), "1033");
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
        // Open a document that has fields
        Document doc = new Document(getMyDir() + "Document.ContainsFields.docx");

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
    /// Document visitor implementation that prints field info
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
        doc.save(getArtifactsDir() + "Field.Compare.docx");
        //ExEnd
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
        FieldIf fieldIf = (FieldIf) builder.insertField(FieldType.FIELD_IF, true);

        // The if field will output either the TrueText or FalseText string into the document, depending on the truth of the statement
        // In this case, "0 = 1" is incorrect, so the output will be "False"
        fieldIf.setLeftExpression("0");
        fieldIf.setComparisonOperator("=");
        fieldIf.setRightExpression("1");
        fieldIf.setTrueText("True");
        fieldIf.setFalseText("False");

        Assert.assertEquals(fieldIf.getFieldCode(), " IF  0 = 1 True False");
        Assert.assertEquals(fieldIf.evaluateCondition(), FieldIfComparisonResult.FALSE);

        // This time, the statement is correct, so the output will be "True"
        builder.write("\nStatement 2: ");
        fieldIf = (FieldIf) builder.insertField(FieldType.FIELD_IF, true);
        fieldIf.setLeftExpression("5");
        fieldIf.setComparisonOperator("=");
        fieldIf.setRightExpression("2 + 3");
        fieldIf.setTrueText("True");
        fieldIf.setFalseText("False");

        Assert.assertEquals(fieldIf.getFieldCode(), " IF  5 = \"2 + 3\" True False");
        Assert.assertEquals(fieldIf.evaluateCondition(), FieldIfComparisonResult.TRUE);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.If.docx");
        //ExEnd
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
        builder.insertField(FieldType.FIELD_AUTO_NUM, true);
        builder.writeln("\tParagraph 1.");
        builder.insertField(FieldType.FIELD_AUTO_NUM, true);
        builder.writeln("\tParagraph 2.");

        for (Field field : doc.getRange().getFields()) {
            if (field.getType() == FieldType.FIELD_AUTO_NUM) {
                // Leaving the FieldAutoNum.SeparatorCharacter field null will set the separator character to '.' by default
                Assert.assertNull(((FieldAutoNum) field).getSeparatorCharacter());

                // The first character of the string entered here will be used as the separator character
                ((FieldAutoNum) field).setSeparatorCharacter(":");

                Assert.assertEquals(field.getFieldCode(), " AUTONUM  \\s :");
            }
        }

        doc.save(getArtifactsDir() + "Field.AutoNum.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:FieldAutoNumLgl
    //ExFor:FieldAutoNumLgl.RemoveTrailingPeriod
    //ExFor:FieldAutoNumLgl.SeparatorCharacter
    //ExSummary:Shows how to organize a document using autonum legal fields
    @Test //ExSkip
    public void fieldAutoNumLgl() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // This string will be our paragraph text that
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

        doc.save(getArtifactsDir() + "Field.AutoNumLegal.docx");
    }

    /// <summary>
    /// Get a document builder to insert a clause numbered by an autonum legal field
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

    @Test
    public void fieldAutoNumOut() throws Exception {
        //ExStart
        //ExFor:FieldAutoNumOut
        //ExSummary:Shows how to number paragraphs using autonum outline fields.
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

        doc.save(getArtifactsDir() + "Field.AutoNumOut.docx");
        //ExEnd
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
        doc.getFieldOptions().setBuiltInTemplatesPaths(new String[]{getMyDir() + "Document.BusinessBrochureTemplate.dotx"});

        // We can also display our building block with a GLOSSARY field
        FieldGlossary fieldGlossary = (FieldGlossary) builder.insertField(FieldType.FIELD_GLOSSARY, true);
        fieldGlossary.setEntryName("MyBlock");

        Assert.assertEquals(fieldGlossary.getFieldCode(), " GLOSSARY  MyBlock");

        // The text content of our building block will be visible in the output
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.AutoText.dotx");
        //ExEnd
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
        field.setEntryName("Right click here to pick an AutoText block"); // This is the text that will be visible in the document
        field.setListStyle("Heading 1");
        field.setScreenTip("Hover tip text for AutoTextList goes here");

        Assert.assertEquals(field.getEntryName(), "Right click here to pick an AutoText block"); //ExSkip
        Assert.assertEquals(field.getListStyle(), "Heading 1"); //ExSkip
        Assert.assertEquals(field.getScreenTip(), "Hover tip text for AutoTextList goes here"); //ExSkip
        Assert.assertEquals(field.getFieldCode(), " AUTOTEXTLIST  \"Right click here to pick an AutoText block\" "
                + "\\s \"Heading 1\" "
                + "\\t \"Hover tip text for AutoTextList goes here\"");

        doc.save(getArtifactsDir() + "Field.AutoTextList.dotx");
    }

    /// <summary>
    /// Create an AutoText entry and add it to a glossary document
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

        // Insert a custom greeting field with document builder, and also some content
        FieldGreetingLine fieldGreetingLine = (FieldGreetingLine) builder.insertField(FieldType.FIELD_GREETING_LINE, true);
        builder.writeln("\n\n\tThis is your custom greeting, created programmatically using Aspose Words!");

        // This array contains strings that correspond to column names in the data table that we will mail merge into our document
        Assert.assertEquals(fieldGreetingLine.getFieldNames().length, 0);

        // To populate that array, we need to specify a format for our greeting line
        fieldGreetingLine.setNameFormat("<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> ");

        // In this case, our greeting line's field names array now has "Courtesy Title" and "Last Name"
        Assert.assertEquals(fieldGreetingLine.getFieldNames().length, 2);

        // This string will cover any cases where the data in the data table is incorrect by substituting the malformed name with a string
        fieldGreetingLine.setAlternateText("Sir or Madam");

        // We can set the language ID here too
        fieldGreetingLine.setLanguageId("1033");

        Assert.assertEquals(fieldGreetingLine.getFieldCode(), " GREETINGLINE  \\f \"<< _BEFORE_ Dear >><< _TITLE0_ >><< _LAST0_ >><< _AFTER_ ,>> \" \\e \"Sir or Madam\" \\l 1033");

        // Create a source table for our mail merge that has columns that our greeting line will look for
        DataTable table = new DataTable("Employees");
        table.getColumns().add("Courtesy Title");
        table.getColumns().add("First Name");
        table.getColumns().add("Last Name");
        table.getRows().add("Mr.", "John", "Doe");
        table.getRows().add("Mrs.", "Jane", "Cardholder");
        table.getRows().add("", "No", "Name"); // This row has an invalid value in the Courtesy Title column, so our greeting will default to the alternate text

        doc.getMailMerge().execute(table);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.GreetingLine.docx");
        //ExEnd
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
        FieldListNum fieldListNum = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // Lists start counting at 1 by default, but we can change this number at any time
        // In this case, we'll do a zero-based count
        fieldListNum.setStartingNumber("0");
        builder.writeln("Paragraph 1");

        // Placing several list num fields in one paragraph increases the list level instead of the current number, in this case resulting in "1)a)i)", list level 3
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.insertField(FieldType.FIELD_LIST_NUM, true);
        builder.writeln("Paragraph 2");

        // The list level resets with new paragraphs, so to keep counting at a desired list level, we need to set the ListLevel property accordingly
        fieldListNum = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);
        fieldListNum.setListLevel("3");
        builder.writeln("Paragraph 3");

        fieldListNum = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);

        // Setting this property to this particular value will emulate the AUTONUMOUT field
        fieldListNum.setListName("OutlineDefault");
        Assert.assertTrue(fieldListNum.hasListName());

        // Start counting from 1
        fieldListNum.setStartingNumber("1");
        builder.writeln("Paragraph 4");

        // Our fields keep track of the count automatically, but the ListName needs to be set with each new field
        fieldListNum = (FieldListNum) builder.insertField(FieldType.FIELD_LIST_NUM, true);
        fieldListNum.setListName("OutlineDefault");
        builder.writeln("Paragraph 5");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FieldListNum.docx");
        //ExEnd
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

        // Insert another merge field for another column
        // We don't need to use every column to perform a mail merge
        fieldMergeField = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        fieldMergeField.setFieldName("Last Name");
        fieldMergeField.setTextAfter(":");

        doc.updateFields();
        doc.getMailMerge().execute(table);
        doc.save(getArtifactsDir() + "Field.MergeField.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:FormField.Accept(DocumentVisitor)
    //ExFor:FormField.CalculateOnExit
    //ExFor:FormField.CheckBoxSize
    //ExFor:FormField.Checked
    //ExFor:FormField.Default
    //ExFor:FormField.DropDownItems
    //ExFor:FormField.DropDownSelectedIndex
    //ExFor:FormField.Enabled
    //ExFor:FormField.EntryMacro
    //ExFor:FormField.ExitMacro
    //ExFor:FormField.HelpText
    //ExFor:FormField.IsCheckBoxExactSize
    //ExFor:FormField.MaxLength
    //ExFor:FormField.OwnHelp
    //ExFor:FormField.OwnStatus
    //ExFor:FormField.SetTextInputValue(Object)
    //ExFor:FormField.StatusText
    //ExFor:FormField.TextInputDefault
    //ExFor:FormField.TextInputFormat
    //ExFor:FormField.TextInputType
    //ExFor:FormFieldCollection.Clear
    //ExFor:FormFieldCollection.Count
    //ExFor:FormFieldCollection.GetEnumerator
    //ExFor:FormFieldCollection.Item(Int32)
    //ExFor:FormFieldCollection.Item(String)
    //ExFor:FormFieldCollection.Remove(String)
    //ExFor:FormFieldCollection.RemoveAt(Int32)
    //ExSummary:Shows how insert different kinds of form fields into a document and process them with a visitor implementation.
    @Test //ExSkip
    public void formField() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a combo box
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[]{"One", "Two", "Three"}, 0);
        comboBox.setCalculateOnExit(true);
        Assert.assertEquals(comboBox.getDropDownItems().getCount(), 3);
        Assert.assertEquals(comboBox.getDropDownSelectedIndex(), 0);
        Assert.assertEquals(comboBox.getEnabled(), true);

        builder.writeln();

        // Use a document builder to insert a check box
        FormField checkBox = builder.insertCheckBox("MyCheckBox", false, 50);
        checkBox.isCheckBoxExactSize(true);
        checkBox.setHelpText("Right click to check this box");
        checkBox.setOwnHelp(true);
        checkBox.setStatusText("Checkbox status text");
        checkBox.setOwnStatus(true);
        Assert.assertEquals(checkBox.getCheckBoxSize(), 50.0d);
        Assert.assertEquals(checkBox.getChecked(), false);
        Assert.assertEquals(checkBox.getDefault(), false);

        builder.writeln();

        // Use a document builder to insert text input form field
        FormField textInput = builder.insertTextInput("MyTextInput", TextFormFieldType.REGULAR, "", "Your text goes here", 50);
        Assert.assertEquals(doc.getRange().getFields().getCount(), 3);
        textInput.setEntryMacro("EntryMacro");
        textInput.setExitMacro("ExitMacro");
        textInput.setTextInputDefault("Regular");
        textInput.setTextInputFormat("FIRST CAPITAL");
        textInput.setTextInputValue("This value overrides the one we set during initialization");
        Assert.assertEquals(textInput.getTextInputType(), TextFormFieldType.REGULAR);
        Assert.assertEquals(textInput.getMaxLength(), 50);

        // Get the collection of form fields that has accumulated in our document
        FormFieldCollection formFields = doc.getRange().getFormFields();
        Assert.assertEquals(formFields.getCount(), 3);

        // Iterate over the collection with an enumerator, accepting a visitor with each form field
        FormFieldVisitor formFieldVisitor = new FormFieldVisitor();

        Iterator<FormField> fieldEnumerator = formFields.iterator();
        while (fieldEnumerator.hasNext()) {
            fieldEnumerator.next().accept(formFieldVisitor);
        }

        System.out.println(formFieldVisitor.getText());

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FormField.docx");
    }

    /// <summary>
    /// Visitor implementation that prints information about visited form fields.
    /// </summary>
    public static class FormFieldVisitor extends DocumentVisitor {
        public FormFieldVisitor() {
            mBuilder = new StringBuilder();
        }

        /// <summary>
        /// Called when a FormField node is encountered in the document.
        /// </summary>
        public int visitFormField(final FormField formField) {
            appendLine(formField.getType() + ": \"" + formField.getName() + "\"");
            appendLine("\tStatus: " + (formField.getEnabled() ? "Enabled" : "Disabled"));
            appendLine("\tHelp Text:  " + formField.getHelpText());
            appendLine("\tEntry macro name: " + formField.getEntryMacro());
            appendLine("\tExit macro name: " + formField.getExitMacro());

            switch (formField.getType()) {
                case FieldType.FIELD_FORM_DROP_DOWN:
                    appendLine("\tDrop down items count: " + formField.getDropDownItems().getCount() + ", default selected item index: " + formField.getDropDownSelectedIndex());
                    appendLine("\tDrop down items: " + String.join(", ", formField.getDropDownItems()));
                    break;
                case FieldType.FIELD_FORM_CHECK_BOX:
                    appendLine("\tCheckbox size: " + formField.getCheckBoxSize());
                    appendLine("\t" + "Checkbox is currently: " + (formField.getChecked() ? "checked, " : "unchecked, ") + "by default: " + (formField.getDefault() ? "checked" : "unchecked"));
                    break;
                case FieldType.FIELD_FORM_TEXT_INPUT:
                    appendLine("\tInput format: " + formField.getTextInputFormat());
                    appendLine("\tCurrent contents: " + formField.getResult());
                    break;
            }

            // Let the visitor continue visiting other nodes.
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Adds newline char-terminated text to the current output.
        /// </summary>
        private void appendLine(final String text) {
            mBuilder.append(text + '\n');
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        private StringBuilder mBuilder;
    }
    //ExEnd

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
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);

        // Limit possible TOC entries to only those within the bookmark we name here
        fieldToc.setBookmarkName("MyBookmark");

        // Normally paragraphs with a "Heading n" style will be the only ones that will be added to a TOC as entries
        // We can set this attribute to include other styles, such as "Quote" and "Intense Quote" in this case
        fieldToc.setCustomStyles("Quote; 6; Intense Quote; 7");

        // Styles are normally separated by a comma (",") but we can use this property to set a custom delimiter
        doc.getFieldOptions().setCustomTocStyleSeparator(";");

        // Filter out any headings that are outside this range
        fieldToc.setHeadingLevelRange("1-3");

        // Headings in this range won't display their page number in their TOC entry
        fieldToc.setPageNumberOmittingLevelRange("2-5");

        fieldToc.setEntrySeparator("-");
        fieldToc.setInsertHyperlinks(true);
        fieldToc.setHideInWebLayout(false);
        fieldToc.setPreserveLineBreaks(true);
        fieldToc.setPreserveTabs(true);
        fieldToc.setUseParagraphOutlineLevel(false);

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

        Assert.assertEquals(fieldToc.getFieldCode(), " TOC  \\b MyBookmark \\t \"Quote; 6; Intense Quote; 7\" \\o 1-3 \\n 2-5 \\p - \\h \\x \\w");

        fieldToc.updatePageNumbers();
        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FieldTOC.docx");
    }

    /// <summary>
    /// Start a new page and insert a paragraph of a specified style
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

        builder.insertBreak(BreakType.PAGE_BREAK);

        // These two entries will appear in the table
        insertTocEntry(builder, "TC field 1", "A", "1");
        insertTocEntry(builder, "TC field 2", "A", "2");

        // These two entries will be omitted because of an incorrect type identifier
        insertTocEntry(builder, "TC field 3", "B", "1");

        // ...and an out-of-range entry level
        insertTocEntry(builder, "TC field 4", "A", "5");

        Assert.assertEquals(fieldToc.getFieldCode(), " TOC  \\f A \\l 1-3");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FieldTOC.TC.docx");
    }

    /// <summary>
    /// Insert a table of contents entry via a document builder
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

    //ExStart
    //ExFor:FieldToc.TableOfFiguresLabel
    //ExFor:FieldToc.CaptionlessTableOfFiguresLabel
    //ExFor:FieldToc.PrefixedSequenceIdentifier
    //ExFor:FieldToc.SequenceSeparator
    //ExFor:FieldSeq
    //ExFor:FieldSeq.BookmarkName
    //ExFor:FieldSeq.InsertNextNumber
    //ExFor:FieldSeq.ResetHeadingLevel
    //ExFor:FieldSeq.ResetNumber
    //ExFor:FieldSeq.SequenceIdentifier
    //ExSummary:Insert a TOC field and build the table with SEQ fields.
    @Test //ExSkip
    public void tocSeqPrefix() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Filter by sequence identifier and a prefix sequence identifier, and change sequence separator
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);
        fieldToc.setTableOfFiguresLabel("MySequence");
        fieldToc.setPrefixedSequenceIdentifier("PrefixSequence");
        fieldToc.setSequenceSeparator(">");

        Assert.assertEquals(fieldToc.getFieldCode(), " TOC  \\c MySequence \\s PrefixSequence \\d >");

        builder.insertBreak(BreakType.PAGE_BREAK);

        // Add two SEQ fields in one paragraph, setting the TOC's sequence and prefix sequence as their sequence identifiers
        FieldSeq fieldSeq = insertSeqField(builder, "PrefixSequence ", "", "PrefixSequence");
        Assert.assertEquals(fieldSeq.getFieldCode(), " SEQ  PrefixSequence");

        fieldSeq = insertSeqField(builder, ", MySequence ", "\n", "MySequence");
        Assert.assertEquals(fieldSeq.getFieldCode(), " SEQ  MySequence");

        insertSeqField(builder, "PrefixSequence ", "", "PrefixSequence");
        insertSeqField(builder, ", MySequence ", "\n", "MySequence");

        // If the sqeuence identifier doesn't match that of the TOC, the entry won't be included
        insertSeqField(builder, "PrefixSequence ", "", "PrefixSequence");
        fieldSeq = insertSeqField(builder, ", MySequence ", "", "OtherSequence");
        builder.writeln(" This text, from a different sequence, won't be included in the same TOC as the one above.");

        Assert.assertEquals(fieldSeq.getFieldCode(), " SEQ  OtherSequence");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOC.SEQ.Prefix.docx");
    }

    @Test(enabled = false, description = "WORDSNET-18083") //ExSkip
    public void tocSeqNumbering() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Filter by sequence identifier and a prefix sequence identifier, and change sequence separator
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);
        fieldToc.setTableOfFiguresLabel("MySequence");

        Assert.assertEquals(fieldToc.getFieldCode(), " TOC  \\c MySequence");

        builder.insertBreak(BreakType.PAGE_BREAK);

        // Set the current number of the sequence to 100
        FieldSeq fieldSeq = insertSeqField(builder, "MySequence ", "\n", "MySequence");
        fieldSeq.setResetNumber("100");
        Assert.assertEquals(fieldSeq.getFieldCode(), " SEQ  MySequence \\r 100");

        insertSeqField(builder, "MySequence ", "\n", "MySequence");

        // Insert a heading
        builder.insertBreak(BreakType.PARAGRAPH_BREAK);
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("My heading");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));

        // Reset sequence when we encounter a heading, resetting the sequence back to 1
        fieldSeq = insertSeqField(builder, "MySequence ", "\n", "MySequence");
        fieldSeq.setResetHeadingLevel("1");
        Assert.assertEquals(" SEQ  MySequence \\s 1", fieldSeq.getFieldCode());

        // Move to the next number
        fieldSeq = insertSeqField(builder, "MySequence ", "\n", "MySequence");
        fieldSeq.setInsertNextNumber(true);
        Assert.assertEquals(fieldSeq.getFieldCode(), " SEQ  MySequence \\n");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOC.SEQ.ResetNumbering.docx");
    }

    @Test(enabled = false, description = "WORDSNET-18084") //ExSkip
    public void tocSeqBookmark() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // This TOC takes in all SEQ fields with "MySequence" inside "TOCBookmark"
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);
        fieldToc.setTableOfFiguresLabel("MySequence");
        fieldToc.setBookmarkName("TOCBookmark");
        builder.insertBreak(BreakType.PAGE_BREAK);

        Assert.assertEquals(fieldToc.getFieldCode(), " TOC  \\c MySequence \\b TOCBookmark");

        insertSeqField(builder, "MySequence ", "", "MySequence");
        builder.writeln(" This text won't show up in the TOC because it is outside of the bookmark.");

        builder.startBookmark("TOCBookmark");

        insertSeqField(builder, "MySequence ", "", "MySequence");
        builder.writeln(" This text will show up in the TOC next to the entry for the above caption.");

        insertSeqField(builder, "OtherSequence ", "", "OtherSequence");
        builder.writeln(" This text, from a different sequence, won't be included in the same TOC as the one above.");

        // The contents of the bookmark we reference here will not appear at the SEQ field, but will appear in the corresponding TOC entry
        FieldSeq fieldSeq = insertSeqField(builder, " MySequence ", "\n", "MySequence");
        fieldSeq.setBookmarkName("SEQBookmark");
        Assert.assertEquals(fieldSeq.getFieldCode(), " SEQ  MySequence SEQBookmark");

        // Add bookmark to reference
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.startBookmark("SEQBookmark");
        insertSeqField(builder, " MySequence ", "", "MySequence");
        builder.writeln(" Text inside SEQBookmark.");
        builder.endBookmark("SEQBookmark");

        builder.endBookmark("TOCBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TOC.SEQ.Bookmark.docx");
    }

    /// <summary>
    /// Insert a sequence field with preceding text and a specified sequence identifier
    /// </summary>
    @Test(enabled = false)
    public FieldSeq insertSeqField(final DocumentBuilder builder, final String textBefore, final String textAfter, final String sequenceIdentifier) throws Exception {
        builder.write(textBefore);
        FieldSeq fieldSeq = (FieldSeq) builder.insertField(FieldType.FIELD_SEQUENCE, true);
        fieldSeq.setSequenceIdentifier(sequenceIdentifier);
        builder.write(textAfter);

        return fieldSeq;
    }
    //ExEnd

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
        Document doc = new Document(getMyDir() + "Document.HasBibliography.docx");

        // Add text that we can cite
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Text to be cited with one source.");

        // Create a citation field using the document builder
        FieldCitation field = (FieldCitation) builder.insertField(FieldType.FIELD_CITATION, true);

        // A simple citation can have just the page number and author's name
        field.setSourceTag("Book1"); // We refer to sources using their tag names
        field.setPageNumber("85");
        field.setSuppressAuthor(false);
        field.setSuppressTitle(true);
        field.setSuppressYear(true);

        Assert.assertEquals(field.getFieldCode(), " CITATION  Book1 \\p 85 \\t \\y");

        // We can make a more detailed citation and make it cite 2 sources
        builder.write("Text to be cited with two sources.");
        field = (FieldCitation) builder.insertField(FieldType.FIELD_CITATION, true);
        field.setSourceTag("Book1");
        field.setAnotherSourceTag("Book2");
        field.setFormatLanguageId("en-US");
        field.setPageNumber("19");
        field.setPrefix("Prefix ");
        field.setSuffix(" Suffix");
        field.setSuppressAuthor(false);
        field.setSuppressTitle(false);
        field.setSuppressYear(false);
        field.setVolumeNumber("VII");

        Assert.assertEquals(field.getFieldCode(), " CITATION  Book1 \\m Book2 \\l en-US \\p 19 \\f \"Prefix \" \\s \" Suffix\" \\v VII");

        // Insert a new page which will contain our bibliography
        builder.insertBreak(BreakType.PAGE_BREAK);

        // All our sources can be displayed using a BIBLIOGRAPHY field
        FieldBibliography fieldBibliography = (FieldBibliography) builder.insertField(FieldType.FIELD_BIBLIOGRAPHY, true);
        fieldBibliography.setFormatLanguageId("1124");

        Assert.assertEquals(fieldBibliography.getFieldCode(), " BIBLIOGRAPHY  \\l 1124");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.Citation.docx");
        //ExEnd
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
        FieldInclude fieldInclude = (FieldInclude) builder.insertField(FieldType.FIELD_INCLUDE, true);
        fieldInclude.setSourceFullName(getMyDir() + "Field.Include.Source.docx");
        fieldInclude.setBookmarkName("Source_paragraph_2");
        fieldInclude.setLockFields(false);
        fieldInclude.setTextConverter("Microsoft Word");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.Include.docx");
        //ExEnd
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
        field.setFileName(getMyDir() + "Database\\Northwind.mdb");
        field.setConnection("DSN=MS Access Databases");
        field.setQuery("SELECT * FROM [Products]");

        // Insert another database field
        field = (FieldDatabase) builder.insertField(FieldType.FIELD_DATABASE, true);
        field.setFileName(getMyDir() + "Database\\Northwind.mdb");
        field.setConnection("DSN=MS Access Databases");

        // This query will sort all the products by their gross sales in descending order
        field.setQuery("SELECT [Products].ProductName, FORMAT(SUM([Order Details].UnitPrice * (1 - [Order Details].Discount) * [Order Details].Quantity), 'Currency') AS GrossSales "
                + "FROM([Products] "
                + "LEFT JOIN[Order Details] ON[Products].[ProductID] = [Order Details].[ProductID]) "
                + "GROUP BY[Products].ProductName "
                + "ORDER BY SUM([Order Details].UnitPrice* (1 - [Order Details].Discount) * [Order Details].Quantity) DESC");

        // You can use these variables instead of a LIMIT clause, to simplify your query
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
        doc.save(getArtifactsDir() + "Field.Database.docx");
        //ExEnd
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
        fieldIncludePicture.setSourceFullName(getImageDir() + "Watermark.png");

        // Here we apply the PNG32.FLT filter
        fieldIncludePicture.setGraphicFilter("PNG32");
        fieldIncludePicture.isLinked(true);
        fieldIncludePicture.setResizeHorizontally(true);
        fieldIncludePicture.setResizeVertically(true);

        // We can do the same thing with an IMPORT field
        FieldImport fieldImport = (FieldImport) builder.insertField(FieldType.FIELD_IMPORT, true);
        fieldImport.setGraphicFilter("PNG32");
        fieldImport.isLinked(true);
        fieldImport.setSourceFullName(getImageDir() + "Watermark.png");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.IncludePicture.docx");
        //ExEnd
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
    //ExSummary:Shows how to create an INCLUDETEXT field and set its properties.
    @Test(enabled = false, description = "WORDSNET-17543") //ExSkip
    public void fieldIncludeText() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert an include text field and perform an XSL transformation on an XML document
        FieldIncludeText fieldIncludeText = createFieldIncludeText(builder, getMyDir() + "Field.IncludeText.Source.xml", false, "text/xml", "XML", "ISO-8859-1");
        fieldIncludeText.setXslTransformation(getMyDir() + "Field.IncludeText.Source.xsl");

        builder.writeln();

        // Use a document builder to insert an include text field and use an XPath to take specific elements
        fieldIncludeText = createFieldIncludeText(builder, getMyDir() + "Field.IncludeText.Source.xml", false, "text/xml", "XML", "ISO-8859-1");
        fieldIncludeText.setNamespaceMappings("xmlns:n='myNamespace'");
        fieldIncludeText.setXPath("/catalog/cd/title");

        doc.save(getArtifactsDir() + "Field.IncludeText.docx");
    }

    /// <summary>
    /// Use a document builder to insert an INCLUDETEXT field and set its properties
    /// </summary>
    @Test(enabled = false)
    public FieldIncludeText createFieldIncludeText(final DocumentBuilder builder, final String sourceFullName,
                                                   final boolean lockFields, final String mimeType, final String textConverter,
                                                   final String encoding) throws Exception {
        FieldIncludeText fieldIncludeText = (FieldIncludeText) builder.insertField(FieldType.FIELD_INCLUDE_TEXT, true);
        fieldIncludeText.setSourceFullName(sourceFullName);
        fieldIncludeText.setLockFields(lockFields);
        fieldIncludeText.setMimeType(mimeType);
        fieldIncludeText.setTextConverter(textConverter);
        fieldIncludeText.setEncoding(encoding);

        return fieldIncludeText;
    }
    //ExEnd

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
        FieldHyperlink fieldHyperlink = (FieldHyperlink) builder.insertField(FieldType.FIELD_HYPERLINK, true);

        // When link is clicked, open a document and place the cursor on the bookmarked location
        fieldHyperlink.setAddress(getMyDir() + "Field.HyperlinkDestination.docx");
        fieldHyperlink.setSubAddress("My_Bookmark");
        fieldHyperlink.setScreenTip("Open " + fieldHyperlink.getAddress() + " on bookmark " + fieldHyperlink.getSubAddress() + " in a new window");

        builder.writeln();

        // Open html file at a specific frame
        fieldHyperlink = (FieldHyperlink) builder.insertField(FieldType.FIELD_HYPERLINK, true);
        fieldHyperlink.setAddress(getMyDir() + "Field.HyperlinkDestination.html");
        fieldHyperlink.setScreenTip("Open " + fieldHyperlink.getAddress());
        fieldHyperlink.setTarget("iframe_3");
        fieldHyperlink.setOpenInNewWindow(true);
        fieldHyperlink.isImageMap(false);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.Hyperlink.docx");
        //ExEnd
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
    @Test
    public void mergeFieldImageDimension() throws Exception {
        Document doc = new Document();

        // Insert a merge field where images will be placed during the mail merge
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Image:ImageColumn");

        // Create a data table for the mail merge
        // The name of the column that contains our image filenames needs to match the name of our merge field
        DataTable dataTable = createDataTable("Images", "ImageColumn",
                new String[]
                        {
                                getImageDir() + "Aspose.Words.gif",
                                getImageDir() + "Watermark.png",
                                getImageDir() + "dotnet-logo.png"
                        });

        doc.getMailMerge().setFieldMergingCallback(new MergedImageResizer(450.0, 200.0, MergeFieldImageDimensionUnit.POINT));
        doc.getMailMerge().execute(dataTable);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.MergeFieldImageDimension.docx");
    }

    /// <summary>
    /// Creates a data table with a single column
    /// </summary>
    private DataTable createDataTable(final String tableName, final String columnName, final String[] columnContents) {
        DataTable dataTable = new DataTable(tableName);
        dataTable.getColumns().add(new DataColumn(columnName));

        for (String s : columnContents) {
            DataRow dataRow = dataTable.newRow();
            dataRow.set(0, s);
            dataTable.getRows().add(dataRow);
        }

        return dataTable;
    }

    /// <summary>
    /// Sets the size of all mail merged images to one defined width and height
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

    //ExStart
    //ExFor:ImageFieldMergingArgs.Image
    //ExSummary:Shows how to set which images to merge during the mail merge.
    @Test //ExSkip
    public void mergeFieldImages() throws Exception {
        Document doc = new Document();

        // Insert a merge field where images will be placed during the mail merge
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField("MERGEFIELD Image:ImageColumn");

        // When we merge images, our data table will normally have the full e. of the images we wish to merge
        // If this is cumbersome, we can move image filename logic to another place and populate the data table with just shorthands for images
        DataTable dataTable = createDataTable("Images", "ImageColumn",
                new String[]
                        {
                                "Aspose logo",
                                ".Net logo",
                                "Watermark"
                        });

        // A custom merging callback will contain filenames that our shorthands will refer to
        doc.getMailMerge().setFieldMergingCallback(new ImageFilenameCallback());
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "Field.MergeFieldImages.docx");
    }

    /// <summary>
    /// Image merging callback that pairs image shorthand names with filenames
    /// </summary>
    private static class ImageFilenameCallback implements IFieldMergingCallback {
        public ImageFilenameCallback() {
            imageFilenames.put("Aspose logo", getImageDir() + "Aspose.Words.gif");
            imageFilenames.put(".Net logo", getImageDir() + "dotnet-logo.png");
            imageFilenames.put("Watermark", getImageDir() + "Watermark.png");
        }

        public void fieldMerging(FieldMergingArgs e) {
            throw new UnsupportedOperationException();
        }

        public void imageFieldMerging(ImageFieldMergingArgs e) throws IOException {
            if (imageFilenames.containsKey(e.getFieldValue().toString())) {
                BufferedImage image = ImageIO.read(new File(imageFilenames.get(e.getFieldValue().toString())));
                e.setImage(image);
            }

            Assert.assertNotNull(e.getImage());
        }

        private HashMap<String, String> imageFilenames = new HashMap<>();
    }
    //ExEnd

    @Test(enabled = false, description = "WORDSNET-17524")
    public void fieldXE() throws Exception {
        //ExStart
        //ExFor:FieldIndex
        //ExFor:FieldIndex.BookmarkName
        //ExFor:FieldIndex.CrossReferenceSeparator
        //ExFor:FieldIndex.EntryType
        //ExFor:FieldIndex.HasPageNumberSeparator
        //ExFor:FieldIndex.HasSequenceName
        //ExFor:FieldIndex.Heading
        //ExFor:FieldIndex.LanguageId
        //ExFor:FieldIndex.LetterRange
        //ExFor:FieldIndex.NumberOfColumns
        //ExFor:FieldIndex.PageNumberListSeparator
        //ExFor:FieldIndex.PageNumberSeparator
        //ExFor:FieldIndex.PageRangeSeparator
        //ExFor:FieldIndex.RunSubentriesOnSameLine
        //ExFor:FieldIndex.SequenceName
        //ExFor:FieldIndex.SequenceSeparator
        //ExFor:FieldIndex.UseYomi
        //ExFor:FieldXE
        //ExFor:FieldXE.EntryType
        //ExFor:FieldXE.HasPageRangeBookmarkName
        //ExFor:FieldXE.IsBold
        //ExFor:FieldXE.IsItalic
        //ExFor:FieldXE.PageNumberReplacement
        //ExFor:FieldXE.PageRangeBookmarkName
        //ExFor:FieldXE.Text
        //ExFor:FieldXE.Yomi
        //ExSummary:Shows how to populate an index field with index entries.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an index field which will contain all the index entries
        FieldIndex index = (FieldIndex) builder.insertField(FieldType.FIELD_INDEX, true);

        // Bookmark that will encompass a section that we want to index
        String mainBookmarkName = "MainBookmark";
        builder.startBookmark(mainBookmarkName);
        index.setBookmarkName(mainBookmarkName);
        index.setCrossReferenceSeparator(":");
        index.setHeading(">");
        index.setLanguageId("1033");
        index.setLetterRange("a-j");
        index.setNumberOfColumns("2");
        index.setPageNumberListSeparator("|");
        index.setPageNumberSeparator("|");
        index.setPageRangeSeparator("/");
        index.setUseYomi(true);
        index.setRunSubentriesOnSameLine(false);
        index.setSequenceName("Chapter");
        index.setSequenceSeparator(":");
        Assert.assertTrue(index.hasPageNumberSeparator());
        Assert.assertTrue(index.hasSequenceName());

        // Our index will take up page 1
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Use a document builder to insert an index entry
        // Index entries are not added to the index manually, it will find them on its own
        FieldXE indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 1");
        indexEntry.setEntryType("Type1");
        indexEntry.isBold(true);
        indexEntry.isItalic(true);
        Assert.assertEquals(indexEntry.hasPageRangeBookmarkName(), false);

        // We can insert a bookmark and have the index field point to it
        String subBookmarkName = "MyBookmark";
        builder.startBookmark(subBookmarkName);
        builder.writeln("Bookmark text contents.");
        builder.endBookmark(subBookmarkName);

        // Put the bookmark and index entry field on different pages
        // Our index will use the page that the bookmark is on, not that of the index entry field, as the page number
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 2");
        indexEntry.setEntryType("Type1");
        indexEntry.setPageRangeBookmarkName(subBookmarkName);
        Assert.assertEquals(indexEntry.hasPageRangeBookmarkName(), true);

        // We can use the PageNumberReplacement property to point to any page we want, even one that may not exist
        builder.insertBreak(BreakType.PAGE_BREAK);
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("Index entry 3");
        indexEntry.setEntryType("Type1");
        indexEntry.setPageNumberReplacement("999");

        // If we are using an East Asian language, we can sort entries phonetically (using Furigana) instead of alphabetically
        indexEntry = (FieldXE) builder.insertField(FieldType.FIELD_INDEX_ENTRY, true);
        indexEntry.setText("æ¼¢å­—");
        indexEntry.setEntryType("Type1");

        // The Yomi field will contain the character looked up for sorting
        indexEntry.setYomi("ã‹");

        // If we are sorting phonetically, we need to notify the index
        index.setUseYomi(true);

        // For all our entry fields, we set the entry type to "Type1"
        // Our field index will not list those entries unless we set its entry type to that of the entries
        index.setEntryType("Type1");

        builder.endBookmark(mainBookmarkName);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.XE.docx");
        //ExEnd
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
        FieldBarcode fieldBarcode = (FieldBarcode) builder.insertField(FieldType.FIELD_BARCODE, true);
        fieldBarcode.setFacingIdentificationMark("C");
        fieldBarcode.setPostalAddress("96801");
        fieldBarcode.isUSPostalAddress(true);

        builder.writeln();

        // Reference a US postal code from a bookmark
        fieldBarcode = (FieldBarcode) builder.insertField(FieldType.FIELD_BARCODE, true);
        fieldBarcode.setPostalAddress("BarcodeBookmark");
        fieldBarcode.isBookmark(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.USAddressBarcode.docx");
        //ExEnd
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
        doc.save(getArtifactsDir() + "Field.DisplayBarcode.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:FieldMergeBarcode
    //ExFor:FieldMergeBarcode.AddStartStopChar
    //ExFor:FieldMergeBarcode.BackgroundColor
    //ExFor:FieldMergeBarcode.BarcodeType
    //ExFor:FieldMergeBarcode.BarcodeValue
    //ExFor:FieldMergeBarcode.CaseCodeStyle
    //ExFor:FieldMergeBarcode.DisplayText
    //ExFor:FieldMergeBarcode.ErrorCorrectionLevel
    //ExFor:FieldMergeBarcode.FixCheckDigit
    //ExFor:FieldMergeBarcode.ForegroundColor
    //ExFor:FieldMergeBarcode.PosCodeStyle
    //ExFor:FieldMergeBarcode.ScalingFactor
    //ExFor:FieldMergeBarcode.SymbolHeight
    //ExFor:FieldMergeBarcode.SymbolRotation
    //ExSummary:Shows how to use MERGEBARCODE fields to integrate barcodes into mail merge operations.
    @Test(enabled = false, description = "Bug!!!") //ExSkip
    public void fieldMergeBarcodeQR() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QR code
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("QR");

        // In a DISPLAYBARCODE field, the BarcodeValue attribute decides what value the barcode will display
        // However in our MERGEBARCODE fields, it has the same function as the FieldName attribute of a MERGEFIELD
        field.setBarcodeValue("MyQRCode");
        field.setBackgroundColor("0xF8BD69");
        field.setForegroundColor("0xB5413B");
        field.setErrorCorrectionLevel("3");
        field.setScalingFactor("250");
        field.setSymbolHeight("1000");
        field.setSymbolRotation("0");

        Assert.assertEquals(field.getFieldCode(), " MERGEBARCODE  MyQRCode QR \\b 0xF8BD69 \\f 0xB5413B \\q 3 \\s 250 \\h 1000 \\r 0");
        builder.writeln();

        // Create a data source for our mail merge
        // This source is a data table, whose column names correspond to the FieldName attributes of MERGEFIELD fields
        // as well as BarcodeValue attributes of DISPLAYBARCODE fields
        DataTable table = createTable("Barcodes", new String[]{"MyQRCode"},
                new String[][]{{"ABC123"}, {"DEF456"}});

        // During the mail merge, all our MERGEBARCODE fields will be converted into DISPLAYBARCODE fields,
        // with values from the data table rows deposited into corresponding BarcodeValue attributes
        doc.getMailMerge().execute(table);

        Assert.assertEquals(doc.getRange().getFields().get(0).getType(), FieldType.FIELD_DISPLAY_BARCODE);
        Assert.assertEquals(doc.getRange().getFields().get(1).getType(), FieldType.FIELD_DISPLAY_BARCODE);

        Assert.assertEquals(doc.getRange().getFields().get(0).getFieldCode(), "DISPLAYBARCODE \"ABC123\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B");
        Assert.assertEquals(doc.getRange().getFields().get(1).getFieldCode(), "DISPLAYBARCODE \"DEF456\" QR \\q 3 \\s 250 \\h 1000 \\r 0 \\b 0xF8BD69 \\f 0xB5413B");

        doc.save(getArtifactsDir() + "Field.MergeBarcode_QR.docx");
    }

    @Test(enabled = false, description = "Bug!!!") //ExSkip
    public void fieldMergeBarcodeEAN13() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a EAN13 barcode
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("EAN13");
        field.setBarcodeValue("MyEAN13Barcode");
        field.setDisplayText(true);
        field.setPosCodeStyle("CASE");
        field.setFixCheckDigit(true);

        Assert.assertEquals(field.getFieldCode(), " MERGEBARCODE  MyEAN13Barcode EAN13 \\t \\p CASE \\x");
        builder.writeln();

        DataTable table = createTable("Barcodes", new String[]{"MyEAN13Barcode"},
                new String[][]{{"501234567890"}, {"123456789012"}});

        doc.getMailMerge().execute(table);

        Assert.assertEquals(doc.getRange().getFields().get(0).getType(), FieldType.FIELD_DISPLAY_BARCODE);
        Assert.assertEquals(doc.getRange().getFields().get(1).getType(), FieldType.FIELD_DISPLAY_BARCODE);

        Assert.assertEquals(doc.getRange().getFields().get(0).getFieldCode(), "DISPLAYBARCODE \"501234567890\" EAN13 \\t \\p CASE \\x");
        Assert.assertEquals(doc.getRange().getFields().get(1).getFieldCode(), "DISPLAYBARCODE \"123456789012\" EAN13 \\t \\p CASE \\x");

        doc.save(getArtifactsDir() + "Field.MergeBarcode_EAN13.docx");
    }

    @Test(enabled = false, description = "Bug!!!") //ExSkip
    public void fieldMergeBarcodeCODE39() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a CODE39 barcode
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("CODE39");
        field.setBarcodeValue("MyCODE39Barcode");
        field.setAddStartStopChar(true);

        Assert.assertEquals(field.getFieldCode(), " MERGEBARCODE  MyCODE39Barcode CODE39 \\d");
        builder.writeln();

        DataTable table = createTable("Barcodes", new String[]{"MyCODE39Barcode"},
                new String[][]{{"12345ABCDE"}, {"67890FGHIJ"}});

        doc.getMailMerge().execute(table);

        Assert.assertEquals(doc.getRange().getFields().get(0).getType(), FieldType.FIELD_DISPLAY_BARCODE);
        Assert.assertEquals(doc.getRange().getFields().get(1).getType(), FieldType.FIELD_DISPLAY_BARCODE);

        Assert.assertEquals(doc.getRange().getFields().get(0).getFieldCode(), "DISPLAYBARCODE \"12345ABCDE\" CODE39 \\d");
        Assert.assertEquals(doc.getRange().getFields().get(1).getFieldCode(), "DISPLAYBARCODE \"67890FGHIJ\" CODE39 \\d");

        doc.save(getArtifactsDir() + "Field.MergeBarcode_CODE39.docx");
    }

    @Test(enabled = false, description = "Bug!!!") //ExSkip
    public void fieldMergeBarcodeITF14() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a ITF14 barcode
        FieldMergeBarcode field = (FieldMergeBarcode) builder.insertField(FieldType.FIELD_MERGE_BARCODE, true);
        field.setBarcodeType("ITF14");
        field.setBarcodeValue("MyITF14Barcode");
        field.setCaseCodeStyle("STD");

        Assert.assertEquals(field.getFieldCode(), " MERGEBARCODE  MyITF14Barcode ITF14 \\c STD");

        DataTable table = createTable("Barcodes", new String[]{"MyITF14Barcode"},
                new String[][]{{"09312345678907"}, {"1234567891234"}});

        doc.getMailMerge().execute(table);

        Assert.assertEquals(doc.getRange().getFields().get(0).getType(), FieldType.FIELD_DISPLAY_BARCODE);
        Assert.assertEquals(doc.getRange().getFields().get(1).getType(), FieldType.FIELD_DISPLAY_BARCODE);

        Assert.assertEquals(doc.getRange().getFields().get(0).getFieldCode().toString(), "DISPLAYBARCODE \"09312345678907\" ITF14 \\c STD");
        Assert.assertEquals(doc.getRange().getFields().get(1).getFieldCode().toString(), "DISPLAYBARCODE \"1234567891234\" ITF14 \\c STD");

        doc.save(getArtifactsDir() + "Field.MergeBarcode_ITF14.docx");
    }

    /// <summary>
    /// Creates a DataTable named by dataTableName, adds a column for every element in columnNames
    /// and fills rows with data from dataSet
    /// </summary>
    @Test(enabled = false)
    public DataTable createTable(final String dataTableName, final String[] columnNames, final Object[][] dataSet) {
        if (!new String(dataTableName).equals("") || columnNames.length != 0) {
            DataTable table = new DataTable(dataTableName);

            for (String columnName : columnNames) {
                table.getColumns().add(columnName);
            }

            for (Object data : dataSet) {
                table.getRows().add(data);
            }

            return table;
        }

        throw new IllegalArgumentException("DataTable name and Column name must be declared.");
    }
    //ExEnd

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
    public void fieldLinkedObjectsAsText(final int insertLinkedObjectAs) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert fields containing text from another document and present them as text (see InsertLinkedObjectAs enum).
        builder.writeln("FieldLink:\n");
        insertFieldLink(builder, insertLinkedObjectAs, "Word.Document.8", getMyDir() + "Document.doc",
                null, true);

        builder.writeln("FieldDde:\n");
        insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Document.Spreadsheet.xlsx",
                "Sheet1!R1C1", true, true);

        builder.writeln("FieldDdeAuto:\n");
        insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Document.Spreadsheet.xlsx",
                "Sheet1!R1C1", true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LinkedObjectsAsText.docx");
    }

    //JAVA-added data provider for test method
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
    public void fieldLinkedObjectsAsImage(final int insertLinkedObjectAs) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert one cell from a spreadsheet as an image (see InsertLinkedObjectAs enum)
        builder.writeln("FieldLink:\n");
        insertFieldLink(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "MySpreadsheet.xlsx",
                "Sheet1!R2C2", true);

        builder.writeln("FieldDde:\n");
        insertFieldDde(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Document.Spreadsheet.xlsx",
                "Sheet1!R1C1", true, true);

        builder.writeln("FieldDdeAuto:\n");
        insertFieldDdeAuto(builder, insertLinkedObjectAs, "Excel.Sheet", getMyDir() + "Document.Spreadsheet.xlsx",
                "Sheet1!R1C1", true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.LinkedObjectsAsImage.docx");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "fieldLinkedObjectsAsImageDataProvider")
    public static Object[][] fieldLinkedObjectsAsImageDataProvider() {
        return new Object[][]
                {
                        {InsertLinkedObjectAs.PICTURE},
                        {InsertLinkedObjectAs.BITMAP},
                };
    }

    /// <summary>
    /// Use a document builder to insert a LINK field and set its properties according to parameters
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
    /// Use a document builder to insert a DDE field and set its properties according to parameters
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
    /// Use a document builder to insert a DDEAUTO field and set its properties according to parameters
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
    public void fieldOptionsCurrentUser() throws Exception {
        //ExStart
        //ExFor:FieldOptions.CurrentUser
        //ExFor:UserInformation
        //ExFor:UserInformation.Name
        //ExFor:UserInformation.Initials
        //ExFor:UserInformation.Address
        //ExFor:UserInformation.DefaultUser
        //ExSummary:Shows how to set user details and display them with fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set user information
        UserInformation userInformation = new UserInformation();
        userInformation.setName("John Doe");
        userInformation.setInitials("J. D.");
        userInformation.setAddress("123 Main Street");
        doc.getFieldOptions().setCurrentUser(userInformation);

        // Insert fields that reference our user information
        Assert.assertEquals(userInformation.getName(), builder.insertField(" USERNAME ").getResult());
        Assert.assertEquals(userInformation.getInitials(), builder.insertField(" USERINITIALS ").getResult());
        Assert.assertEquals(userInformation.getAddress(), builder.insertField(" USERADDRESS ").getResult());

        // The field options object also has a static default user value that fields from many documents can refer to
        UserInformation.getDefaultUser().setName("Default User");
        UserInformation.getDefaultUser().setInitials("D. U.");
        UserInformation.getDefaultUser().setAddress("One Microsoft Way");
        doc.getFieldOptions().setCurrentUser(UserInformation.getDefaultUser());

        Assert.assertEquals(builder.insertField(" USERNAME ").getResult(), "Default User");
        Assert.assertEquals(builder.insertField(" USERINITIALS ").getResult(), "D. U.");
        Assert.assertEquals(builder.insertField(" USERADDRESS ").getResult(), "One Microsoft Way");
        //ExEnd
    }

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
        Assert.assertEquals(fieldUserAddress.getResult(), userInformation.getAddress());

        Assert.assertEquals(fieldUserAddress.getFieldCode(), " USERADDRESS ");
        Assert.assertEquals(fieldUserAddress.getResult(), "123 Main Street");

        // We can set this attribute to get our field to display a different value
        fieldUserAddress.setUserAddress("456 North Road");
        fieldUserAddress.update();

        Assert.assertEquals(fieldUserAddress.getFieldCode(), " USERADDRESS  \"456 North Road\"");
        Assert.assertEquals(fieldUserAddress.getResult(), "456 North Road");

        // This does not change the value in the user information object
        Assert.assertEquals(doc.getFieldOptions().getCurrentUser().getAddress(), "123 Main Street");
        //ExEnd
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
        Assert.assertEquals(fieldUserInitials.getResult(), userInformation.getInitials());

        Assert.assertEquals(fieldUserInitials.getFieldCode(), " USERINITIALS ");
        Assert.assertEquals(fieldUserInitials.getResult(), "J. D.");

        // We can set this attribute to get our field to display a different value
        fieldUserInitials.setUserInitials("J. C.");
        fieldUserInitials.update();

        Assert.assertEquals(fieldUserInitials.getFieldCode(), " USERINITIALS  \"J. C.\"");
        Assert.assertEquals(fieldUserInitials.getResult(), "J. C.");

        // This does not change the value in the user information object
        Assert.assertEquals(doc.getFieldOptions().getCurrentUser().getInitials(), "J. D.");
        //ExEnd
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
        Assert.assertEquals(fieldUserName.getResult(), userInformation.getName());

        Assert.assertEquals(fieldUserName.getFieldCode(), " USERNAME ");
        Assert.assertEquals(fieldUserName.getResult(), "John Doe");

        // We can set this attribute to get our field to display a different value
        fieldUserName.setUserName("Jane Doe");
        fieldUserName.update();

        Assert.assertEquals(fieldUserName.getFieldCode(), " USERNAME  \"Jane Doe\"");
        Assert.assertEquals(fieldUserName.getResult(), "Jane Doe");

        // This does not change the value in the user information object
        Assert.assertEquals(doc.getFieldOptions().getCurrentUser().getName(), "John Doe");
        //ExEnd
    }

    @Test
    public void fieldOptionsFileName() throws Exception {
        //ExStart
        //ExFor:FieldOptions.FileName
        //ExFor:FieldFileName
        //ExFor:FieldFileName.IncludeFullPath
        //ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToDocumentEnd();
        builder.writeln();

        // This FILENAME field will display the file name of the document we opened
        FieldFileName field = (FieldFileName) builder.insertField(FieldType.FIELD_FILE_NAME, true);
        field.update();

        Assert.assertEquals(field.getFieldCode(), " FILENAME ");
        Assert.assertEquals(field.getResult(), "Document.docx");

        builder.writeln();

        // By default, the FILENAME field does not show the full path, and we can change this
        field = (FieldFileName) builder.insertField(FieldType.FIELD_FILE_NAME, true);
        field.setIncludeFullPath(true);

        // We can override the values displayed by our FILENAME fields by setting this attribute
        Assert.assertNull(doc.getFieldOptions().getFileName());
        doc.getFieldOptions().setFileName("Field.FileName.docx");
        field.update();

        Assert.assertEquals(field.getFieldCode(), " FILENAME  \\p");
        Assert.assertEquals(field.getResult(), "Field.FileName.docx");

        doc.updateFields();
        doc.save(getArtifactsDir() + "" + doc.getFieldOptions().getFileName());
        //ExEnd
    }

    @Test
    public void fieldOptionsBidi() throws Exception {
        //ExStart
        //ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
        //ExSummary:Shows how to use FieldOptions to ensure that bi-directional text is properly supported during the field update.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure that any field operation involving right-to-left text is performed correctly
        doc.getFieldOptions().isBidiTextSupportedOnUpdate(true);

        // Use a document builder to insert a field which contains right-to-left text
        FormField comboBox = builder.insertComboBox("MyComboBox", new String[]{"×¢Ö¶×©Ö°×‚×¨Ö´×™×", "×©Ö°××œ×•Ö¹×©Ö´××™×", "×Ö·×¨Ö°×‘Ö¸Ö¼×¢Ö´×™×", "×—Ö²×žÖ´×©Ö´Ö¼××™×", "×©Ö´××©Ö´Ö¼××™×"}, 0);
        comboBox.setCalculateOnExit(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FieldOptionsBidi.docx");
        //ExEnd
    }

    @Test
    public void fieldOptionsLegacyNumberFormat() throws Exception {
        //ExStart
        //ExFor:FieldOptions.LegacyNumberFormat
        //ExSummary:Shows how use FieldOptions to change the number format.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Field field = builder.insertField("= 2 + 3 \\# $##");

        Assert.assertEquals(field.getResult(), "$ 5");

        doc.getFieldOptions().setLegacyNumberFormat(true);
        field.update();

        Assert.assertEquals(field.getResult(), "$5");
        //ExEnd
    }

    @Test
    public void fieldOptionsToaCategories() throws Exception {
        //ExStart
        //ExFor:FieldOptions.ToaCategories
        //ExFor:ToaCategories
        //ExFor:ToaCategories.Item(Int32)
        //ExFor:ToaCategories.DefaultCategories
        //ExSummary:Shows how to specify a table of authorities categories for a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // There are default category values we can use, or we can make our own like this
        ToaCategories toaCategories = new ToaCategories();
        doc.getFieldOptions().setToaCategories(toaCategories);

        toaCategories.set(1, "My Category 1"); // Replaces default value "Cases"
        toaCategories.set(2, "My Category 2"); // Replaces default value "Statutes"

        // Even if we changed the categories in the FieldOptions object, the default categories are still available here
        Assert.assertEquals(ToaCategories.getDefaultCategories().get(1), "Cases");
        Assert.assertEquals(ToaCategories.getDefaultCategories().get(2), "Statutes");

        // Insert 2 tables of authorities, one per category
        builder.insertField("TOA \\c 1 \\h", null);
        builder.insertField("TOA \\c 2 \\h", null);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Insert table of authorities entries across 2 categories
        builder.insertField("TA \\c 2 \\l \"entry 1\"");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertField("TA \\c 1 \\l \"entry 2\"");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertField("TA \\c 2 \\l \"entry 3\"");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.TableOfAuthorities.Categories.docx");
        //ExEnd
    }

    @Test
    public void fieldOptionsUseInvariantCultureNumberFormat() throws Exception {
        //ExStart
        //ExFor:FieldOptions.UseInvariantCultureNumberFormat
        //ExSummary:Shows how to format numbers according to the invariant culture.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Locale defaultLocale = Locale.getDefault();
        Locale.setDefault(new Locale("de-DE"));

        Field field = builder.insertField(" = 1234567,89 \\# $#,###,###.##");
        field.update();

        // The combination of field, number format and thread culture can sometimes produce an unsuitable result
        Assert.assertFalse(doc.getFieldOptions().getUseInvariantCultureNumberFormat());
        Assert.assertEquals(field.getResult(), "$123,456,789.  ");

        // We can set this attribute to avoid changing the whole thread culture just for numeric formats
        doc.getFieldOptions().setUseInvariantCultureNumberFormat(true);
        field.update();
        Assert.assertEquals(field.getResult(), "$123,456,789.  ");

        Locale.setDefault(defaultLocale);
        //ExEnd
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
        FieldStyleRef fieldStyleRef = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        fieldStyleRef.setStyleName("List Paragraph");

        // Place a STYLEREF field in the footer and have it display the last text
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        fieldStyleRef = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        fieldStyleRef.setStyleName("List Paragraph");
        fieldStyleRef.setSearchFromBottom(true);

        builder.moveToDocumentEnd();

        // We can also use STYLEREF fields to reference the list numbers of lists
        builder.write("\nParagraph number: ");
        fieldStyleRef = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        fieldStyleRef.setStyleName("Quote");
        fieldStyleRef.setInsertParagraphNumber(true);

        builder.write("\nParagraph number, relative context: ");
        fieldStyleRef = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        fieldStyleRef.setStyleName("Quote");
        fieldStyleRef.setInsertParagraphNumberInRelativeContext(true);

        builder.write("\nParagraph number, full context: ");
        fieldStyleRef = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        fieldStyleRef.setStyleName("Quote");
        fieldStyleRef.setInsertParagraphNumberInFullContext(true);

        builder.write("\nParagraph number, full context, non-delimiter chars suppressed: ");
        fieldStyleRef = (FieldStyleRef) builder.insertField(FieldType.FIELD_STYLE_REF, true);
        fieldStyleRef.setStyleName("Quote");
        fieldStyleRef.setInsertParagraphNumberInFullContext(true);
        fieldStyleRef.setSuppressNonDelimiters(true);

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FieldStyleRef.docx");
        //ExEnd
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
        FieldDate fieldDate = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);

        // Set the field's date to the current date of the Islamic Lunar Calendar
        fieldDate.setUseLunarCalendar(true);
        Assert.assertEquals(fieldDate.getFieldCode(), " DATE  \\h");
        builder.writeln();

        // Insert a date field with the current date of the Umm al-Qura calendar
        fieldDate = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        fieldDate.setUseUmAlQuraCalendar(true);
        Assert.assertEquals(fieldDate.getFieldCode(), " DATE  \\u");
        builder.writeln();

        // Insert a date field with the current date of the Indian national calendar
        fieldDate = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        fieldDate.setUseSakaEraCalendar(true);
        Assert.assertEquals(fieldDate.getFieldCode(), " DATE  \\s");
        builder.writeln();

        // Insert a date field with the current date of the calendar used in the (Insert > Date and Time) dialog box
        fieldDate = (FieldDate) builder.insertField(FieldType.FIELD_DATE, true);
        fieldDate.setUseLastFormat(true);
        Assert.assertEquals(fieldDate.getFieldCode(), " DATE  \\l");
        builder.writeln();

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.Date.docx");
        //ExEnd
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
        FieldCreateDate fieldCreateDate = (FieldCreateDate) builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        fieldCreateDate.setUseLunarCalendar(true);
        Assert.assertEquals(fieldCreateDate.getFieldCode(), " CREATEDATE  \\h");
        builder.writeln();

        // Display the date using the Umm al-Qura Calendar
        builder.write("According to the Umm al-Qura Calendar - ");
        fieldCreateDate = (FieldCreateDate) builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        fieldCreateDate.setUseUmAlQuraCalendar(true);
        Assert.assertEquals(fieldCreateDate.getFieldCode(), " CREATEDATE  \\u");
        builder.writeln();

        // Display the date using the Indian National Calendar
        builder.write("According to the Indian National Calendar - ");
        fieldCreateDate = (FieldCreateDate) builder.insertField(FieldType.FIELD_CREATE_DATE, true);
        fieldCreateDate.setUseSakaEraCalendar(true);
        Assert.assertEquals(fieldCreateDate.getFieldCode(), " CREATEDATE  \\s");
        builder.writeln();

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.CreateDate.docx");
        //ExEnd
    }

    @Test
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
        FieldSaveDate fieldSaveDate = (FieldSaveDate) builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        fieldSaveDate.setUseLunarCalendar(true);
        Assert.assertEquals(fieldSaveDate.getFieldCode(), " SAVEDATE  \\h");
        builder.writeln();

        // Display the date using the Umm al-Qura Calendar
        builder.write("According to the Umm al-Qura calendar - ");
        fieldSaveDate = (FieldSaveDate) builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        fieldSaveDate.setUseUmAlQuraCalendar(true);
        Assert.assertEquals(fieldSaveDate.getFieldCode(), " SAVEDATE  \\u");
        builder.writeln();

        // Display the date using the Indian National Calendar
        builder.write("According to the Indian National calendar - ");
        fieldSaveDate = (FieldSaveDate) builder.insertField(FieldType.FIELD_SAVE_DATE, true);
        fieldSaveDate.setUseSakaEraCalendar(true);
        Assert.assertEquals(fieldSaveDate.getFieldCode(), " SAVEDATE  \\s");
        builder.writeln();

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.SaveDate.docx");
        //ExEnd
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
        doc.save(getArtifactsDir() + "Field.FieldBuilder.docx");
        //ExEnd
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

        Assert.assertEquals(fieldDocVariable.getFieldCode(), " DOCVARIABLE  \"My Variable\"");
        Assert.assertEquals(fieldDocVariable.getResult(), "My variable's value");
        //ExEnd
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
        Assert.assertEquals(doc.getBuiltInDocumentProperties().getSubject(), "My new subject");
        //ExEnd
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

        Assert.assertEquals(field.getResult(), "My comment.");

        // We can override the comment from the document's built in properties and display any text we put here instead
        field.setText("My overriding comment.");
        field.update();

        Assert.assertEquals(field.getResult(), "My overriding comment.");

        doc.save(getArtifactsDir() + "Field.Comments.docx");
        //ExEnd
    }

    @Test
    public void fieldFileSize() throws Exception {
        //ExStart
        //ExFor:FieldFileSize
        //ExFor:FieldFileSize.IsInKilobytes
        //ExFor:FieldFileSize.IsInMegabytes
        //ExSummary:Shows how to display the file size of a document with a FILESIZE field.
        Document doc = new Document(getMyDir() + "Document.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToDocumentEnd();

        // By default, file size is displayed in bytes
        FieldFileSize field = (FieldFileSize) builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.update();
        Assert.assertEquals(field.getResult(), "23040");

        // Set the field to display size in kilobytes
        field = (FieldFileSize) builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.isInKilobytes(true);
        field.update();
        Assert.assertEquals(field.getResult(), "23");

        // Set the field to display size in megabytes
        field = (FieldFileSize) builder.insertField(FieldType.FIELD_FILE_SIZE, true);
        field.isInMegabytes(true);
        field.update();
        Assert.assertEquals(field.getResult(), "0");
        //ExEnd
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
        doc.save(getArtifactsDir() + "Field.GoToButton.docx");
        //ExEnd
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
        doc.save(getArtifactsDir() + "Field.FillIn.docx");
    }

    /// <summary>
    /// IFieldUserPromptRespondent implementation that appends a line to the default response of an FILLIN field during a mail merge
    /// </summary>
    private static class PromptRespondent implements IFieldUserPromptRespondent {
        public String respond(final String promptText, final String defaultResponse) {
            return "Response from PromptRespondent. " + defaultResponse;
        }
    }
    //ExEnd

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

        doc.save(getArtifactsDir() + "Field.Info.docx");
        //ExEnd
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
        Document doc = new Document(getMyDir() + "Document.HasMacro.docm");
        DocumentBuilder builder = new DocumentBuilder(doc);
        Assert.assertTrue(doc.hasMacros());

        // Insert a MACROBUTTON field and reference by name a macro that exists within the input document
        FieldMacroButton field = (FieldMacroButton) builder.insertField(FieldType.FIELD_MACRO_BUTTON, true);
        field.setMacroName("MyMacro");
        field.setDisplayText("Double click to run macro: " + field.getMacroName());

        Assert.assertEquals(field.getFieldCode(), " MACROBUTTON  MyMacro Double click to run macro: MyMacro");

        builder.insertParagraph();

        // Reference "ViewZoom200", a macro that was shipped with Microsoft Word, found under "Word commands"
        // If our document has a macro of the same name as one from another source, the field will select ours to run
        field = (FieldMacroButton) builder.insertField(FieldType.FIELD_MACRO_BUTTON, true);
        field.setMacroName("ViewZoom200");
        field.setDisplayText("Run " + field.getMacroName());

        Assert.assertEquals(field.getFieldCode(), " MACROBUTTON  ViewZoom200 Run ViewZoom200");

        // Save the document as a macro-enabled document type
        doc.save(getArtifactsDir() + "Field.MacroButton.docm");
        //ExEnd
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

        doc.save(getArtifactsDir() + "Field.Keywords.docx");
        //ExEnd
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
        Document doc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");
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
        doc.save(getArtifactsDir() + "Field.Num.docx");
        //ExEnd
    }

    @Test
    public void fieldPrint() throws Exception {
        //ExStart
        //ExFor:FieldPrint
        //ExFor:FieldPrint.PostScriptGroup
        //ExFor:FieldPrint.PrinterInstructions
        //ExFor:FieldPrintDate
        //ExFor:FieldPrintDate.UseLunarCalendar
        //ExFor:FieldPrintDate.UseSakaEraCalendar
        //ExFor:FieldPrintDate.UseUmAlQuraCalendar
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

        Assert.assertEquals(field.getFieldCode(), " PRINT  erasepage \\p para");

        builder.insertParagraph();

        // PRINTDATE field will display "0/0/0000" by default
        // When a document is printed by a printer or printed as a PDF (but not exported as PDF),
        // these fields will display the date/time of the printing operation, in various calendars
        FieldPrintDate fieldPrintDate = (FieldPrintDate) builder.insertField(FieldType.FIELD_PRINT_DATE, true);
        fieldPrintDate.setUseLunarCalendar(true);
        builder.writeln();

        Assert.assertEquals(fieldPrintDate.getFieldCode(), " PRINTDATE  \\h");

        fieldPrintDate = (FieldPrintDate) builder.insertField(FieldType.FIELD_PRINT_DATE, true);
        fieldPrintDate.setUseSakaEraCalendar(true);
        builder.writeln();

        Assert.assertEquals(fieldPrintDate.getFieldCode(), " PRINTDATE  \\s");

        fieldPrintDate = (FieldPrintDate) builder.insertField(FieldType.FIELD_PRINT_DATE, true);
        fieldPrintDate.setUseUmAlQuraCalendar(true);
        builder.writeln();

        Assert.assertEquals(fieldPrintDate.getFieldCode(), " PRINTDATE  \\u");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.Print.docx");
        //ExEnd
    }

    @Test
    public void fieldQuote() throws Exception {
        //ExStart
        //ExFor:FieldQuote
        //ExFor:FieldQuote.Text
        //ExSummary:Shows to use the QUOTE field.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a QUOTE field, which will display content from the Text attribute
        FieldQuote field = (FieldQuote) builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("\"Quoted text\"");

        Assert.assertEquals(field.getFieldCode(), " QUOTE  \"\\\"Quoted text\\\"\"");

        builder.insertParagraph();

        // Insert a QUOTE field with a nested DATE field
        // DATE fields normally update their value to the current date every time the document is opened
        // Nesting the DATE field inside the QUOTE field like this will freeze its value to the date when we created the document
        builder.write("Document creation date: ");
        field = (FieldQuote) builder.insertField(FieldType.FIELD_QUOTE, true);
        builder.moveTo(field.getSeparator());
        builder.insertField(FieldType.FIELD_DATE, true);

        LocalDateTime actualDate = LocalDateTime.now();
        String actualDateFormated = DateTimeFormatter.ofPattern("M/d/yyyy").format(actualDate);

        Assert.assertEquals(field.getFieldCode(), " QUOTE \u0013 DATE \u0014" + actualDateFormated + "\u0015");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.Quote.docx");
        //ExEnd
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

        Assert.assertEquals(fieldNext.getFieldCode(), " NEXT ");
        Assert.assertEquals(fieldNextIf.getFieldCode(), " NEXTIF  5 = \"2 + 3\"");

        doc.save(getArtifactsDir() + "Field.Next.docx");
    }

    /// <summary>
    /// Uses a document builder to insert merge fields for a data table that has "Courtesy Title", "First Name" and "Last Name" columns
    /// </summary>
    @Test(enabled = false)
    public void insertMergeFields(final DocumentBuilder builder, final String firstFieldTextBefore) throws Exception {
        insertMergeField(builder, "Courtesy Title", firstFieldTextBefore, " ");
        insertMergeField(builder, "First Name", null, " ");
        insertMergeField(builder, "Last Name", null, null);
        builder.insertParagraph();
    }

    /// <summary>
    /// Uses a document builder to insert a merge field
    /// </summary>
    @Test(enabled = false)
    public void insertMergeField(final DocumentBuilder builder, final String fieldName, final String textBefore, final String textAfter) throws Exception {
        FieldMergeField field = (FieldMergeField) builder.insertField(FieldType.FIELD_MERGE_FIELD, true);
        field.setFieldName(fieldName);
        field.setTextBefore(textBefore);
        field.setTextAfter(textAfter);
    }
    //ExEnd

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
        Assert.assertEquals(" NOTEREF  MyBookmark2 \\h \\f \\p",
                insertFieldNoteRef(builder, "MyBookmark2", true, true, true, "Bookmark2, with footnote number ").getFieldCode());

        builder.insertBreak(BreakType.PAGE_BREAK);
        insertBookmarkWithFootnote(builder, "MyBookmark2", "Contents of MyBookmark2", "Footnote from MyBookmark2");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.NoteRef.docx");
    }

    /// <summary>
    /// Uses a document builder to insert a NOTEREF field and sets its attributes
    /// </summary>
    private FieldNoteRef insertFieldNoteRef(final DocumentBuilder builder, final String bookmarkName, final boolean insertHyperlink,
                                            final boolean insertRelativePosition, final boolean insertReferenceMark,
                                            final String textBefore) throws Exception {
        builder.write(textBefore);

        FieldNoteRef field = (FieldNoteRef) builder.insertField(FieldType.FIELD_NOTE_REF, true);
        field.setBookmarkName(bookmarkName);
        field.setInsertHyperlink(insertHyperlink);
        field.setInsertReferenceMark(insertReferenceMark);
        field.setInsertRelativePosition(insertRelativePosition);
        builder.writeln();

        return field;
    }

    /// <summary>
    /// Uses a document builder to insert a named bookmark with a footnote at the end
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

    @Test
    public void footnoteRef() throws Exception {
        //ExStart
        //ExFor:FieldFootnoteRef
        //ExSummary:Shows how to cross-reference footnotes with the FOOTNOTEREF field
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
        //Field field = builder.insertField(" ftnref ");
        FieldFootnoteRef field = (FieldFootnoteRef) builder.insertField(FieldType.FIELD_FOOTNOTE_REF, true);

        // Get this field to reference a bookmark
        // The bookmark that we chose contains a footnote marker belonging to the footnote we inserted, which will be displayed by the field, just by itself
        builder.moveTo(field.getSeparator());
        builder.write("CrossRefBookmark");

        Assert.assertEquals(field.getFieldCode(), " FOOTNOTEREF CrossRefBookmark");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.FootnoteRef.docx");
        //ExEnd
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
        Assert.assertEquals(insertFieldPageRef(builder, "MyBookmark3", true, false, "Hyperlink to Bookmark3, on page: ").getFieldCode(),
                " PAGEREF  MyBookmark3 \\h");

        // Setting the \p flag makes the field display the relative position of the bookmark to the field instead of a page number
        // Bookmark1 is on the same page and above this field, so the result will be "above" on update
        Assert.assertEquals(insertFieldPageRef(builder, "MyBookmark1", true, true, "Bookmark1 is ").getFieldCode(),
                " PAGEREF  MyBookmark1 \\h \\p");

        // Bookmark2 will be on the same page and below this field, so the field will display "below"
        Assert.assertEquals(insertFieldPageRef(builder, "MyBookmark2", true, true, "Bookmark2 is ").getFieldCode(),
                " PAGEREF  MyBookmark2 \\h \\p");

        // Bookmark3 will be on a different page, so the field will display "on page 2"
        Assert.assertEquals(insertFieldPageRef(builder, "MyBookmark3", true, true, "Bookmark3 is ").getFieldCode(),
                " PAGEREF  MyBookmark3 \\h \\p");

        insertAndNameBookmark(builder, "MyBookmark2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        insertAndNameBookmark(builder, "MyBookmark3");

        doc.updateFields();
        doc.save(getArtifactsDir() + "Field.PageRef.docx");
    }

    /// <summary>
    /// Uses a document builder to insert a PAGEREF field and sets its attributes
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
    /// Uses a document builder to insert a named bookmark
    /// </summary>
    private void insertAndNameBookmark(final DocumentBuilder builder, final String bookmarkName) {
        builder.startBookmark(bookmarkName);
        builder.writeln(MessageFormat.format("Contents of bookmark \"{0}\".", bookmarkName));
        builder.endBookmark(bookmarkName);
    }
    //ExEnd

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
        doc.save(getArtifactsDir() + "Field.Ref.docx");
    }

    /// <summary>
    /// Get the document builder to insert a REF field, reference a bookmark with it, and add text before and after
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
        doc.save(getArtifactsDir() + "Field.RefDoc.docx");
        //ExEnd
    }

    @Test
    public void skipIf() throws Exception {
        //ExStart
        //ExFor:FieldSkipIf
        //ExFor:FieldSkipIf.ComparisonOperator
        //ExFor:FieldSkipIf.LeftExpression
        //ExFor:FieldSkipIf.RightExpression
        //ExSummary:Shows how to skip pages in a mail merge using the SKIPIF field
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

        bookmark = doc.getRange().getBookmarks().get("MyBookmark");
        Assert.assertEquals("New text", bookmark.getText());
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
        //ExSummary:Shows how to use the SYMBOL field
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
        // With a font that supports Shift-JIS, this symbol will display "ã‚", which is the large Hiragana letter "A"
        field = (FieldSymbol) builder.insertField(FieldType.FIELD_SYMBOL, true);
        field.setFontName("MS Gothic");
        field.setCharacterCode(Integer.toString(0x82A0));
        field.isShiftJis(true);

        Assert.assertEquals(field.getFieldCode(), " SYMBOL  33440 \\f \"MS Gothic\" \\j");

        builder.write("Line 3");

        doc.save(getArtifactsDir() + "Field.SYMBOL.docx");
        //ExEnd
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

        builder.writeln();

        // Set the Text attribute to display a different value
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
    }

    /// <summary>
    /// Get a builder to insert a TA field, specifying its long citation and category,
    /// then insert a page break and return the field we created
    /// </summary>
    private FieldTA insertToaEntry(final DocumentBuilder builder, final String entryCategory, final String longCitation) throws Exception {
        FieldTA field = (FieldTA) builder.insertField(FieldType.FIELD_TOA_ENTRY, false);
        field.setEntryCategory(entryCategory);
        field.setLongCitation(longCitation);

        builder.insertBreak(BreakType.PAGE_BREAK);

        return field;
    }
    //ExEnd

    @Test
    public void fieldAddin() throws Exception {
        //ExStart
        //ExFor:FieldAddIn
        //ExSummary:Shows how to process an ADDIN field.
        // Open a document that contains an ADDIN field
        Document doc = new Document(getMyDir() + "Field.Addin.docx");

        // Aspose.Words does not support inserting ADDIN fields, but they can be read
        FieldAddIn field = (FieldAddIn) doc.getRange().getFields().get(0);
        Assert.assertEquals(field.getFieldCode(), " ADDIN \"My value\" ");
        //ExEnd
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

        Assert.assertEquals(field.getFieldCode(), " EQ \\f(1,4)");

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
        insertFieldEQ(builder, "\\a \\ac \\vs1 \\co1(lim,nâ†’âˆž) \\b (\\f(n,n2 + 12) + \\f(n,n2 + 22) + ... + \\f(n,n2 + n2))");
        insertFieldEQ(builder, "\\i (,,  \\b(\\f(x,x2 + 3x + 2))) \\s \\up10(2)");
        insertFieldEQ(builder, "\\i \\in( tan x, \\s \\up2(sec x), \\b(\\r(3) )\\s \\up4(t) \\s \\up7(2)  dt)");

        doc.save(getArtifactsDir() + "Field.EQ.docx");
    }

    /// <summary>
    /// Use a document builder to insert an EQ field, set its arguments and start a new paragraph
    /// </summary>
    private FieldEQ insertFieldEQ(final DocumentBuilder builder, final String args) throws Exception {
        FieldEQ field = (FieldEQ) builder.insertField(FieldType.FIELD_EQUATION, true);
        builder.moveTo(field.getSeparator());
        builder.write(args);
        builder.moveTo(field.getStart().getParentNode());

        builder.insertParagraph();
        return field;
    }
    //ExEnd

    @Test
    public void fieldForms() throws Exception {
        //ExStart
        //ExFor:FieldFormCheckBox
        //ExFor:FieldFormDropDown
        //ExFor:FieldFormText
        //ExSummary:Shows how to process FORMCHECKBOX, FORMDROPDOWN and FORMTEXT fields.
        // These fields are legacy equivalents of the FormField, and they can be read and not inserted by Aspose.Words,
        // and are inserted in Microsoft Word 2019 via the Legacy Tools menu in the Developer tab
        Document doc = new Document(getMyDir() + "Field.FieldForms.doc");

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
        DataTable table = createTable("Employees", new String[]{"Name"},
                new String[][]{{"Jane Doe"}, {"John Doe"}, {"Joe Bloggs"}});

        // Execute mail merge and save document
        doc.getMailMerge().execute(table);
        doc.save(getArtifactsDir() + "Field.MERGEREC.MERGESEQ.docx");
        //ExEnd
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
    }

    //ExStart
    //ExFor:FieldPrivate
    //ExSummary:Shows how to process PRIVATE fields.
    @Test //ExSkip
    public void fieldPrivate() throws Exception {
        // Open a Corel WordPerfect document that was converted to .docx format
        Document doc = new Document(getMyDir() + "Field.FromWpd.docx");

        // WordPerfect 5.x/6.x documents like the one we opened may contain PRIVATE fields
        // The PRIVATE field is a WordPerfect artifact that is preserved when a file is opened and saved in Microsoft Word
        // However, they have no functionality in Microsoft Word
        FieldPrivate field = (FieldPrivate) doc.getRange().getFields().get(0);
        Assert.assertEquals(field.getFieldCode(), " PRIVATE \"My value\" ");

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
    /// Visitor implementation that removes all PRIVATE fields that it comes across.
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
        Assert.assertEquals(field.getFieldCode(), " BIDIOUTLINE ");
        builder.writeln("שלום");

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
    }

    @Test
    public void legacy() throws Exception {
        //ExStart
        //ExFor:FieldEmbed
        //ExFor:FieldShape
        //ExFor:FieldShape.Text
        //ExSummary:Shows how some older Microsoft Word fields such as SHAPE and EMBED are handled.
        // Open a document that was created in Microsoft Word 2003
        Document doc = new Document(getMyDir() + "Field.Legacy.doc");

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
        Assert.assertEquals(shape.getShapeType(), ShapeType.OLE_OBJECT);
        //ExEnd
    }

    @Test
    public void fieldDisplayResult() throws Exception {
        //ExStart
        //ExFor:Field.DisplayResult
        //ExSummary:Shows how to get the text that represents the displayed field result.
        Document document = new Document(getMyDir() + "Field.FieldDisplayResult.docx");

        FieldCollection fields = document.getRange().getFields();

        Assert.assertEquals(fields.get(0).getDisplayResult(), "111");
        Assert.assertEquals(fields.get(1).getDisplayResult(), "222");
        Assert.assertEquals(fields.get(2).getDisplayResult(), "Multi\rLine\rText");
        Assert.assertEquals(fields.get(3).getDisplayResult(), "%");
        Assert.assertEquals(fields.get(4).getDisplayResult(), "Macro Button Text");
        Assert.assertEquals(fields.get(5).getDisplayResult(), "");

        // Method must be called to obtain correct value for the "FieldListNum", "FieldAutoNum",
        // "FieldAutoNumOut" and "FieldAutoNumLgl" fields
        document.updateListLabels();

        Assert.assertEquals(fields.get(5).getDisplayResult(), "1)");
        //ExEnd
    }
}
