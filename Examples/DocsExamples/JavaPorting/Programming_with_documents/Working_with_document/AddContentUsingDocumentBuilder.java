package DocsExamples.Programming_with_Documents.Working_with_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Table;
import com.aspose.words.CellVerticalAlignment;
import com.aspose.words.HeightRule;
import com.aspose.words.TextOrientation;
import com.aspose.words.AutoFitBehavior;
import com.aspose.words.Shape;
import com.aspose.words.HorizontalRuleFormat;
import com.aspose.words.HorizontalRuleAlignment;
import java.awt.Color;
import com.aspose.words.BreakType;
import com.aspose.words.TextFormFieldType;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.Font;
import com.aspose.words.Underline;
import com.aspose.words.ParagraphFormat;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.FindReplaceOptions;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.ms.System.msString;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;
import com.aspose.ms.System.msConsole;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.words.Section;
import com.aspose.words.ParagraphCollection;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Field;


class AddContentUsingDocumentBuilder extends DocsExamplesBase
{
    @Test
    public void createNewDocument() throws Exception
    {
        //ExStart:CreateNewDocument
        //GistId:1d626c7186a318d22d022dc96dd91d55
        Document doc = new Document();

        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello World!");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.CreateNewDocument.docx");
        //ExEnd:CreateNewDocument
    }

    @Test
    public void insertBookmark() throws Exception
    {
        //ExStart:InsertBookmark
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("FineBookmark");
        builder.writeln("This is just a fine bookmark.");
        builder.endBookmark("FineBookmark");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertBookmark.docx");
        //ExEnd:InsertBookmark
    }

    @Test
    public void buildTable() throws Exception
    {
        //ExStart:BuildTable
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.write("This is row 1 cell 1");

        builder.insertCell();
        builder.write("This is row 1 cell 2");

        builder.endRow();

        builder.insertCell();
        
        builder.getRowFormat().setHeight(100.0);
        builder.getRowFormat().setHeightRule(HeightRule.EXACTLY);
        builder.getCellFormat().setOrientation(TextOrientation.UPWARD);
        builder.writeln("This is row 2 cell 1");

        builder.insertCell();
        builder.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        builder.writeln("This is row 2 cell 2");

        builder.endRow();
        builder.endTable();

        table.autoFit(AutoFitBehavior.FIXED_COLUMN_WIDTHS);

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.BuildTable.docx");
        //ExEnd:BuildTable
    }

    @Test
    public void insertHorizontalRule() throws Exception
    {
        //ExStart:InsertHorizontalRule
        //GistId:ad463bf5f128fe6e6c1485df3c046a4c
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Insert a horizontal rule shape into the document.");
        builder.insertHorizontalRule();

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertHorizontalRule.docx");
        //ExEnd:InsertHorizontalRule
    }

    @Test
    public void horizontalRuleFormat() throws Exception
    {
        //ExStart:HorizontalRuleFormat
        //GistId:ad463bf5f128fe6e6c1485df3c046a4c
        DocumentBuilder builder = new DocumentBuilder();

        Shape shape = builder.insertHorizontalRule();
        
        HorizontalRuleFormat horizontalRuleFormat = shape.getHorizontalRuleFormat();
        horizontalRuleFormat.setAlignment(HorizontalRuleAlignment.CENTER);
        horizontalRuleFormat.setWidthPercent(70.0);
        horizontalRuleFormat.setHeight(3.0);
        horizontalRuleFormat.setColor(Color.BLUE);
        horizontalRuleFormat.setNoShade(true);

        builder.getDocument().save(getArtifactsDir() + "AddContentUsingDocumentBuilder.HorizontalRuleFormat.docx");
        //ExEnd:HorizontalRuleFormat
    }

    @Test
    public void insertBreak() throws Exception
    {
        //ExStart:InsertBreak
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("This is page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.writeln("This is page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.writeln("This is page 3.");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertBreak.docx");
        //ExEnd:InsertBreak
    }

    @Test
    public void insertTextInputFormField() throws Exception
    {
        //ExStart:InsertTextInputFormField
        //GistId:b09907fef4643433271e4e0e912921b0
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertTextInput("TextInput", TextFormFieldType.REGULAR, "", "Hello", 0);

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertTextInputFormField.docx");
        //ExEnd:InsertTextInputFormField
    }

    @Test
    public void insertCheckBoxFormField() throws Exception
    {
        //ExStart:InsertCheckBoxFormField
        //GistId:b09907fef4643433271e4e0e912921b0
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertCheckBox("CheckBox", true, true, 0);

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertCheckBoxFormField.docx");
        //ExEnd:InsertCheckBoxFormField
    }

    @Test
    public void insertComboBoxFormField() throws Exception
    {
        //ExStart:InsertComboBoxFormField
        //GistId:b09907fef4643433271e4e0e912921b0
        String[] items = { "One", "Two", "Three" };

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertComboBox("DropDown", items, 0);

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertComboBoxFormField.docx");
        //ExEnd:InsertComboBoxFormField
    }

    @Test
    public void insertHtml() throws Exception
    {
        //ExStart:InsertHtml
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertHtml(
            "<P align='right'>Paragraph right</P>" +
            "<b>Implicit paragraph left</b>" +
            "<div align='center'>Div center</div>" +
            "<h1 align='left'>Heading 1 left.</h1>");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertHtml.docx");
        //ExEnd:InsertHtml
    }

    @Test
    public void insertHyperlink() throws Exception
    {
        //ExStart:InsertHyperlink
        //GistId:0213851d47551e83af42233f4d075cf6
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Please make sure to visit ");

        builder.getFont().setStyle(doc.getStyles().getByStyleIdentifier(StyleIdentifier.HYPERLINK));
        builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);
        builder.getFont().clearFormatting();

        builder.write(" for more information.");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertHyperlink.docx");
        //ExEnd:InsertHyperlink
    }

    @Test
    public void insertTableOfContents() throws Exception
    {
        //ExStart:InsertTableOfContents
        //GistId:db118a3e1559b9c88355356df9d7ea10
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        
        // Start the actual document content on the second page.
        builder.insertBreak(BreakType.PAGE_BREAK);

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        builder.writeln("Heading 2");
        builder.writeln("Heading 3");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 3.1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);

        builder.writeln("Heading 3.1.1");
        builder.writeln("Heading 3.1.2");
        builder.writeln("Heading 3.1.3");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 3.2");
        builder.writeln("Heading 3.3");

        //ExStart:UpdateFields
        //GistId:db118a3e1559b9c88355356df9d7ea10
        // The newly inserted table of contents will be initially empty.
        // It needs to be populated by updating the fields in the document.
        doc.updateFields();
        //ExEnd:UpdateFields

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertTableOfContents.docx");
        //ExEnd:InsertTableOfContents
    }

    @Test
    public void insertInlineImage() throws Exception
    {
        //ExStart:InsertInlineImage
        //GistId:6f849e51240635a6322ab0460938c922
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImagesDir() + "Transparent background logo.png");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertInlineImage.docx");
        //ExEnd:InsertInlineImage
    }

    @Test
    public void insertFloatingImage() throws Exception
    {
        //ExStart:InsertFloatingImage
        //GistId:6f849e51240635a6322ab0460938c922
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertImage(getImagesDir() + "Transparent background logo.png",
            RelativeHorizontalPosition.MARGIN,
            100.0,
            RelativeVerticalPosition.MARGIN,
            100.0,
            200.0,
            100.0,
            WrapType.SQUARE);

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertFloatingImage.docx");
        //ExEnd:InsertFloatingImage
    }

    @Test
    public void insertParagraph() throws Exception
    {
        //ExStart:InsertParagraph
        //GistId:ecf2c438314e6c8318ca9833c7f62326
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Font font = builder.getFont();
        font.setSize(16.0);
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setName("Arial");
        font.setUnderline(Underline.DASH);

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        paragraphFormat.setFirstLineIndent(8.0);
        paragraphFormat.setAlignment(ParagraphAlignment.JUSTIFY);
        paragraphFormat.setKeepTogether(true);

        builder.writeln("A whole paragraph.");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertParagraph.docx");
        //ExEnd:InsertParagraph
    }

    @Test
    public void insertTcField() throws Exception
    {
        //ExStart:InsertTcField
        //GistId:db118a3e1559b9c88355356df9d7ea10
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("TC \"Entry Text\" \\f t");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.InsertTcField.docx");
        //ExEnd:InsertTcField
    }

    @Test
    public void insertTcFieldsAtText() throws Exception
    {
        //ExStart:InsertTcFieldsAtText
        //GistId:db118a3e1559b9c88355356df9d7ea10
        Document doc = new Document();

        FindReplaceOptions options = new FindReplaceOptions();
        options.getApplyFont().setHighlightColor(msColor.getDarkOrange());
        options.setReplacingCallback(new InsertTCFieldHandler("Chapter 1", "\\l 1"));

        doc.getRange().replaceInternal(new Regex("The Beginning"), "", options);
        //ExEnd:InsertTcFieldsAtText
    }

    //ExStart:InsertTCFieldHandler
    public final static class InsertTCFieldHandler implements IReplacingCallback
    {
        // Store the text and switches to be used for the TC fields.
        private /*final*/ String mFieldText;
        private /*final*/ String mFieldSwitches;

        /// <summary>
        /// The display text and switches to use for each TC field. Display name can be an empty string or null.
        /// </summary>
        public InsertTCFieldHandler(String text, String switches)
        {
            mFieldText = text;
            mFieldSwitches = switches;
        }

        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs args) throws Exception
        {
            DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
            builder.moveTo(args.getMatchNode());

            // If the user-specified text to be used in the field as display text, then use that,
            // otherwise use the match string as the display text.
            String insertText = !msString.isNullOrEmpty(mFieldText) ? mFieldText : args.getMatchInternal().getValue();

            builder.insertField($"TC \"{insertText}\" {mFieldSwitches}");

            return ReplaceAction.SKIP;
        }
    }
    //ExEnd:InsertTCFieldHandler
    
    @Test
    public void cursorPosition() throws Exception
    {
        //ExStart:CursorPosition
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Node curNode = builder.getCurrentNode();
        Paragraph curParagraph = builder.getCurrentParagraph();
        //ExEnd:CursorPosition

        System.out.println("\nCursor move to paragraph: " + curParagraph.getText());
    }

    @Test
    public void moveToNode() throws Exception
    {
        //ExStart:MoveToNode
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        //ExStart:MoveToBookmark
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a bookmark and add content to it using a DocumentBuilder.
        builder.startBookmark("MyBookmark");
        builder.writeln("Bookmark contents.");
        builder.endBookmark("MyBookmark");

        // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark.
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getBookmarkEnd(), builder.getCurrentParagraph().getFirstChild());

        // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this.
        builder.moveToBookmark("MyBookmark");

        // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it.
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getBookmarkStart(), builder.getCurrentParagraph().getFirstChild());

        // We can move the builder to an individual node,
        // which in this case will be the first node of the first paragraph, like this.
        builder.moveTo(doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.ANY, false).get(0));
        //ExEnd:MoveToBookmark

        Assert.assertEquals(NodeType.BOOKMARK_START, builder.getCurrentNode().getNodeType());
        Assert.assertTrue(builder.isAtStartOfParagraph());

        // A shorter way of moving the very start/end of a document is with these methods.
        builder.moveToDocumentEnd();
        Assert.assertTrue(builder.isAtEndOfParagraph());
        builder.moveToDocumentStart();
        Assert.assertTrue(builder.isAtStartOfParagraph());
        //ExEnd:MoveToNode
    }

    @Test
    public void moveToDocumentStartEnd() throws Exception
    {
        //ExStart:MoveToDocumentStartEnd
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor position to the beginning of your document.
        builder.moveToDocumentStart();
        System.out.println("\nThis is the beginning of the document.");

        // Move the cursor position to the end of your document.
        builder.moveToDocumentEnd();
        System.out.println("\nThis is the end of the document.");
        //ExEnd:MoveToDocumentStartEnd
    }

    @Test
    public void moveToSection() throws Exception
    {
        //ExStart:MoveToSection
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        Document doc = new Document();
        doc.appendChild(new Section(doc));

        // Move a DocumentBuilder to the second section and add text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToSection(1);
        builder.writeln("Text added to the 2nd section.");

        // Create document with paragraphs.
        doc = new Document(getMyDir() + "Paragraphs.docx");
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        Assert.assertEquals(22, paragraphs.getCount());

        // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
        // and any content added by the DocumentBuilder will just be prepended to the document.
        builder = new DocumentBuilder(doc);
        Assert.assertEquals(0, paragraphs.indexOf(builder.getCurrentParagraph()));

        // You can move the cursor to any position in a paragraph.
        builder.moveToParagraph(2, 10);
        Assert.assertEquals(2, paragraphs.indexOf(builder.getCurrentParagraph()));
        builder.writeln("This is a new third paragraph. ");
        Assert.assertEquals(3, paragraphs.indexOf(builder.getCurrentParagraph()));
        //ExEnd:MoveToSection
    }

    @Test
    public void moveToHeadersFooters() throws Exception
    {
        //ExStart:MoveToHeadersFooters
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want headers and footers different for first, even and odd pages.
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        // Create the headers.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header for even pages");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for all other pages");

        // Create two pages in the document.
        builder.moveToSection(0);
        builder.writeln("Page1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page2");

        doc.save(getArtifactsDir() + "AddContentUsingDocumentBuilder.MoveToHeadersFooters.docx");
        //ExEnd:MoveToHeadersFooters
    }

    @Test
    public void moveToParagraph() throws Exception
    {
        //ExStart:MoveToParagraph
        Document doc = new Document(getMyDir() + "Paragraphs.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToParagraph(2, 0);
        builder.writeln("This is the 3rd paragraph.");
        //ExEnd:MoveToParagraph
    }

    @Test
    public void moveToTableCell() throws Exception
    {
        //ExStart:MoveToTableCell
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        Document doc = new Document(getMyDir() + "Tables.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to row 3, cell 4 of the first table.
        builder.moveToCell(0, 2, 3, 0);
        builder.write("\nCell contents added by DocumentBuilder");
        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        Assert.assertEquals(table.getRows().get(2).getCells().get(3), builder.getCurrentNode().getParentNode().getParentNode());
        Assert.assertEquals("Cell contents added by DocumentBuilderCell 3 contents\u0007", table.getRows().get(2).getCells().get(3).getText().trim());
        //ExEnd:MoveToTableCell
    }

    @Test
    public void moveToBookmarkEnd() throws Exception
    {
        //ExStart:MoveToBookmarkEnd
        //GistId:ecf2c438314e6c8318ca9833c7f62326
        Document doc = new Document(getMyDir() + "Bookmarks.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToBookmark("MyBookmark1", false, true);
        builder.writeln("This is a bookmark.");
        //ExEnd:MoveToBookmarkEnd
    }

    @Test
    public void moveToMergeField() throws Exception
    {
        //ExStart:MoveToMergeField
        //GistId:1a2c340d1a9dde6fe70c2733084d9aab
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a field using the DocumentBuilder and add a run of text after it.
        Field field = builder.insertField("MERGEFIELD field");
        builder.write(" Text after the field.");

        // The builder's cursor is currently at end of the document.
        Assert.assertNull(builder.getCurrentNode());
        // We can move the builder to a field like this, placing the cursor at immediately after the field.
        builder.moveToField(field, true);

        // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field.
        // If we wish to move the DocumentBuilder to inside a field,
        // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method.
        Assert.assertEquals(field.getEnd(), builder.getCurrentNode().getPreviousSibling());
        builder.write(" Text immediately after the field.");
        //ExEnd:MoveToMergeField
    }
}
