package DocsExamples.Programming_with_Documents.Contents_Management;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.Table;
import java.util.ArrayList;
import com.aspose.words.Node;
import java.util.Collections;
import com.aspose.words.Section;
import com.aspose.words.Bookmark;
import com.aspose.words.BookmarkStart;
import com.aspose.words.BookmarkEnd;
import com.aspose.words.CommentRangeStart;
import com.aspose.words.CommentRangeEnd;
import com.aspose.words.Run;
import com.aspose.ms.System.msConsole;
import com.aspose.words.SaveFormat;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldSeparator;
import com.aspose.words.FieldEnd;
import com.aspose.words.ControlChar;
import com.aspose.words.Body;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.HeaderFooter;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.aspose.words.FieldType;
import com.aspose.words.FieldHyperlink;
import com.aspose.words.NodeCollection;
import com.aspose.ms.System.msString;
import com.aspose.words.Shape;


public class ExtractContent extends DocsExamplesBase
{
    @Test
    public void extractContentBetweenBlockLevelNodes() throws Exception
    {
        //ExStart:ExtractContentBetweenBlockLevelNodes
        Document doc = new Document(getMyDir() + "Extract content.docx");

        Paragraph startPara = (Paragraph) doc.getLastSection().getChild(NodeType.PARAGRAPH, 2, true);
        Table endTable = (Table) doc.getLastSection().getChild(NodeType.TABLE, 0, true);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        ArrayList<Node> extractedNodes = ExtractContentHelper.extractContent(startPara, endTable, true);

        // Let's reverse the array to make inserting the content back into the document easier.
        Collections.reverse(extractedNodes);

        while (extractedNodes.size() > 0)
        {
            // Insert the last node from the reversed list.
            endTable.getParentNode().insertAfter((Node) extractedNodes.get(0), endTable);
            // Remove this node from the list after insertion.
            extractedNodes.remove(0);
        }

        doc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenBlockLevelNodes.docx");
        //ExEnd:ExtractContentBetweenBlockLevelNodes
    }

    @Test
    public void extractContentBetweenBookmark() throws Exception
    {
        //ExStart:ExtractContentBetweenBookmark
        Document doc = new Document(getMyDir() + "Extract content.docx");

        Section section = doc.getSections().get(0);
        section.getPageSetup().setLeftMargin(70.85);

        // Retrieve the bookmark from the document.
        Bookmark bookmark = doc.getRange().getBookmarks().get("Bookmark1");
        // We use the BookmarkStart and BookmarkEnd nodes as markers.
        BookmarkStart bookmarkStart = bookmark.getBookmarkStart();
        BookmarkEnd bookmarkEnd = bookmark.getBookmarkEnd();

        // Firstly, extract the content between these nodes, including the bookmark.
        ArrayList<Node> extractedNodesInclusive = ExtractContentHelper.extractContent(bookmarkStart, bookmarkEnd, true);
        
        Document dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesInclusive);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenBookmark.IncludingBookmark.docx");

        // Secondly, extract the content between these nodes this time without including the bookmark.
        ArrayList<Node> extractedNodesExclusive = ExtractContentHelper.extractContent(bookmarkStart, bookmarkEnd, false);
        
        dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesExclusive);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenBookmark.WithoutBookmark.docx");
        //ExEnd:ExtractContentBetweenBookmark
    }

    @Test
    public void extractContentBetweenCommentRange() throws Exception
    {
        //ExStart:ExtractContentBetweenCommentRange
        Document doc = new Document(getMyDir() + "Extract content.docx");

        // This is a quick way of getting both comment nodes.
        // Your code should have a proper method of retrieving each corresponding start and end node.
        CommentRangeStart commentStart = (CommentRangeStart) doc.getChild(NodeType.COMMENT_RANGE_START, 0, true);
        CommentRangeEnd commentEnd = (CommentRangeEnd) doc.getChild(NodeType.COMMENT_RANGE_END, 0, true);

        // Firstly, extract the content between these nodes including the comment as well.
        ArrayList<Node> extractedNodesInclusive = ExtractContentHelper.extractContent(commentStart, commentEnd, true);
        
        Document dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesInclusive);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenCommentRange.IncludingComment.docx");

        // Secondly, extract the content between these nodes without the comment.
        ArrayList<Node> extractedNodesExclusive = ExtractContentHelper.extractContent(commentStart, commentEnd, false);
        
        dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodesExclusive);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenCommentRange.WithoutComment.docx");
        //ExEnd:ExtractContentBetweenCommentRange
    }

    @Test
    public void extractContentBetweenParagraphs() throws Exception
    {
        //ExStart:ExtractContentBetweenParagraphs
        Document doc = new Document(getMyDir() + "Extract content.docx");

        Paragraph startPara = (Paragraph) doc.getFirstSection().getBody().getChild(NodeType.PARAGRAPH, 6, true);
        Paragraph endPara = (Paragraph) doc.getFirstSection().getBody().getChild(NodeType.PARAGRAPH, 10, true);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        ArrayList<Node> extractedNodes = ExtractContentHelper.extractContent(startPara, endPara, true);

        Document dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenParagraphs.docx");
        //ExEnd:ExtractContentBetweenParagraphs
    }

    @Test
    public void extractContentBetweenParagraphStyles() throws Exception
    {
        //ExStart:ExtractContentBetweenParagraphStyles
        Document doc = new Document(getMyDir() + "Extract content.docx");

        // Gather a list of the paragraphs using the respective heading styles.
        ArrayList<Paragraph> parasStyleHeading1 = ExtractContentHelper.paragraphsByStyleName(doc, "Heading 1");
        ArrayList<Paragraph> parasStyleHeading3 = ExtractContentHelper.paragraphsByStyleName(doc, "Heading 3");

        // Use the first instance of the paragraphs with those styles.
        Node startPara1 = parasStyleHeading1.get(0);
        Node endPara1 = parasStyleHeading3.get(0);

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        ArrayList<Node> extractedNodes = ExtractContentHelper.extractContent(startPara1, endPara1, false);

        Document dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentBetweenParagraphStyles.docx");
        //ExEnd:ExtractContentBetweenParagraphStyles
    }

    @Test
    public void extractContentBetweenRuns() throws Exception
    {
        //ExStart:ExtractContentBetweenRuns
        Document doc = new Document(getMyDir() + "Extract content.docx");

        Paragraph para = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 7, true);

        Run startRun = para.getRuns().get(1);
        Run endRun = para.getRuns().get(4);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        ArrayList<Node> extractedNodes = ExtractContentHelper.extractContent(startRun, endRun, true);

        Node node = (Node) extractedNodes.get(0);
        System.out.println(node.toString(SaveFormat.TEXT));
        //ExEnd:ExtractContentBetweenRuns
    }

    @Test
    public void extractContentUsingDocumentVisitor() throws Exception
    {
        //ExStart:ExtractContentUsingDocumentVisitor
        Document doc = new Document(getMyDir() + "Absolute position tab.docx");

        MyDocToTxtWriter myConverter = new MyDocToTxtWriter();

        // This is the well known Visitor pattern. Get the model to accept a visitor.
        // The model will iterate through itself by calling the corresponding methods.
        // On the visitor object (this is called visiting). 
        // Note that every node in the object model has the accept method so the visiting
        // can be executed not only for the whole document, but for any node in the document.
        doc.accept(myConverter);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // That in this example, has accumulated in the visitor.
        System.out.println(myConverter.getText());
        //ExEnd:ExtractContentUsingDocumentVisitor
    }

    //ExStart:MyDocToTxtWriter
    /// <summary>
    /// Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
    /// </summary>
    static class MyDocToTxtWriter extends DocumentVisitor
    {
        public MyDocToTxtWriter()
        {
            mIsSkipText = false;
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
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            appendText(run.getText());

            // Let the visitor continue visiting other nodes.
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldStart(FieldStart fieldStart)
        {
            // In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
            // after a field start character. We want to skip field codes and output field.
            // Result only, therefore we use a flag to suspend the output while inside a field code.
            // Note this is a very simplistic implementation and will not work very well.
            // If you have nested fields in a document.
            mIsSkipText = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldSeparator(FieldSeparator fieldSeparator)
        {
            // Once reached a field separator node, we enable the output because we are
            // now entering the field result nodes.
            mIsSkipText = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldEnd(FieldEnd fieldEnd)
        {
            // Make sure we enable the output when reached a field end because some fields
            // do not have field separator and do not have field result.
            mIsSkipText = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when visiting of a Paragraph node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitParagraphEnd(Paragraph paragraph)
        {
            // When outputting to plain text we output Cr+Lf characters.
            appendText(ControlChar.CR_LF);

            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitBodyStart(Body body)
        {
            // We can detect beginning and end of all composite nodes such as Section, Body, 
            // Table, Paragraph etc and provide custom handling for them.
            msStringBuilder.append(mBuilder, "*** Body Started ***\r\n");

            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitBodyEnd(Body body)
        {
            msStringBuilder.append(mBuilder, "*** Body Ended ***\r\n");
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a HeaderFooter node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitHeaderFooterStart(HeaderFooter headerFooter)
        {
            // Returning this value from a visitor method causes visiting of this
            // Node to stop and move on to visiting the next sibling node
            // The net effect in this example is that the text of headers and footers
            // Is not included in the resulting output
            return VisitorAction.SKIP_THIS_NODE;
        }

        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        private void appendText(String text)
        {
            if (!mIsSkipText)
                msStringBuilder.append(mBuilder, text);
        }

        private /*final*/ StringBuilder mBuilder;
        private boolean mIsSkipText;
    }
    //ExEnd:MyDocToTxtWriter
    
    @Test
    public void extractContentUsingField() throws Exception
    {
        //ExStart:ExtractContentUsingField
        Document doc = new Document(getMyDir() + "Extract content.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.moveToMergeField("Fullname", false, false);

        // The builder cursor should be positioned at the start of the field.
        FieldStart startField = (FieldStart) builder.getCurrentNode();
        Paragraph endPara = (Paragraph) doc.getFirstSection().getChild(NodeType.PARAGRAPH, 5, true);

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        ArrayList<Node> extractedNodes = ExtractContentHelper.extractContent(startField, endPara, false);

        Document dstDoc = ExtractContentHelper.generateDocument(doc, extractedNodes);
        dstDoc.save(getArtifactsDir() + "ExtractContent.ExtractContentUsingField.docx");
        //ExEnd:ExtractContentUsingField
    }

    @Test
    public void extractTableOfContents() throws Exception
    {
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        for (Field field : doc.getRange().getFields())
        {
            if (field.getType() == FieldType.FIELD_HYPERLINK)
            {
                FieldHyperlink hyperlink = (FieldHyperlink) field;
                if (hyperlink.getSubAddress() != null && hyperlink.getSubAddress().startsWith("_Toc"))
                {
                    Paragraph tocItem = (Paragraph) field.getStart().getAncestor(NodeType.PARAGRAPH);
                    
                    System.out.println(tocItem.toString(SaveFormat.TEXT).trim());
                    System.out.println("------------------");

                    Bookmark bm = doc.getRange().getBookmarks().get(hyperlink.getSubAddress());
                    Paragraph pointer = (Paragraph) bm.getBookmarkStart().getAncestor(NodeType.PARAGRAPH);
                    
                    System.out.println(pointer.toString(SaveFormat.TEXT));
                }
            }
        }
    }

    @Test
    public void extractTextOnly() throws Exception
    {
        //ExStart:ExtractTextOnly
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertField("MERGEFIELD Field");

        System.out.println("GetText() Result: " + doc.getText());

        // When converted to text it will not retrieve fields code or special characters,
        // but will still contain some natural formatting characters such as paragraph markers etc. 
        // This is the same as "viewing" the document as if it was opened in a text editor.
        System.out.println("ToString() Result: " + doc.toString(SaveFormat.TEXT));
        //ExEnd:ExtractTextOnly            
    }

    @Test
    public void extractContentBasedOnStyles() throws Exception
    {
        //ExStart:ExtractContentBasedOnStyles
        Document doc = new Document(getMyDir() + "Styles.docx");

        final String PARA_STYLE = "Heading 1";
        final String RUN_STYLE = "Intense Emphasis";

        ArrayList<Paragraph> paragraphs = paragraphsByStyleName(doc, PARA_STYLE);
        System.out.println("Paragraphs with \"{paraStyle}\" styles ({paragraphs.Count}):");
        
        for (Paragraph paragraph : paragraphs)
            msConsole.write(paragraph.toString(SaveFormat.TEXT));

        ArrayList<Run> runs = runsByStyleName(doc, RUN_STYLE);
        System.out.println("\nRuns with \"{runStyle}\" styles ({runs.Count}):");
        
        for (Run run : runs)
            System.out.println(run.getRange().getText());
        //ExEnd:ExtractContentBasedOnStyles
    }

    //ExStart:ParagraphsByStyleName
    public ArrayList<Paragraph> paragraphsByStyleName(Document doc, String styleName)
    {
        ArrayList<Paragraph> paragraphsWithStyle = new ArrayList<Paragraph>();
        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);
        
        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs)
        {
            if (msString.equals(paragraph.getParagraphFormat().getStyle().getName(), styleName))
                paragraphsWithStyle.add(paragraph);
        }

        return paragraphsWithStyle;
    }
    //ExEnd:ParagraphsByStyleName
    
    //ExStart:RunsByStyleName
    public ArrayList<Run> runsByStyleName(Document doc, String styleName)
    {
        ArrayList<Run> runsWithStyle = new ArrayList<Run>();
        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);
        
        for (Run run : (Iterable<Run>) runs)
        {
            if (msString.equals(run.getFont().getStyle().getName(), styleName))
                runsWithStyle.add(run);
        }

        return runsWithStyle;
    }
    //ExEnd:RunsByStyleName

    @Test
    public void extractPrintText() throws Exception
    {
        //ExStart:ExtractText
        Document doc = new Document(getMyDir() + "Tables.docx");

        
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // The range text will include control characters such as "\a" for a cell.
        // You can call ToString and pass SaveFormat.Text on the desired node to find the plain text content.

        System.out.println("Contents of the table: ");
        System.out.println(table.getRange().getText());
        //ExEnd:ExtractText   

        //ExStart:PrintTextRangeOFRowAndTable
        System.out.println("\nContents of the row: ");
        System.out.println(table.getRows().get(1).getRange().getText());

        System.out.println("\nContents of the cell: ");
        System.out.println(table.getLastRow().getLastCell().getRange().getText());
        //ExEnd:PrintTextRangeOFRowAndTable
    }

    @Test
    public void extractImagesToFiles() throws Exception
    {
        //ExStart:ExtractImagesToFiles
        Document doc = new Document(getMyDir() + "Images.docx");

        NodeCollection shapes = doc.getChildNodes(NodeType.SHAPE, true);
        int imageIndex = 0;
        
        for (Shape shape : (Iterable<Shape>) shapes)
        {
            if (shape.hasImage())
            {
                String imageFileName =
                    $"Image.ExportImages.{imageIndex}_{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";

                shape.getImageData().save(getArtifactsDir() + imageFileName);
                imageIndex++;
            }
        }
        //ExEnd:ExtractImagesToFiles
    }
}
