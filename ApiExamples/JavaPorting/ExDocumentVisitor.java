// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.NodeType;
import com.aspose.words.Section;
import com.aspose.words.NodeCollection;
import com.aspose.words.Body;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.SubDocument;
import com.aspose.ms.System.Text.msStringBuilder;
import org.testng.Assert;
import com.aspose.words.Table;
import com.aspose.words.Row;
import com.aspose.ms.System.msString;
import com.aspose.words.Cell;
import com.aspose.words.CommentRangeStart;
import com.aspose.words.CommentRangeEnd;
import com.aspose.words.Comment;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldEnd;
import com.aspose.words.FieldSeparator;
import com.aspose.words.HeaderFooter;
import com.aspose.words.EditableRangeStart;
import com.aspose.words.EditableRangeEnd;
import com.aspose.words.Footnote;
import com.aspose.words.OfficeMath;
import com.aspose.words.SmartTag;
import com.aspose.words.StructuredDocumentTag;


@Test
public class ExDocumentVisitor extends ApiExampleBase
{
    //ExStart
    //ExFor:Document.Accept(DocumentVisitor)
    //ExFor:Body.Accept(DocumentVisitor)
    //ExFor:SubDocument.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor
    //ExFor:DocumentVisitor.VisitRun(Run)
    //ExFor:DocumentVisitor.VisitDocumentEnd(Document)
    //ExFor:DocumentVisitor.VisitDocumentStart(Document)
    //ExFor:DocumentVisitor.VisitSectionEnd(Section)
    //ExFor:DocumentVisitor.VisitSectionStart(Section)
    //ExFor:DocumentVisitor.VisitBodyStart(Body)
    //ExFor:DocumentVisitor.VisitBodyEnd(Body)
    //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
    //ExFor:DocumentVisitor.VisitParagraphEnd(Paragraph)
    //ExFor:DocumentVisitor.VisitSubDocument(SubDocument)
    //ExSummary:Traverse a document with a visitor that prints all structure nodes that it encounters.
    @Test //ExSkip
    public void docStructureToText() throws Exception
    {
        // Open the document that has nodes we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        DocStructurePrinter visitor = new DocStructurePrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testDocStructureToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about sections, bodies, paragraphs and runs encountered in the document.
    /// </summary>
    public static class DocStructurePrinter extends DocumentVisitor
    {
        public DocStructurePrinter()
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
        /// Called when a Document node is encountered.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitDocumentStart(Document doc)
        {
            int childNodeCount = doc.getChildNodes(NodeType.ANY, true).getCount();

            // A Document node is at the root of every document, so if we let a document accept a visitor, this will be the first visitor action to be carried out
            indentAndAppendLine("[Document start] Child nodes: " + childNodeCount);
            mDocTraversalDepth++;

            // Let the visitor continue visiting other nodes
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Document is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitDocumentEnd(Document doc)
        {
            // If we let a document accept a visitor, this will be the last visitor action to be carried out
            mDocTraversalDepth--;
            indentAndAppendLine("[Document end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Section node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSectionStart(Section section)
        {
            // Get the index of our section within the document
            NodeCollection docSections = section.getDocument().getChildNodes(NodeType.SECTION, false);
            int sectionIndex = docSections.indexOf(section);

            indentAndAppendLine("[Section start] Section index: " + sectionIndex);
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Section node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSectionEnd(Section section)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Section end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Body node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitBodyStart(Body body)
        {
            int paragraphCount = body.getParagraphs().getCount();
            indentAndAppendLine("[Body start] Paragraphs: " + paragraphCount);
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Body node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitBodyEnd(Body body)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Body end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitParagraphStart(Paragraph paragraph)
        {
            indentAndAppendLine("[Paragraph start]");
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Paragraph node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitParagraphEnd(Paragraph paragraph)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Paragraph end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a SubDocument node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSubDocument(SubDocument subDocument)
        {
            indentAndAppendLine("[SubDocument]");
            
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testDocStructureToText(DocStructurePrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[Document start]"));
        Assert.assertTrue(visitorText.contains("[Document end]"));
        Assert.assertTrue(visitorText.contains("[Section start]"));
        Assert.assertTrue(visitorText.contains("[Section end]"));
        Assert.assertTrue(visitorText.contains("[Body start]"));
        Assert.assertTrue(visitorText.contains("[Body end]"));
        Assert.assertTrue(visitorText.contains("[Paragraph start]"));
        Assert.assertTrue(visitorText.contains("[Paragraph end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
        Assert.assertTrue(visitorText.contains("[SubDocument]"));
    }

    //ExStart
    //ExFor:Cell.Accept(DocumentVisitor)
    //ExFor:Cell.IsFirstCell
    //ExFor:Cell.IsLastCell
    //ExFor:DocumentVisitor.VisitTableEnd(Tables.Table)
    //ExFor:DocumentVisitor.VisitTableStart(Tables.Table)
    //ExFor:DocumentVisitor.VisitRowEnd(Tables.Row)
    //ExFor:DocumentVisitor.VisitRowStart(Tables.Row)
    //ExFor:DocumentVisitor.VisitCellStart(Tables.Cell)
    //ExFor:DocumentVisitor.VisitCellEnd(Tables.Cell)
    //ExFor:Row.Accept(DocumentVisitor)
    //ExFor:Row.FirstCell
    //ExFor:Row.GetText
    //ExFor:Row.IsFirstRow
    //ExFor:Row.LastCell
    //ExFor:Row.ParentTable
    //ExSummary:Traverse a document with a visitor that prints all tables that it encounters.
    @Test //ExSkip
    public void tableToText() throws Exception
    {
        // Open the document that has tables we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        TableInfoPrinter visitor = new TableInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testTableToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about and contents of all tables encountered in the document.
    /// </summary>
    public static class TableInfoPrinter extends DocumentVisitor
    {
        public TableInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideTable = false;
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
            // We want to print the contents of runs, but only if they consist of text from cells
            // So we are only interested in runs that are children of table nodes
            if (mVisitorIsInsideTable) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Table is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitTableStart(Table table)
        {
            int rows = 0;
            int columns = 0;

            if (table.getRows().getCount() > 0)
            {
                rows = table.getRows().getCount();
                columns = table.getFirstRow().getCount();
            }

            indentAndAppendLine("[Table start] Size: " + rows + "x" + columns);
            mDocTraversalDepth++;
            mVisitorIsInsideTable = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Table node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitTableEnd(Table table)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Table end]");
            mVisitorIsInsideTable = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Row node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRowStart(Row row)
        {
            String rowContents = msString.trimEnd(row.getText(), new char[]{ '\u0007', ' ' }).replace("\u0007", ", ");
            int rowWidth = row.indexOf(row.getLastCell()) + 1;
            int rowIndex = row.getParentTable().indexOf(row);
            String rowStatusInTable = row.isFirstRow() && row.isLastRow() ? "only" : row.isFirstRow() ? "first" : row.isLastRow() ? "last" : "";
            if (!"".equals(rowStatusInTable))
            {
                rowStatusInTable = $", the {rowStatusInTable} row in this table,";
            }

            indentAndAppendLine($"[Row start] Row #{++rowIndex}{rowStatusInTable} width {rowWidth}, \"{rowContents}\"");
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Row node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRowEnd(Row row)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Row end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Cell node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCellStart(Cell cell)
        {
            Row row = cell.getParentRow();
            Table table = row.getParentTable();
            String cellStatusInRow = cell.isFirstCell() && cell.isLastCell() ? "only" : cell.isFirstCell() ? "first" : cell.isLastCell() ? "last" : "";
            if (!"".equals(cellStatusInRow))
            {
                cellStatusInRow = $", the {cellStatusInRow} cell in this row";
            }

            indentAndAppendLine($"[Cell start] Row {table.IndexOf(row) + 1}, Col {row.IndexOf(cell) + 1}{cellStatusInRow}");
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Cell node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCellEnd(Cell cell)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Cell end]");
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                msStringBuilder.append(mBuilder, "|  ");
            }

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideTable;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testTableToText(TableInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[Table start]"));
        Assert.assertTrue(visitorText.contains("[Table end]"));
        Assert.assertTrue(visitorText.contains("[Row start]"));
        Assert.assertTrue(visitorText.contains("[Row end]"));
        Assert.assertTrue(visitorText.contains("[Cell start]"));
        Assert.assertTrue(visitorText.contains("[Cell end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitCommentStart(Comment)
    //ExFor:DocumentVisitor.VisitCommentEnd(Comment)
    //ExFor:DocumentVisitor.VisitCommentRangeEnd(CommentRangeEnd)
    //ExFor:DocumentVisitor.VisitCommentRangeStart(CommentRangeStart)
    //ExSummary:Traverse a document with a visitor that prints all comment nodes that it encounters.
    @Test //ExSkip
    public void commentsToText() throws Exception
    {
        // Open the document that has comments/comment ranges we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        CommentInfoPrinter visitor = new CommentInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testCommentsToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about and contents of comments and comment ranges encountered in the document.
    /// </summary>
    public static class CommentInfoPrinter extends DocumentVisitor
    {
        public CommentInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideComment = false;
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
            if (mVisitorIsInsideComment) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentRangeStart(CommentRangeStart commentRangeStart)
        {
            indentAndAppendLine("[Comment range start] ID: " + commentRangeStart.getId());
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment range end]");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentStart(Comment comment)
        {
            indentAndAppendLine(
                $"[Comment start] For comment range ID {comment.Id}, By {comment.Author} on {comment.DateTime}");
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Comment node is ended in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitCommentEnd(Comment comment)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment end]");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                msStringBuilder.append(mBuilder, "|  ");
            }

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideComment;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testCommentsToText(CommentInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[Comment range start]"));
        Assert.assertTrue(visitorText.contains("[Comment range end]"));
        Assert.assertTrue(visitorText.contains("[Comment start]"));
        Assert.assertTrue(visitorText.contains("[Comment end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitFieldStart
    //ExFor:DocumentVisitor.VisitFieldEnd
    //ExFor:DocumentVisitor.VisitFieldSeparator
    //ExSummary:Traverse a document with a visitor that prints all fields that it encounters.
    @Test //ExSkip
    public void fieldToText() throws Exception
    {
        // Open the document that has fields that we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        FieldInfoPrinter visitor = new FieldInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testFieldToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about fields encountered in the document.
    /// </summary>
    public static class FieldInfoPrinter extends DocumentVisitor
    {
        public FieldInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideField = false;
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
            if (mVisitorIsInsideField) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldStart(FieldStart fieldStart)
        {
            indentAndAppendLine("[Field start] FieldType: " + fieldStart.getFieldType());
            mDocTraversalDepth++;
            mVisitorIsInsideField = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldEnd(FieldEnd fieldEnd)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Field end]");
            mVisitorIsInsideField = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFieldSeparator(FieldSeparator fieldSeparator)
        {
            indentAndAppendLine("[FieldSeparator]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++)
            {
                msStringBuilder.append(mBuilder, "|  ");
            }

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideField;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testFieldToText(FieldInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[Field start]"));
        Assert.assertTrue(visitorText.contains("[Field end]"));
        Assert.assertTrue(visitorText.contains("[FieldSeparator]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitHeaderFooterStart(HeaderFooter)
    //ExFor:DocumentVisitor.VisitHeaderFooterEnd(HeaderFooter)
    //ExFor:HeaderFooter.Accept(DocumentVisitor)
    //ExFor:HeaderFooterCollection.ToArray
    //ExFor:Run.Accept(DocumentVisitor)
    //ExFor:Run.GetText
    //ExSummary:Traverse a document with a visitor that prints all header/footer nodes that it encounters.
    @Test //ExSkip
    public void headerFooterToText() throws Exception
    {
        // Open the document that has headers and/or footers we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        HeaderFooterInfoPrinter visitor = new HeaderFooterInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());

        // An alternative way of visiting a document's header/footers section-by-section is by accessing the collection
        // We can also turn it into an array
        HeaderFooter[] headerFooters = doc.getFirstSection().getHeadersFooters().toArray();
        Assert.assertEquals(3, headerFooters.length);
        testHeaderFooterToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about HeaderFooter nodes encountered in the document.
    /// </summary>
    public static class HeaderFooterInfoPrinter extends DocumentVisitor
    {
        public HeaderFooterInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideHeaderFooter = false;
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
            if (mVisitorIsInsideHeaderFooter) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a HeaderFooter node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitHeaderFooterStart(HeaderFooter headerFooter)
        {
            indentAndAppendLine("[HeaderFooter start] HeaderFooterType: " + headerFooter.getHeaderFooterType());
            mDocTraversalDepth++;
            mVisitorIsInsideHeaderFooter = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a HeaderFooter node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitHeaderFooterEnd(HeaderFooter headerFooter)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[HeaderFooter end]");
            mVisitorIsInsideHeaderFooter = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideHeaderFooter;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testHeaderFooterToText(HeaderFooterInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[HeaderFooter start] HeaderFooterType: HeaderPrimary"));
        Assert.assertTrue(visitorText.contains("[HeaderFooter end]"));
        Assert.assertTrue(visitorText.contains("[HeaderFooter start] HeaderFooterType: HeaderFirst"));
        Assert.assertTrue(visitorText.contains("[HeaderFooter start] HeaderFooterType: HeaderEven"));
        Assert.assertTrue(visitorText.contains("[HeaderFooter start] HeaderFooterType: FooterPrimary"));
        Assert.assertTrue(visitorText.contains("[HeaderFooter start] HeaderFooterType: FooterFirst"));
        Assert.assertTrue(visitorText.contains("[HeaderFooter start] HeaderFooterType: FooterEven"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
    //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
    //ExSummary:Traverse a document with a visitor that prints all editable ranges that it encounters.
    @Test //ExSkip
    public void editableRangeToText() throws Exception
    {
        // Open the document that has editable ranges we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        EditableRangeInfoPrinter visitor = new EditableRangeInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        Paragraph p = new Paragraph(doc);
        p.appendChild(new Run(doc, "Paragraph with editable range text."));

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testEditableRangeToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about editable ranges encountered in the document.
    /// </summary>
    public static class EditableRangeInfoPrinter extends DocumentVisitor
    {
        public EditableRangeInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideEditableRange = false;
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
            // We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes
            if (mVisitorIsInsideEditableRange) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an EditableRange node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitEditableRangeStart(EditableRangeStart editableRangeStart)
        {
            indentAndAppendLine("[EditableRange start] ID: " + editableRangeStart.getId() + " Owner: " +
                                editableRangeStart.getEditableRange().getSingleUser());
            mDocTraversalDepth++;
            mVisitorIsInsideEditableRange = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a EditableRange node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[EditableRange end]");
            mVisitorIsInsideEditableRange = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideEditableRange;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd
    
    private void testEditableRangeToText(EditableRangeInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[EditableRange start]"));
        Assert.assertTrue(visitorText.contains("[EditableRange end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:Footnote.Accept(DocumentVisitor)
    //ExSummary:Traverse a document with a visitor that prints all footnotes that it encounters.
    @Test //ExSkip
    public void footnoteToText() throws Exception
    {
        // Open the document that has footnotes we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        FootnoteInfoPrinter visitor = new FootnoteInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testFootnoteToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about footnotes encountered in the document.
    /// </summary>
    public static class FootnoteInfoPrinter extends DocumentVisitor
    {
        public FootnoteInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideFootnote = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText()
        {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Footnote node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFootnoteStart(Footnote footnote)
        {
            indentAndAppendLine("[Footnote start] Type: " + footnote.getFootnoteType());
            mDocTraversalDepth++;
            mVisitorIsInsideFootnote = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a Footnote node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitFootnoteEnd(Footnote footnote)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[Footnote end]");
            mVisitorIsInsideFootnote = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            if (mVisitorIsInsideFootnote) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideFootnote;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testFootnoteToText(FootnoteInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[Footnote start] Type: Footnote"));
        Assert.assertTrue(visitorText.contains("[Footnote end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitOfficeMathEnd(Math.OfficeMath)
    //ExFor:DocumentVisitor.VisitOfficeMathStart(Math.OfficeMath)
    //ExFor:Math.MathObjectType
    //ExFor:Math.OfficeMath.Accept(DocumentVisitor)
    //ExFor:Math.OfficeMath.MathObjectType
    //ExSummary:Traverse a document with a visitor that prints all OfficeMath nodes that it encounters.
    @Test //ExSkip
    public void officeMathToText() throws Exception
    {
        // Open the document that has office math objects we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        OfficeMathInfoPrinter visitor = new OfficeMathInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testOfficeMathToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about office math objects encountered in the document.
    /// </summary>
    public static class OfficeMathInfoPrinter extends DocumentVisitor
    {
        public OfficeMathInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideOfficeMath = false;
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
            if (mVisitorIsInsideOfficeMath) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an OfficeMath node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitOfficeMathStart(OfficeMath officeMath)
        {
            indentAndAppendLine("[OfficeMath start] Math object type: " + officeMath.getMathObjectType());
            mDocTraversalDepth++;
            mVisitorIsInsideOfficeMath = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a OfficeMath node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitOfficeMathEnd(OfficeMath officeMath)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[OfficeMath end]");
            mVisitorIsInsideOfficeMath = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideOfficeMath;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testOfficeMathToText(OfficeMathInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: OMathPara"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: OMath"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: Argument"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: Supercript"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: SuperscriptPart"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: Fraction"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: Numerator"));
        Assert.assertTrue(visitorText.contains("[OfficeMath start] Math object type: Denominator"));
        Assert.assertTrue(visitorText.contains("[OfficeMath end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitSmartTagEnd(Markup.SmartTag)
    //ExFor:DocumentVisitor.VisitSmartTagStart(Markup.SmartTag)
    //ExSummary:Traverse a document with a visitor that prints all smart tag nodes that it encounters.
    @Test //ExSkip
    public void smartTagToText() throws Exception
    {
        // Open the document that has smart tags we want to print the info of
        Document doc = new Document(getMyDir() + "Smart tags.doc");

        // Create an object that inherits from the DocumentVisitor class
        SmartTagInfoPrinter visitor = new SmartTagInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testSmartTagToText(visitor); //ExEnd
    }

    /// <summary>
    /// This Visitor implementation prints information about smart tags encountered in the document.
    /// </summary>
    public static class SmartTagInfoPrinter extends DocumentVisitor
    {
        public SmartTagInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideSmartTag = false;
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
            if (mVisitorIsInsideSmartTag) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSmartTagStart(SmartTag smartTag)
        {
            indentAndAppendLine("[SmartTag start] Name: " + smartTag.getElement());
            mDocTraversalDepth++;
            mVisitorIsInsideSmartTag = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a SmartTag node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitSmartTagEnd(SmartTag smartTag)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[SmartTag end]");
            mVisitorIsInsideSmartTag = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private boolean mVisitorIsInsideSmartTag;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testSmartTagToText(SmartTagInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: address"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: Street"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: PersonName"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: title"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: GivenName"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: Sn"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: stockticker"));
        Assert.assertTrue(visitorText.contains("[SmartTag start] Name: date"));
        Assert.assertTrue(visitorText.contains("[SmartTag end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:StructuredDocumentTag.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagEnd(Markup.StructuredDocumentTag)
    //ExFor:DocumentVisitor.VisitStructuredDocumentTagStart(Markup.StructuredDocumentTag)
    //ExSummary:Traverse a document with a visitor that prints all structured document tag nodes that it encounters.
    @Test //ExSkip
    public void structuredDocumentTagToText() throws Exception
    {
        // Open the document that has structured document tags we want to print the info of
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");

        // Create an object that inherits from the DocumentVisitor class
        StructuredDocumentTagInfoPrinter visitor = new StructuredDocumentTagInfoPrinter();

        // Accepting a visitor lets it start traversing the nodes in the document, 
        // starting with the node that accepted it to then recursively visit every child
        doc.accept(visitor);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor
        System.out.println(visitor.getText());
        testStructuredDocumentTagToText(visitor); //ExSkip
    }

    /// <summary>
    /// This Visitor implementation prints information about structured document tags encountered in the document.
    /// </summary>
    public static class StructuredDocumentTagInfoPrinter extends DocumentVisitor
    {
        public StructuredDocumentTagInfoPrinter()
        {
            mBuilder = new StringBuilder();
            mVisitorIsInsideStructuredDocumentTag = false;
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
            if (mVisitorIsInsideStructuredDocumentTag) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a StructuredDocumentTag node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitStructuredDocumentTagStart(StructuredDocumentTag sdt)
        {
            indentAndAppendLine("[StructuredDocumentTag start] Title: " + sdt.getTitle());
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a StructuredDocumentTag node is ended.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
        {
            mDocTraversalDepth--;
            indentAndAppendLine("[StructuredDocumentTag end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(String text)
        {
            for (int i = 0; i < mDocTraversalDepth; i++) msStringBuilder.append(mBuilder, "|  ");

            msStringBuilder.appendLine(mBuilder, text);
        }

        private /*final*/ boolean mVisitorIsInsideStructuredDocumentTag;
        private int mDocTraversalDepth;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    private void testStructuredDocumentTagToText(StructuredDocumentTagInfoPrinter visitor)
    {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[StructuredDocumentTag start]"));
        Assert.assertTrue(visitorText.contains("[StructuredDocumentTag end]"));
    }
}
