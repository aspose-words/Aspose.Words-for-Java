package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;

@Test
public class ExDocumentVisitor extends ApiExampleBase {
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
    //ExSummary:Shows how to use a document visitor to print a document's node structure.
    @Test //ExSkip
    public void docStructureToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        DocStructurePrinter visitor = new DocStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testDocStructureToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's tree of child nodes.
    /// Creates a map of this tree in the form of a string.
    /// </summary>
    public static class DocStructurePrinter extends DocumentVisitor {
        public DocStructurePrinter() {
            mAcceptingNodeChildTree = new StringBuilder();
        }

        public String getText() {
            return mAcceptingNodeChildTree.toString();
        }

        /// <summary>
        /// Called when a Document node is encountered.
        /// </summary>
        public int visitDocumentStart(Document doc) {
            int childNodeCount = doc.getChildNodes(NodeType.ANY, true).getCount();

            indentAndAppendLine("[Document start] Child nodes: " + childNodeCount);
            mDocTraversalDepth++;

            // Allow the visitor to continue visiting other nodes.
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Document node have been visited.
        /// </summary>
        public int visitDocumentEnd(Document doc) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Document end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Section node is encountered in the document.
        /// </summary>
        public int visitSectionStart(final Section section) {
            // Get the index of our section within the document
            NodeCollection docSections = section.getDocument().getChildNodes(NodeType.SECTION, false);
            int sectionIndex = docSections.indexOf(section);

            indentAndAppendLine("[Section start] Section index: " + sectionIndex);
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Section node have been visited.
        /// </summary>
        public int visitSectionEnd(final Section section) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Section end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Body node is encountered in the document.
        /// </summary>
        public int visitBodyStart(final Body body) {
            int paragraphCount = body.getParagraphs().getCount();
            indentAndAppendLine("[Body start] Paragraphs: " + paragraphCount);
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Body node have been visited.
        /// </summary>
        public int visitBodyEnd(final Body body) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Body end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Paragraph node is encountered in the document.
        /// </summary>
        public int visitParagraphStart(final Paragraph paragraph) {
            indentAndAppendLine("[Paragraph start]");
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Paragraph node have been visited.
        /// </summary>
        public int visitParagraphEnd(final Paragraph paragraph) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Paragraph end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a SubDocument node is encountered in the document.
        /// </summary>
        public int visitSubDocument(final SubDocument subDocument) {
            indentAndAppendLine("[SubDocument]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mAcceptingNodeChildTree.append("|  ");
            }

            mAcceptingNodeChildTree.append(text + "\r\n");
        }

        private int mDocTraversalDepth;
        private final StringBuilder mAcceptingNodeChildTree;
    }
    //ExEnd

    private void testDocStructureToText(DocStructurePrinter visitor) {
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
    //ExSummary:Shows how to print the node structure of every table in a document.
    @Test //ExSkip
    public void tableToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        TableStructurePrinter visitor = new TableStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testTableToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Table nodes and their children.
    /// </summary>
    public static class TableStructurePrinter extends DocumentVisitor {
        public TableStructurePrinter() {
            mVisitedTables = new StringBuilder();
            mVisitorIsInsideTable = false;
        }

        public String getText() {
            return mVisitedTables.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// Runs that are not within tables are not recorded.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run) {
            if (mVisitorIsInsideTable) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Table is encountered in the document.
        /// </summary>
        public int visitTableStart(final Table table) {
            int rows = 0;
            int columns = 0;

            if (table.getRows().getCount() > 0) {
                rows = table.getRows().getCount();
                columns = table.getFirstRow().getCount();
            }

            indentAndAppendLine("[Table start] Size: " + rows + "x" + columns);
            mDocTraversalDepth++;
            mVisitorIsInsideTable = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Table node have been visited.
        /// </summary>
        public int visitTableEnd(final Table table) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Table end]");
            mVisitorIsInsideTable = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Row node is encountered in the document.
        /// </summary>
        public int visitRowStart(final Row row) {
            String rowContents = row.getText().replaceAll("\\u0007", ", ").replaceAll(", , ", "");
            int rowWidth = row.indexOf(row.getLastCell()) + 1;
            int rowIndex = row.getParentTable().indexOf(row);
            String rowStatusInTable = row.isFirstRow() && row.isLastRow() ? "only" : row.isFirstRow() ? "first" : row.isLastRow() ? "last" : "";
            if (!"".equals(rowStatusInTable)) {
                rowStatusInTable = MessageFormat.format(", the {0} row in this table,", rowStatusInTable);
            }

            indentAndAppendLine(MessageFormat.format("[Row start] Row #{0}{1} width {2}, \"{3}\"", ++rowIndex, rowStatusInTable, rowWidth, rowContents));
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Row node have been visited.
        /// </summary>
        public int visitRowEnd(final Row row) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Row end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Cell node is encountered in the document.
        /// </summary>
        public int visitCellStart(final Cell cell) {
            Row row = cell.getParentRow();
            Table table = row.getParentTable();
            String cellStatusInRow = cell.isFirstCell() && cell.isLastCell() ? "only" : cell.isFirstCell() ? "first" : cell.isLastCell() ? "last" : "";
            if (!"".equals(cellStatusInRow)) {
                cellStatusInRow = MessageFormat.format(", the {0} cell in this row", cellStatusInRow);
            }

            indentAndAppendLine(MessageFormat.format("[Cell start] Row {0}, Col {1}{2}", table.indexOf(row) + 1, row.indexOf(cell) + 1, cellStatusInRow));
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Cell node have been visited.
        /// </summary>
        public int visitCellEnd(final Cell cell) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Cell end]");
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into the current table's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mVisitedTables.append("|  ");
            }

            mVisitedTables.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideTable;
        private int mDocTraversalDepth;
        private final /*final*/ StringBuilder mVisitedTables;
    }
    //ExEnd

    private void testTableToText(TableStructurePrinter visitor) {
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
    //ExSummary:Shows how to print the node structure of every comment and comment range in a document.
    @Test //ExSkip
    public void commentsToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        CommentStructurePrinter visitor = new CommentStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testCommentsToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Comment/CommentRange nodes and their children.
    /// </summary>
    public static class CommentStructurePrinter extends DocumentVisitor {
        public CommentStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideComment = false;
        }

        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// A Run is only recorded if it is a child of a Comment or CommentRange node.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideComment) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeStart node is encountered in the document.
        /// </summary>
        public int visitCommentRangeStart(final CommentRangeStart commentRangeStart) {
            indentAndAppendLine("[Comment range start] ID: " + commentRangeStart.getId());
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a CommentRangeEnd node is encountered in the document.
        /// </summary>
        public int visitCommentRangeEnd(final CommentRangeEnd commentRangeEnd) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment range end]");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Comment node is encountered in the document.
        /// </summary>
        public int visitCommentStart(final Comment comment) {
            indentAndAppendLine(MessageFormat.format("[Comment start] For comment range ID {0}, By {1} on {2}", comment.getId(),
                    comment.getAuthor(), comment.getDateTime()));
            mDocTraversalDepth++;
            mVisitorIsInsideComment = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Comment node have been visited.
        /// </summary>
        public int visitCommentEnd(final Comment comment) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Comment end]");
            mVisitorIsInsideComment = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into a comment/comment range's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideComment;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testCommentsToText(CommentStructurePrinter visitor) {
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
    //ExSummary:Shows how to print the node structure of every field in a document.
    @Test //ExSkip
    public void fieldToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        FieldStructurePrinter visitor = new FieldStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testFieldToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Field nodes and their children.
    /// </summary>
    public static class FieldStructurePrinter extends DocumentVisitor {
        public FieldStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideField = false;
        }

        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideField) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldStart node is encountered in the document.
        /// </summary>
        public int visitFieldStart(final FieldStart fieldStart) {
            indentAndAppendLine("[Field start] FieldType: " + fieldStart.getFieldType());
            mDocTraversalDepth++;
            mVisitorIsInsideField = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldEnd node is encountered in the document.
        /// </summary>
        public int visitFieldEnd(final FieldEnd fieldEnd) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Field end]");
            mVisitorIsInsideField = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a FieldSeparator node is encountered in the document.
        /// </summary>
        public int visitFieldSeparator(final FieldSeparator fieldSeparator) {
            indentAndAppendLine("[FieldSeparator]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
        /// into the field's tree of child nodes.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideField;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testFieldToText(FieldStructurePrinter visitor) {
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
    //ExSummary:Shows how to print the node structure of every header and footer in a document.
    @Test //ExSkip
    public void headerFooterToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        HeaderFooterStructurePrinter visitor = new HeaderFooterStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());

        // An alternative way of accessing a document's header/footers section-by-section is by accessing the collection.
        HeaderFooter[] headerFooters = doc.getFirstSection().getHeadersFooters().toArray();
        Assert.assertEquals(3, headerFooters.length);
        testHeaderFooterToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered HeaderFooter nodes and their children.
    /// </summary>
    public static class HeaderFooterStructurePrinter extends DocumentVisitor {
        public HeaderFooterStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideHeaderFooter = false;
        }

        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideHeaderFooter) indentAndAppendLine("[Run] \"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a HeaderFooter node is encountered in the document.
        /// </summary>
        public int visitHeaderFooterStart(final HeaderFooter headerFooter) {
            indentAndAppendLine("[HeaderFooter start] HeaderFooterType: " + headerFooter.getHeaderFooterType());
            mDocTraversalDepth++;
            mVisitorIsInsideHeaderFooter = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a HeaderFooter node have been visited.
        /// </summary>
        public int visitHeaderFooterEnd(final HeaderFooter headerFooter) {
            mDocTraversalDepth--;
            indentAndAppendLine("[HeaderFooter end]");
            mVisitorIsInsideHeaderFooter = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideHeaderFooter;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testHeaderFooterToText(HeaderFooterStructurePrinter visitor) {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains(MessageFormat.format("[HeaderFooter start] HeaderFooterType: {0}", HeaderFooterType.HEADER_PRIMARY)));
        Assert.assertTrue(visitorText.contains("[HeaderFooter end]"));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[HeaderFooter start] HeaderFooterType: {0}", HeaderFooterType.HEADER_FIRST)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[HeaderFooter start] HeaderFooterType: {0}", HeaderFooterType.HEADER_EVEN)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[HeaderFooter start] HeaderFooterType: {0}", HeaderFooterType.FOOTER_PRIMARY)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[HeaderFooter start] HeaderFooterType: {0}", HeaderFooterType.FOOTER_FIRST)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[HeaderFooter start] HeaderFooterType: {0}", HeaderFooterType.FOOTER_EVEN)));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
    //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
    //ExSummary:Shows how to print the node structure of every editable range in a document.
    @Test //ExSkip
    public void editableRangeToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        EditableRangeStructurePrinter visitor = new EditableRangeStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testEditableRangeToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered EditableRange nodes and their children.
    /// </summary>
    public static class EditableRangeStructurePrinter extends DocumentVisitor {
        public EditableRangeStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideEditableRange = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            // We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes.
            if (mVisitorIsInsideEditableRange) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an EditableRange node is encountered in the document.
        /// </summary>
        public int visitEditableRangeStart(final EditableRangeStart editableRangeStart) {
            indentAndAppendLine("[EditableRange start] ID: " + editableRangeStart.getId() + " Owner: "
                    + editableRangeStart.getEditableRange().getSingleUser());
            mDocTraversalDepth++;
            mVisitorIsInsideEditableRange = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when the visiting of a EditableRange node is ended.
        /// </summary>
        public int visitEditableRangeEnd(final EditableRangeEnd editableRangeEnd) {
            mDocTraversalDepth--;
            indentAndAppendLine("[EditableRange end]");
            mVisitorIsInsideEditableRange = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideEditableRange;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testEditableRangeToText(EditableRangeStructurePrinter visitor) {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[EditableRange start]"));
        Assert.assertTrue(visitorText.contains("[EditableRange end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
    //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
    //ExFor:Footnote.Accept(DocumentVisitor)
    //ExSummary:Shows how to print the node structure of every footnote in a document.
    @Test //ExSkip
    public void footnoteToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        FootnoteStructurePrinter visitor = new FootnoteStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testFootnoteToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered Footnote nodes and their children.
    /// </summary>
    public static class FootnoteStructurePrinter extends DocumentVisitor {
        public FootnoteStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideFootnote = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Footnote node is encountered in the document.
        /// </summary>
        public int visitFootnoteStart(final Footnote footnote) {
            indentAndAppendLine("[Footnote start] Type: " + footnote.getFootnoteType());
            mDocTraversalDepth++;
            mVisitorIsInsideFootnote = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a Footnote node have been visited.
        /// </summary>
        public int visitFootnoteEnd(final Footnote footnote) {
            mDocTraversalDepth--;
            indentAndAppendLine("[Footnote end]");
            mVisitorIsInsideFootnote = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideFootnote) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideFootnote;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testFootnoteToText(FootnoteStructurePrinter visitor) {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains(MessageFormat.format("[Footnote start] Type: {0}", FootnoteType.FOOTNOTE)));
        Assert.assertTrue(visitorText.contains("[Footnote end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitOfficeMathEnd(Math.OfficeMath)
    //ExFor:DocumentVisitor.VisitOfficeMathStart(Math.OfficeMath)
    //ExFor:Math.MathObjectType
    //ExFor:Math.OfficeMath.Accept(DocumentVisitor)
    //ExFor:Math.OfficeMath.MathObjectType
    //ExSummary:Shows how to print the node structure of every office math node in a document.
    @Test //ExSkip
    public void officeMathToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        OfficeMathStructurePrinter visitor = new OfficeMathStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testOfficeMathToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered OfficeMath nodes and their children.
    /// </summary>
    public static class OfficeMathStructurePrinter extends DocumentVisitor {
        public OfficeMathStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideOfficeMath = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideOfficeMath) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an OfficeMath node is encountered in the document.
        /// </summary>
        public int visitOfficeMathStart(final OfficeMath officeMath) {
            indentAndAppendLine("[OfficeMath start] Math object type: " + officeMath.getMathObjectType());
            mDocTraversalDepth++;
            mVisitorIsInsideOfficeMath = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of an OfficeMath node have been visited.
        /// </summary>
        public int visitOfficeMathEnd(final OfficeMath officeMath) {
            mDocTraversalDepth--;
            indentAndAppendLine("[OfficeMath end]");
            mVisitorIsInsideOfficeMath = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideOfficeMath;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testOfficeMathToText(OfficeMathStructurePrinter visitor) {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.O_MATH_PARA)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.O_MATH)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.ARGUMENT)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.SUPERCRIPT)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.SUPERSCRIPT_PART)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.FRACTION)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.NUMERATOR)));
        Assert.assertTrue(visitorText.contains(MessageFormat.format("[OfficeMath start] Math object type: {0}", MathObjectType.DENOMINATOR)));
        Assert.assertTrue(visitorText.contains("[OfficeMath end]"));
        Assert.assertTrue(visitorText.contains("[Run]"));
    }

    //ExStart
    //ExFor:DocumentVisitor.VisitSmartTagEnd(Markup.SmartTag)
    //ExFor:DocumentVisitor.VisitSmartTagStart(Markup.SmartTag)
    //ExSummary:Shows how to print the node structure of every smart tag in a document.
    @Test //ExSkip
    public void smartTagToText() throws Exception {
        Document doc = new Document(getMyDir() + "Smart tags.doc");
        SmartTagStructurePrinter visitor = new SmartTagStructurePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testSmartTagToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered SmartTag nodes and their children.
    /// </summary>
    public static class SmartTagStructurePrinter extends DocumentVisitor {
        public SmartTagStructurePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideSmartTag = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideSmartTag) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a SmartTag node is encountered in the document.
        /// </summary>
        public int visitSmartTagStart(final SmartTag smartTag) {
            indentAndAppendLine("[SmartTag start] Name: " + smartTag.getElement());
            mDocTraversalDepth++;
            mVisitorIsInsideSmartTag = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a SmartTag node have been visited.
        /// </summary>
        public int visitSmartTagEnd(final SmartTag smartTag) {
            mDocTraversalDepth--;
            indentAndAppendLine("[SmartTag end]");
            mVisitorIsInsideSmartTag = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private boolean mVisitorIsInsideSmartTag;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testSmartTagToText(SmartTagStructurePrinter visitor) {
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
    //ExSummary:Shows how to print the node structure of every structured document tag in a document.
    @Test //ExSkip
    public void structuredDocumentTagToText() throws Exception {
        Document doc = new Document(getMyDir() + "DocumentVisitor-compatible features.docx");
        StructuredDocumentTagNodePrinter visitor = new StructuredDocumentTagNodePrinter();

        // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
        // and then traverses all the node's children in a depth-first manner.
        // The visitor can read and modify each visited node.
        doc.accept(visitor);

        System.out.println(visitor.getText());
        testStructuredDocumentTagToText(visitor); //ExSkip
    }

    /// <summary>
    /// Traverses a node's non-binary tree of child nodes.
    /// Creates a map in the form of a string of all encountered StructuredDocumentTag nodes and their children.
    /// </summary>
    public static class StructuredDocumentTagNodePrinter extends DocumentVisitor {
        public StructuredDocumentTagNodePrinter() {
            mBuilder = new StringBuilder();
            mVisitorIsInsideStructuredDocumentTag = false;
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            if (mVisitorIsInsideStructuredDocumentTag) {
                indentAndAppendLine("[Run] \"" + run.getText() + "\"");
            }

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a StructuredDocumentTag node is encountered in the document.
        /// </summary>
        public int visitStructuredDocumentTagStart(final StructuredDocumentTag sdt) {
            indentAndAppendLine("[StructuredDocumentTag start] Title: " + sdt.getTitle());
            mDocTraversalDepth++;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called after all the child nodes of a StructuredDocumentTag node have been visited.
        /// </summary>
        public int visitStructuredDocumentTagEnd(final StructuredDocumentTag sdt) {
            mDocTraversalDepth--;
            indentAndAppendLine("[StructuredDocumentTag end]");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
        /// </summary>
        /// <param name="text"></param>
        private void indentAndAppendLine(final String text) {
            for (int i = 0; i < mDocTraversalDepth; i++) {
                mBuilder.append("|  ");
            }

            mBuilder.append(text + "\r\n");
        }

        private final boolean mVisitorIsInsideStructuredDocumentTag;
        private int mDocTraversalDepth;
        private final StringBuilder mBuilder;
    }
    //ExEnd

    private void testStructuredDocumentTagToText(StructuredDocumentTagNodePrinter visitor) {
        String visitorText = visitor.getText();

        Assert.assertTrue(visitorText.contains("[StructuredDocumentTag start]"));
        Assert.assertTrue(visitorText.contains("[StructuredDocumentTag end]"));
    }
}
