package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayInputStream;
import java.util.regex.Pattern;

public class InsertDocumentIntoAnotherDocument {

    public static final String dataDir = Utils.getSharedDataDir(InsertDocumentIntoAnotherDocument.class) + "InsertDocumentIntoAnother/";

    public static void main(String[] args) throws Exception {
        //Insert a Document at a Bookmark
        insertADocumentAtABookmark();

        //Insert a Document During Mail Merge
        insertDocumentAtMailMerge();

        //Insert a Document During Replace
        insertDocumentAtReplace();
    }

    public static void insertADocumentAtABookmark() throws Exception {
        //ExStart:insertADocumentAtABookmark
        Document mainDoc = new Document(dataDir + "InsertDocument1.doc");
        Document subDoc = new Document(dataDir + "InsertDocument2.doc");

        Bookmark bookmark = mainDoc.getRange().getBookmarks().get("insertionPlace");
        insertDocument(bookmark.getBookmarkStart().getParentNode(), subDoc);

        mainDoc.save(dataDir + "InsertDocumentAtBookmark_out.doc");
        //ExEnd:insertADocumentAtABookmark
    }

    public static void insertDocumentAtMailMerge() throws Exception {
        //ExStart:insertDocumentAtMailMerge
        // Open the main document.
        Document mainDoc = new Document(dataDir + "InsertDocument1.doc");

        // Add a handler to MergeField event
        mainDoc.getMailMerge().setFieldMergingCallback(new InsertDocumentAtMailMergeHandler());

        // The main document has a merge field in it called "Document_1".
        // The corresponding data for this field contains fully qualified path to the document
        // that should be inserted to this field.
        mainDoc.getMailMerge().execute(new String[]{"Document_1"}, new String[]{dataDir + "InsertDocument2.doc"});

        mainDoc.save(dataDir + "InsertDocument.InsertAtMailMerge.docx");
        //ExEnd:insertDocumentAtMailMerge
    }

    public static void insertDocumentAtReplace() throws Exception {
        //ExStart: insertDocumentAtReplace
        Document mainDoc = new Document(dataDir + "InsertDocument1.doc");

        // Set find and replace options.
        FindReplaceOptions options = new FindReplaceOptions();
        options.setDirection(FindReplaceDirection.BACKWARD);
        options.setReplacingCallback(new InsertDocumentAtReplaceHandler());

        mainDoc.getRange().replace(Pattern.compile("\\[MY_DOCUMENT\\]"), "", options);
        mainDoc.save(dataDir + "InsertDocument.InsertDocumentAtReplace.docx");
        //ExEnd: insertDocumentAtReplace
    }

    //ExStart:InsertDocumentAtReplaceHandler
    private static class InsertDocumentAtReplaceHandler implements IReplacingCallback {

        public int replacing(ReplacingArgs e) throws Exception {
            Document subDoc = new Document(dataDir + "InsertDocument2.doc");

            // Insert a document after the paragraph, containing the match text.
            Paragraph para = (Paragraph) e.getMatchNode().getParentNode();
            insertDocument(para, subDoc);

            // Remove the paragraph with the match text.
            para.remove();

            return ReplaceAction.SKIP;
        }
    }
    //ExEnd:InsertDocumentAtReplaceHandler

    //ExStart:InsertDocumentAtMailMergeHandler
    private static class InsertDocumentAtMailMergeHandler implements IFieldMergingCallback {
        /**
         * This handler makes special processing for the "Document_1" field. The
         * field value contains the path to load the document. We load the
         * document and insert it into the current merge field.
         */
        public void fieldMerging(FieldMergingArgs e) throws Exception {
            if ("Document_1".equals(e.getDocumentFieldName())) {
                // Use document builder to navigate to the merge field with the specified name.
                DocumentBuilder builder = new DocumentBuilder(e.getDocument());
                builder.moveToMergeField(e.getDocumentFieldName());

                // The name of the document to load and insert is stored in the field value.
                Document subDoc = new Document((String) e.getFieldValue());

                // Insert the document.
                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now and you probably want to delete it.
                if (!builder.getCurrentParagraph().hasChildNodes())
                    builder.getCurrentParagraph().remove();

                // Indicate to the mail merge engine that we have inserted what we wanted.
                e.setText(null);
            }
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
            // Do nothing.
        }
    }
    //ExEnd:InsertDocumentAtMailMergeHandler

    //ExStart:InsertDocumentAtMailMergeBlobHandler
    //Load a document from a BLOB database field
    private class InsertDocumentAtMailMergeBlobHandler implements IFieldMergingCallback {
        /**
         * This handler makes special processing for the "Document_1" field. The
         * field value contains the path to load the document. We load the
         * document and insert it into the current merge field.
         */
        public void fieldMerging(FieldMergingArgs e) throws Exception {
            if ("Document_1".equals(e.getDocumentFieldName())) {
                // Use document builder to navigate to the merge field with the specified name.
                DocumentBuilder builder = new DocumentBuilder(e.getDocument());
                builder.moveToMergeField(e.getDocumentFieldName());

                // Load the document from the blob field.
                ByteArrayInputStream inStream = new ByteArrayInputStream((byte[]) e.getFieldValue());
                Document subDoc = new Document(inStream);
                inStream.close();

                // Insert the document.
                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now and you probably want to delete it.
                if (!builder.getCurrentParagraph().hasChildNodes())
                    builder.getCurrentParagraph().remove();

                // Indicate to the mail merge engine that we have inserted what we wanted.
                e.setText(null);
            }
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
            // Do nothing.
        }
    }
//ExEnd:InsertDocumentAtMailMergeBlobHandler

//ExStart:insertDocument
    /**
     * Inserts content of the external document after the specified node.
     * Section breaks and section formatting of the inserted document are
     * ignored.
     *
     * @param insertionDestination Node in the destination document after which the content
     *                        should be inserted. This node should be a block level node
     *                        (paragraph or table).
     * @param docToInsert     The document to insert.
     */
    public static void insertDocument(Node insertionDestination, Document docToInsert)
    {
        // Make sure that the node is either a paragraph or table.
        if (((insertionDestination.getNodeType()) == (NodeType.PARAGRAPH)) || ((insertionDestination.getNodeType()) == (NodeType.TABLE)))
        {
            // We will be inserting into the parent of the destination paragraph.
            CompositeNode dstStory = insertionDestination.getParentNode();

            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // Loop through all block level nodes in the body of the section
            for (Section srcSection : docToInsert.getSections().toArray())
                for (Node srcNode : srcSection.getBody())
                {
                    // Let's skip the node if it is a last empty paragraph in a section
                    if (((srcNode.getNodeType()) == (NodeType.PARAGRAPH)))
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document.
                    Node newNode = importer.importNode(srcNode, true);

                    // Insert new node after the reference node.
                    dstStory.insertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
        }
        else
        {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd:insertDocument

    //ExStart:insertDocumentWithSectionFormatting

    /**
     * Inserts content of the external document after the specified node.
     *
     * @param insertAfterNode Node in the destination document after which the content
     *                        should be inserted. This node should be a block level node
     *                        (paragraph or table).
     * @param srcDoc          The document to insert.
     */
    public static void insertDocumentWithSectionFormatting(Node insertAfterNode, Document srcDoc) throws Exception {
        // Make sure that the node is either a pargraph or table.
        if ((insertAfterNode.getNodeType() != NodeType.PARAGRAPH) & (insertAfterNode.getNodeType() != NodeType.TABLE))
            throw new Exception("The destination node should be either a paragraph or table.");

        // Document to insert srcDoc into.
        Document dstDoc = (Document) insertAfterNode.getDocument();

        // To retain section formatting, split the current section into two at the marker node and then import the content from srcDoc as whole sections.
        // The section of the node which the insert marker node belongs to

        Section currentSection = (Section) insertAfterNode.getAncestor(NodeType.SECTION);

        // Don't clone the content inside the section, we just want the properties of the section retained.
        Section cloneSection = (Section) currentSection.deepClone(false);

        // However make sure the clone section has a body, but no empty first paragraph.
        cloneSection.ensureMinimum();

        cloneSection.getBody().getFirstParagraph().remove();

        // Insert the cloned section into the document after the original section.
        insertAfterNode.getDocument().insertAfter(cloneSection, currentSection);

        // Append all nodes after the marker node to the new section. This will split the content at the section level at
        // the marker so the sections from the other document can be inserted directly.
        Node currentNode = insertAfterNode.getNextSibling();
        while (currentNode != null) {
            Node nextNode = currentNode.getNextSibling();
            cloneSection.getBody().appendChild(currentNode);
            currentNode = nextNode;
        }

        // This object will be translating styles and lists during the import.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        // Loop through all sections in the source document.
        for (Section srcSection : srcDoc.getSections()) {
            Node newNode = importer.importNode(srcSection, true);
            // Append each section to the destination document. Start by inserting it after the split section.
            dstDoc.insertAfter(newNode, currentSection);
            currentSection = (Section) newNode;
        }
    }
    //ExEnd:insertDocumentWithSectionFormatting
}
