package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.util.regex.Pattern;

public class ExInsertDocument extends ApiExampleBase {
    //ExStart
    //ExFor:Paragraph.IsEndOfSection
    //ExFor:NodeImporter
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
    //ExFor:NodeImporter.ImportNode(Node, Boolean)
    //ExSummary:This is a method that inserts contents of one document at a specified location in another document.

    /**
     * Inserts content of the external document after the specified node.
     * Section breaks and section formatting of the inserted document are ignored.
     *
     * @param insertAfterNode Node in the destination document after which the content
     *                        should be inserted. This node should be a block level node (paragraph or table).
     * @param srcDoc          The document to insert.
     */
    public static void insertDocument(Node insertAfterNode, final Document srcDoc) {
        // Make sure that the node is either a paragraph or table.
        if ((insertAfterNode.getNodeType() != NodeType.PARAGRAPH) & (insertAfterNode.getNodeType() != NodeType.TABLE)) {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }

        // We will be inserting into the parent of the destination paragraph
        CompositeNode dstStory = insertAfterNode.getParentNode();

        // This object will be translating styles and lists during the import
        NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // Loop through all sections in the source document
        for (Section srcSection : srcDoc.getSections()) {
            // Loop through all block level nodes (paragraphs and tables) in the body of the section
            for (Node srcNode : srcSection.getBody()) {
                // Let's skip the node if it is a last empty paragraph in a section
                if (srcNode.getNodeType() == (NodeType.PARAGRAPH)) {
                    Paragraph para = (Paragraph) srcNode;
                    if (para.isEndOfSection() && !para.hasChildNodes()) {
                        continue;
                    }
                }

                // This creates a clone of the node, suitable for insertion into the destination document
                Node newNode = importer.importNode(srcNode, true);

                // Insert new node after the reference node
                dstStory.insertAfter(newNode, insertAfterNode);
                insertAfterNode = newNode;
            }
        }
    }
    //ExEnd

    @Test
    public void insertAtBookmark() throws Exception {
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");
        Document subDoc = new Document(getMyDir() + "Document.docx");

        Bookmark bookmark = mainDoc.getRange().getBookmarks().get("insertionPlace");
        insertDocument(bookmark.getBookmarkStart().getParentNode(), subDoc);

        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertAtBookmark.doc");
    }

    //ExStart
    //ExFor:CompositeNode.HasChildNodes
    //ExSummary:Demonstrates how to use the InsertDocument method to insert a document into a merge field during mail merge.
    @Test //ExSkip
    public void insertAtMailMerge() throws Exception {
        // Open the main document
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");

        // Add a handler to MergeField event
        mainDoc.getMailMerge().setFieldMergingCallback(new InsertDocumentAtMailMergeHandler());

        // The main document has a merge field in it called "Document_1"
        // The corresponding data for this field contains fully qualified path to the document
        // that should be inserted to this field
        mainDoc.getMailMerge().execute(new String[]{"Document_1"}, new String[]{getMyDir() + "Document.docx"});

        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertAtMailMerge.docx");
    }

    private class InsertDocumentAtMailMergeHandler implements IFieldMergingCallback {
        /**
         * This handler makes special processing for the "Document_1" field.
         * The field value contains the path to load the document.
         * We load the document and insert it into the current merge field.
         */
        public void fieldMerging(final FieldMergingArgs args) throws Exception {
            if ("Document_1".equals(args.getDocumentFieldName())) {
                // Use document builder to navigate to the merge field with the specified name
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());

                // The name of the document to load and insert is stored in the field value
                Document subDoc = new Document((String) args.getFieldValue());

                // Insert the document
                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now and you probably want to delete it
                if (!builder.getCurrentParagraph().hasChildNodes()) {
                    builder.getCurrentParagraph().remove();
                }

                // Indicate to the mail merge engine that we have inserted what we wanted
                args.setText(null);
            }
        }

        public void imageFieldMerging(final ImageFieldMergingArgs args) {
            // Do nothing
        }
    }
    //ExEnd

    private class InsertDocumentAtMailMergeBlobHandler implements IFieldMergingCallback {
        /**
         * This handler makes special processing for the "Document_1" field.
         * The field value contains the path to load the document.
         * We load the document and insert it into the current merge field.
         */
        public void fieldMerging(final FieldMergingArgs args) throws Exception {
            if ("Document_1".equals(args.getDocumentFieldName())) {
                // Use document builder to navigate to the merge field with the specified name
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());

                // Load the document from the blob field
                ByteArrayInputStream inStream = new ByteArrayInputStream((byte[]) args.getFieldValue());
                Document subDoc = new Document(inStream);
                inStream.close();

                // Insert the document
                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now and you probably want to delete it
                if (!builder.getCurrentParagraph().hasChildNodes()) {
                    builder.getCurrentParagraph().remove();
                }

                // Indicate to the mail merge engine that we have inserted what we wanted
                args.setText(null);
            }
        }

        public void imageFieldMerging(final ImageFieldMergingArgs args) {
            // Do nothing
        }
    }


    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:IReplacingCallback
    //ExFor:ReplaceAction
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:ReplacingArgs.MatchNode
    //ExFor:FindReplaceDirection
    //ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
    @Test //ExSkip
    public void insertDocumentAtReplace() throws Exception {
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setDirection(FindReplaceDirection.BACKWARD);
        options.setReplacingCallback(new InsertDocumentAtReplaceHandler());

        mainDoc.getRange().replace(Pattern.compile("\\[MY_DOCUMENT\\]"), "", options);
        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertDocumentAtReplace.doc");
    }

    private class InsertDocumentAtReplaceHandler implements IReplacingCallback {
        public int replacing(final ReplacingArgs args) throws Exception {
            Document subDoc = new Document(getMyDir() + "Document.docx");

            // Insert a document after the paragraph, containing the match text
            Paragraph para = (Paragraph) args.getMatchNode().getParentNode();
            insertDocument(para, subDoc);

            // Remove the paragraph with the match text
            para.remove();

            return ReplaceAction.SKIP;
        }
    }
    //ExEnd
}

