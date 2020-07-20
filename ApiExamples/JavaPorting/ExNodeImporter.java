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
import com.aspose.words.Bookmark;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.CompositeNode;
import com.aspose.words.NodeImporter;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Section;
import com.aspose.words.Paragraph;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImageFieldMergingArgs;


@Test
public class ExNodeImporter extends ApiExampleBase
{
    //ExStart
    //ExFor:Paragraph.IsEndOfSection
    //ExFor:NodeImporter
    //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
    //ExFor:NodeImporter.ImportNode(Node, Boolean)
    //ExSummary:Shows how to insert the contents of one document to a bookmark in another document.
    @Test
    public void insertAtBookmark() throws Exception
    {
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");
        Document docToInsert = new Document(getMyDir() + "Document.docx");

        Bookmark bookmark = mainDoc.getRange().getBookmarks().get("insertionPlace");
        insertDocument(bookmark.getBookmarkStart().getParentNode(), docToInsert);

        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertAtBookmark.docx");
        testInsertAtBookmark(new Document(getArtifactsDir() + "InsertDocument.InsertAtBookmark.docx")); //ExSkip
    }

    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    static void insertDocument(Node insertionDestination, Document docToInsert)
    {
        // Make sure that the node is either a paragraph or table
        if (((insertionDestination.getNodeType()) == (NodeType.PARAGRAPH)) || ((insertionDestination.getNodeType()) == (NodeType.TABLE)))
        {
            // We will be inserting into the parent of the destination paragraph
            CompositeNode dstStory = insertionDestination.getParentNode();

            // This object will be translating styles and lists during the import
            NodeImporter importer =
                new NodeImporter(docToInsert, insertionDestination.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // Loop through all block level nodes in the body of the section
            for (Section srcSection : docToInsert.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
                for (Node srcNode : (Iterable<Node>) srcSection.getBody())
                {
                    // Let's skip the node if it is a last empty paragraph in a section
                    if (((srcNode.getNodeType()) == (NodeType.PARAGRAPH)))
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document
                    Node newNode = importer.importNode(srcNode, true);

                    // Insert new node after the reference node
                    dstStory.insertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
        }
        else
        {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd

    private void testInsertAtBookmark(Document doc)
    {
        Assert.assertEquals("1) At text that can be identified by regex:\r[MY_DOCUMENT]\r" +
                        "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                        "3) At a bookmark:\r\rHello World!", msString.trim(doc.getFirstSection().getBody().getText()));
    }

    @Test
    public void insertAtMailMerge() throws Exception
    {
        // Open the main document
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");

        // Add a handler to MergeField event
        mainDoc.getMailMerge().setFieldMergingCallback(new InsertDocumentAtMailMergeHandler());

        // The main document has a merge field in it called "Document_1"
        // The corresponding data for this field contains fully qualified path to the document
        // that should be inserted to this field
        mainDoc.getMailMerge().execute(new String[] { "Document_1" }, new Object[] { getMyDir() + "Document.docx" });

        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertAtMailMerge.docx");
        testInsertAtMailMerge(new Document(getArtifactsDir() + "InsertDocument.InsertAtMailMerge.docx")); //ExSkip
    }

    private static class InsertDocumentAtMailMergeHandler implements IFieldMergingCallback
    {
        /// <summary>
        /// This handler makes special processing for the "Document_1" field.
        /// The field value contains the path to load the document. 
        /// We load the document and insert it into the current merge field.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            if ("Document_1".equals(args.getDocumentFieldName()))
            {
                // Use document builder to navigate to the merge field with the specified name
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());

                // The name of the document to load and insert is stored in the field value
                Document subDoc = new Document((String)args.getFieldValue());

                // Insert the document
                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now and you probably want to delete it
                if (!builder.getCurrentParagraph().hasChildNodes())
                    builder.getCurrentParagraph().remove();

                // Indicate to the mail merge engine that we have inserted what we wanted
                args.setText(null);
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing
        }
    }

    private void testInsertAtMailMerge(Document doc)
    {
        Assert.assertEquals("1) At text that can be identified by regex:\r[MY_DOCUMENT]\r" +
                        "2) At a MERGEFIELD:\rHello World!\r" +
                        "3) At a bookmark:", msString.trim(doc.getFirstSection().getBody().getText()));
    }
}
