package DocsExamples.Programming_with_Documents.Working_with_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.FindReplaceDirection;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.Bookmark;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.CompositeNode;
import com.aspose.words.NodeImporter;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.Section;
import com.aspose.words.Paragraph;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;


class CloneAndCombineDocuments extends DocsExamplesBase
{
    @Test
    public void cloningDocument() throws Exception
    {
        //ExStart:CloningDocument
        Document doc = new Document(getMyDir() + "Document.docx");

        Document clone = doc.deepClone();
        clone.save(getArtifactsDir() + "CloneAndCombineDocuments.CloningDocument.docx");
        //ExEnd:CloningDocument
    }

    @Test
    public void insertDocumentAtReplace() throws Exception
    {
        //ExStart:InsertDocumentAtReplace
        Document mainDoc = new Document(getMyDir() + "Document insertion 1.docx");

        // Set find and replace options.
        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setDirection(FindReplaceDirection.BACKWARD); 
            options.setReplacingCallback(new InsertDocumentAtReplaceHandler());
        }

        // Call the replace method.
        mainDoc.getRange().replaceInternal(new Regex("\\[MY_DOCUMENT\\]"), "", options);
        mainDoc.save(getArtifactsDir() + "CloneAndCombineDocuments.InsertDocumentAtReplace.docx");
        //ExEnd:InsertDocumentAtReplace
    }

    @Test
    public void insertDocumentAtBookmark() throws Exception
    {
        //ExStart:InsertDocumentAtBookmark         
        Document mainDoc = new Document(getMyDir() + "Document insertion 1.docx");
        Document subDoc = new Document(getMyDir() + "Document insertion 2.docx");

        Bookmark bookmark = mainDoc.getRange().getBookmarks().get("insertionPlace");
        insertDocument(bookmark.getBookmarkStart().getParentNode(), subDoc);
        
        mainDoc.save(getArtifactsDir() + "CloneAndCombineDocuments.InsertDocumentAtBookmark.docx");
        //ExEnd:InsertDocumentAtBookmark
    }

    @Test
    public void insertDocumentAtMailMerge() throws Exception
    {
        //ExStart:InsertDocumentAtMailMerge   
        Document mainDoc = new Document(getMyDir() + "Document insertion 1.docx");

        mainDoc.getMailMerge().setFieldMergingCallback(new InsertDocumentAtMailMergeHandler());
        // The main document has a merge field in it called "Document_1".
        // The corresponding data for this field contains a fully qualified path to the document.
        // That should be inserted to this field.
        mainDoc.getMailMerge().execute(new String[] { "Document_1" }, new Object[] { getMyDir() + "Document insertion 2.docx" });

        mainDoc.save(getArtifactsDir() + "CloneAndCombineDocuments.InsertDocumentAtMailMerge.doc");
        //ExEnd:InsertDocumentAtMailMerge
    }

    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// Section breaks and section formatting of the inserted document are ignored.
    /// </summary>
    /// <param name="insertionDestination">Node in the destination document after which the content
    /// Should be inserted. This node should be a block level node (paragraph or table).</param>
    /// <param name="docToInsert">The document to insert.</param>
    //ExStart:InsertDocument
    private static void insertDocument(Node insertionDestination, Document docToInsert)
    {
        if (insertionDestination.getNodeType() == NodeType.PARAGRAPH || insertionDestination.getNodeType() == NodeType.TABLE)
        {
            CompositeNode destinationParent = insertionDestination.getParentNode();

            NodeImporter importer =
                new NodeImporter(docToInsert, insertionDestination.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

            // Loop through all block-level nodes in the section's body,
            // then clone and insert every node that is not the last empty paragraph of a section.
            for (Section srcSection : docToInsert.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
            for (Node srcNode : (Iterable<Node>) srcSection.getBody())
            {
                if (srcNode.getNodeType() == NodeType.PARAGRAPH)
                {
                    Paragraph para = (Paragraph)srcNode;
                    if (para.isEndOfSection() && !para.hasChildNodes())
                        continue;
                }

                Node newNode = importer.importNode(srcNode, true);

                destinationParent.insertAfter(newNode, insertionDestination);
                insertionDestination = newNode;
            }
        }
        else
        {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }
    }
    //ExEnd:InsertDocument

    //ExStart:InsertDocumentWithSectionFormatting
    /// <summary>
    /// Inserts content of the external document after the specified node.
    /// </summary>
    /// <param name="insertAfterNode">Node in the destination document after which the content
    /// Should be inserted. This node should be a block level node (paragraph or table).</param>
    /// <param name="srcDoc">The document to insert.</param>
    private void insertDocumentWithSectionFormatting(Node insertAfterNode, Document srcDoc)
    {
        if (insertAfterNode.getNodeType() != NodeType.PARAGRAPH &&
            insertAfterNode.getNodeType() != NodeType.TABLE)
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");

        Document dstDoc = (Document) insertAfterNode.getDocument();
        // To retain section formatting, split the current section into two at the marker node and then import the content
        // from srcDoc as whole sections. The section of the node to which the insert marker node belongs.
        Section currentSection = (Section) insertAfterNode.getAncestor(NodeType.SECTION);

        // Don't clone the content inside the section, we just want the properties of the section retained.
        Section cloneSection = (Section) currentSection.deepClone(false);

        // However, make sure the clone section has a body but no empty first paragraph.
        cloneSection.ensureMinimum();
        cloneSection.getBody().getFirstParagraph().remove();

        insertAfterNode.getDocument().insertAfter(cloneSection, currentSection);

        // Append all nodes after the marker node to the new section. This will split the content at the section level at.
        // The marker so the sections from the other document can be inserted directly.
        Node currentNode = insertAfterNode.getNextSibling();
        while (currentNode != null)
        {
            Node nextNode = currentNode.getNextSibling();
            cloneSection.getBody().appendChild(currentNode);
            currentNode = nextNode;
        }

        // This object will be translating styles and lists during the import.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.USE_DESTINATION_STYLES);

        for (Section srcSection : (Iterable<Section>) srcDoc.getSections())
        {
            Node newNode = importer.importNode(srcSection, true);

            dstDoc.insertAfter(newNode, currentSection);
            currentSection = (Section) newNode;
        }
    }
    //ExEnd:InsertDocumentWithSectionFormatting

    //ExStart:InsertDocumentAtMailMergeHandler
    private static class InsertDocumentAtMailMergeHandler implements IFieldMergingCallback
    {
        // This handler makes special processing for the "Document_1" field.
        // The field value contains the path to load the document. 
        // We load the document and insert it into the current merge field.
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs args) throws Exception
        {
            if ("Document_1".equals(args.getDocumentFieldName()))
            {
                // Use document builder to navigate to the merge field with the specified name.
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                builder.moveToMergeField(args.getDocumentFieldName());

                // The name of the document to load and insert is stored in the field value.
                Document subDoc = new Document((String)args.getFieldValue());
                
                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now, and you probably want to delete it.
                if (!builder.getCurrentParagraph().hasChildNodes())
                    builder.getCurrentParagraph().remove();

                // Indicate to the mail merge engine that we have inserted what we wanted.
                args.setText(null);
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }
    }
    //ExEnd:InsertDocumentAtMailMergeHandler

    //ExStart:InsertDocumentAtMailMergeBlobHandler
    private static class InsertDocumentAtMailMergeBlobHandler implements IFieldMergingCallback
    {
        /// <summary>
        /// This handler makes special processing for the "Document_1" field.
        /// The field value contains the path to load the document.
        /// We load the document and insert it into the current merge field.
        /// </summary>
        public void /*IFieldMergingCallback.*/fieldMerging(FieldMergingArgs e) throws Exception
        {
            if ("Document_1".equals(e.getDocumentFieldName()))
            {
                DocumentBuilder builder = new DocumentBuilder(e.getDocument());
                builder.moveToMergeField(e.getDocumentFieldName());

                MemoryStream stream = new MemoryStream((byte[]) e.getFieldValue());
                Document subDoc = new Document(stream);

                insertDocument(builder.getCurrentParagraph(), subDoc);

                // The paragraph that contained the merge field might be empty now, and you probably want to delete it.
                if (!builder.getCurrentParagraph().hasChildNodes())
                    builder.getCurrentParagraph().remove();

                e.setText(null);
            }
        }

        public void /*IFieldMergingCallback.*/imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }
    }
    //ExEnd:InsertDocumentAtMailMergeBlobHandler
    
    //ExStart:InsertDocumentAtReplaceHandler
    private static class InsertDocumentAtReplaceHandler implements IReplacingCallback
    {
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs args) throws Exception
        {
            Document subDoc = new Document(getMyDir() + "Document insertion 2.docx");

            // Insert a document after the paragraph, containing the match text.
            Paragraph para = (Paragraph)args.getMatchNode().getParentNode();
            insertDocument(para, subDoc);
            
            // Remove the paragraph with the match text.
            para.remove();
            return ReplaceAction.SKIP;
        }
    }
    //ExEnd:InsertDocumentAtReplaceHandler
}
