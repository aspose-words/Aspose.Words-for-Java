/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
  
package quickstart.workingwithnodes.java;
import com.aspose.words.*;
import org.testng.annotations.Test;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;
import java.util.List;
import java.util.regex.Pattern;
import com.aspose.words.Bookmark;
import com.aspose.words.BookmarkEnd;
import com.aspose.words.BookmarkStart;
import com.aspose.words.CompositeNode;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.NodeImporter;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
public class WorkingWithNodes {
        Document myDocument=null;
        DocumentBuilder myDocumentBuilder=null;
        @Test
        public void test() {
            //Variables declaration
            BookmarkEnd bookmarkEnd =null;
            BookmarkStart bookmarkStart=null;
            ArrayList extractedNodesInclusive = null;
            Document dstDoc =null;
            InputStream templateStream  = null;
            File templateFile =null;
            Bookmark sourceBmark=null;
            Bookmark destBmark =null;
            try {
                License license = new License();
                license.setLicense("C:\\X\\awuex\\Licenses\\Aspose.Words.Java.lic");
        String dataDir = "src/quickstart/workingwithnodes/data/";
                //Your word document location
                templateFile = new File(dataDir + "input.docx");
                templateStream  = new FileInputStream(templateFile);
                //document & documentbuilder
                myDocument = new Document(templateStream);
                myDocumentBuilder = new DocumentBuilder(myDocument);
                Hashtable<String,String> names = new Hashtable<String,String>();
                names.put("pierre", "name1");
                names.put("jean", "name2");
                names.put("paul", "name3");
                List list = Arrays.asList(myDocument.getChildNodes(NodeType.ANY, true).toArray());
                Enumeration<String> enumKey = names.keys();
                while(enumKey.hasMoreElements()) {
                    String key = enumKey.nextElement();
                    String val = names.get(key);
                    //We get the handle on the existing bookmarks
                    sourceBmark= myDocument.getRange().getBookmarks().get(key);
                    destBmark = myDocument.getRange().getBookmarks().get(val);
                    //Getting bookmarkStart and BookmarkEnd
                    bookmarkStart=sourceBmark.getBookmarkStart();
                    bookmarkEnd = sourceBmark.getBookmarkEnd();
                    //Extract the nodes contained inside the bookmark
                    extractedNodesInclusive = extractContent(bookmarkStart,bookmarkEnd, false);
                    //We generate a temporary aspose document
                    dstDoc= generateDocument(myDocument,extractedNodesInclusive);
                    //We erase the line of text in the destination bookmark ("destination")
                    myDocument.getRange().replace(val, "", false, false);
                    //Insertion of the temporary doc inside the bookmark destination
                    insertDocumentAtBookmark(destBmark.getName(),myDocument, dstDoc);
                }
                myDocument.save(dataDir + "result2.docx");
                sourceBmark= myDocument.getRange().getBookmarks().get("line");
                bookmarkStart=sourceBmark.getBookmarkStart();
                bookmarkEnd = sourceBmark.getBookmarkEnd();
                extractedNodesInclusive = extractContent(bookmarkStart,bookmarkEnd, false);
                dstDoc= generateDocument(myDocument,extractedNodesInclusive);
                /**Process With static bookmark**/
                /** Copy of the line with a static bookmark **/
                destBmark = myDocument.getRange().getBookmarks().get("copy_side");
                myDocument.getRange().replace("copy_side", "", false, false);
                //Insertion of the temporary doc inside the bookmark destination
                insertDocumentAtBookmark(destBmark.getName(), myDocument, dstDoc);
                /**End static bookmark process**/
                /**Process With dynamic bookmark**/
                /** Copy of the line with a dymanic bookmark **/
                //We move the cursor after the first destination bookmark
                myDocumentBuilder.moveToBookmark("copy_side", false, true);
                //Creation of a second bookmark name destination2
                myDocumentBuilder.startBookmark("line_copy");
                myDocumentBuilder.writeln("");
                myDocumentBuilder.endBookmark("line_copy");
                //Get the handle on the destination 2 bookmark
                Bookmark des2tBmark = myDocument.getRange().getBookmarks().get("line_copy");
                myDocument.save(dataDir + "result3.docx");
                //We insert the same temporary document inside the second bookmark
               // insertDocument(des2tBmark.getBookmarkStart().getParentNode(), dstDoc);
                insertDocumentAtBookmark(des2tBmark.getName(), myDocument, dstDoc);
                /**End dynamic bookmark process**/
                Bookmark bookmark4 = myDocument.getRange().getBookmarks().get("p1");
                extractedNodesInclusive = extractContent(bookmark4.getBookmarkStart(),bookmark4.getBookmarkEnd(), false);
                dstDoc= generateDocument(myDocument,extractedNodesInclusive);
                Bookmark dstBookmark = myDocument.getRange().getBookmarks().get("p1_copy");
                myDocument.getRange().replace("p1_copy", "", false, false);
                insertDocumentAtBookmark(dstBookmark.getName(), myDocument, dstDoc);
                //Save the word doc to your destination
                myDocument.save(dataDir + "result.docx");
                /**
                 * With the getParentNode().remove(), the paragraphs from the main doc are getting removed from the .doc!
                 *
                 */
                templateFile = new File(dataDir + "input_with_replace.docx");
                templateStream  = new FileInputStream(templateFile);
                //document & documentbuilder
                myDocument = new Document(templateStream);
                myDocumentBuilder = new DocumentBuilder(myDocument);
                //We create the first temp doc
                Bookmark bookmark1 = myDocument.getRange().getBookmarks().get("templating");
                extractedNodesInclusive = extractContent(bookmark1.getBookmarkStart(),bookmark1.getBookmarkEnd(), false);
                dstDoc= generateDocument(myDocument,extractedNodesInclusive);
                String startTextToRemove = "{%Text to remove}";
                String endTextToRemove = "{%End of text to remove}";
                //We replace the text that need to be removed with an empty string
                dstDoc.getRange().replace(startTextToRemove,"", true, false);
                dstDoc.getRange().replace(endTextToRemove,"", true, false);
                //Move after the first bookmark and creation of a new bookmark
                myDocumentBuilder.moveToBookmark("templating", false, true);
                myDocumentBuilder.startBookmark("templating_copy");
             //   myDocumentBuilder.writeln("");
                myDocumentBuilder.endBookmark("templating_copy");
                //Get the bookmark and insert the document inside
                Bookmark bookmark2 = myDocument.getRange().getBookmarks().get("templating_copy");
                insertDocumentAtBookmark(bookmark2.getName(), myDocument, dstDoc);
                //Move after the first bookmark and creation of a new bookmark
                myDocumentBuilder.moveToBookmark("templating_copy", false, true);
                myDocumentBuilder.startBookmark("templating_copy2");
              //  myDocumentBuilder.writeln("");
                myDocumentBuilder.endBookmark("templating_copy2");
                //we recreate a new template doc (not needed in this example but considered that the other doc could have been modified)
                extractedNodesInclusive = extractContent(bookmark1.getBookmarkStart(),bookmark1.getBookmarkEnd(), false);
                dstDoc= generateDocument(myDocument,extractedNodesInclusive);
                //We replace the text that need to be removed with an empty string
                dstDoc.getRange().replace(startTextToRemove,"", true, false);
                dstDoc.getRange().replace(endTextToRemove,"", true, false);
                //Get the bookmark and insert the document inside
                Bookmark bookmark3 = myDocument.getRange().getBookmarks().get("templating_copy2");
                insertDocumentAtBookmark(bookmark3.getName(), myDocument, dstDoc);
                //Remove the first bookmark
                bookmark1.setText("");
                bookmark1.remove();
                //Save the word doc to your destination
                myDocument.save(dataDir + "result test.docx");
            } catch (Exception e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }finally{
                System.out.print("Process ended");
            }
        }
    public static void insertDocumentAtBookmark(String bookmarkName, Document dstDoc, Document srcDoc)  throws Exception
    {
        //Create DocumentBuilder
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        //Move cursor to bookmark and insert paragraph break
        builder.moveToBookmark(bookmarkName);
        builder.writeln();
        // If current paragraph is a list item, we should clear its formating.
        if(builder.getCurrentParagraph().isListItem())
            builder.getListFormat().removeNumbers();
        //Content of srcdoc will be inserted after this node
        Node insertAfterNode = builder.getCurrentParagraph().getPreviousSibling();
        //Content of first paragraph of srcDoc will be apended to this parafraph
        Paragraph insertAfterParagraph = null;
        if(insertAfterNode.getNodeType() == NodeType.PARAGRAPH)
            insertAfterParagraph = (Paragraph)insertAfterNode;
        //Content of last paragraph of srcDoc will be apended to this parafraph
        Paragraph insertBeforeParagraph = builder.getCurrentParagraph();
        //We will be inserting into the parent of the destination paragraph.
        CompositeNode dstStory = insertAfterNode.getParentNode();
        //Remove empty paragraphs from the end of document
        while (!((CompositeNode)srcDoc.getLastSection().getBody().getLastChild()).hasChildNodes())
        {
            srcDoc.getLastSection().getBody().getLastParagraph().remove();
            if (srcDoc.getLastSection().getBody().getLastChild() == null)
                break;
        }
        //Loop through all sections in the source document.
        for(Section srcSection : srcDoc.getSections())
        {
            //Loop through all block level nodes (paragraphs and tables) in the body of the section.
            for (int nodeIdx=0; nodeIdx<srcSection.getBody().getChildNodes().getCount(); nodeIdx++)
            {
                Node srcNode = srcSection.getBody().getChildNodes().get(nodeIdx);
                //Do not insert node if it is a last empty paragarph in the section.
                Paragraph para = null;
                if(srcNode.getNodeType() == NodeType.PARAGRAPH)
                    para = (Paragraph)srcNode;
                if ((para != null) && para.isEndOfSection() && (!para.hasChildNodes()))
                    break;
                //If current paragraph is first paragraph of srcDoc
                //then appent its content to insertAfterParagraph
                boolean nodeInserted = false;
                if (para != null && para.equals(srcDoc.getFirstSection().getBody().getFirstChild()))
                {
                    nodeInserted = true; // set this flag to know that we already processed this node.
                    for (int i=0; i<para.getChildNodes().getCount(); i++)
                    {
                        Node node = para.getChildNodes().get(i);
                        Node dstNode = dstDoc.importNode(node, true, ImportFormatMode.KEEP_SOURCE_FORMATTING);
                        insertAfterParagraph.appendChild(dstNode);
                    }
                    //If subdocument contains only one paragraph
                    //then copy content of insertBeforeParagraph to insertAfterParagraph
                    //and remove insertBeforeParagraph
                    if (srcDoc.getFirstSection().getBody().getFirstParagraph().equals(getLastParagraphInDocumentWithText(srcDoc)))
                    {
                        while (insertBeforeParagraph.hasChildNodes())
                            insertAfterParagraph.appendChild(insertBeforeParagraph.getFirstChild());
                        insertBeforeParagraph.remove();
                    }
                }
                //If current paragraph is last paragraph of srcDoc
                //then appent its content to insertBeforeParagraph
                if (para != null && para.equals(getLastParagraphInDocumentWithText(srcDoc)))
                {
                    nodeInserted = true; // set this flag to know that we already processed this node.
                    Node previouseNode = null;
                    for (int i=0; i<para.getChildNodes().getCount(); i++)
                    {
                        Node node = para.getChildNodes().get(i);
                        Node dstNode = dstDoc.importNode(node, true, ImportFormatMode.KEEP_SOURCE_FORMATTING);
                        if (previouseNode == null)
                            insertBeforeParagraph.insertBefore(dstNode, insertBeforeParagraph.getFirstChild());
                        else
                            insertBeforeParagraph.insertAfter(dstNode, previouseNode);
                        previouseNode = dstNode;
                    }
                }
                if(!nodeInserted)
                {
                    //This creates a clone of the node, suitable for insertion into the destination document.
                    Node newNode = dstDoc.importNode(srcNode, true, ImportFormatMode.KEEP_SOURCE_FORMATTING);
                    //Insert new node after the reference node.
                    dstStory.insertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }
    }
    public static Node getLastParagraphInDocumentWithText(Document doc)
    {
         Node currentNode = doc.getLastSection().getBody().getLastChild();
        while(currentNode != null && currentNode.getNodeType() == NodeType.PARAGRAPH && currentNode.getRange().getText().trim().equals(""))
             currentNode = currentNode.getPreviousSibling();
            return currentNode;
    }
        public static Document generateDocument(Document srcDoc, ArrayList nodes)
                throws Exception {
            // Create a blank document.
            Document dstDoc = new Document();
            // Remove the first paragraph from the empty document.
            dstDoc.getFirstSection().getBody().removeAllChildren();
            // Import each node from the list into the new document. Keep the
            // original formatting of the node.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc,
                    ImportFormatMode.KEEP_SOURCE_FORMATTING);
            for (com.aspose.words.Node node : (Iterable<com.aspose.words.Node>) nodes) {
                com.aspose.words.Node importNode = importer.importNode(node, true);
                dstDoc.getFirstSection().getBody().appendChild(importNode);
            }
            // Return the generated document.
            return dstDoc;
        }
        public void insertFormatedNodes(ArrayList nodes) {
            NodeImporter importer = new NodeImporter(this.myDocument,
                    this.myDocument, ImportFormatMode.KEEP_SOURCE_FORMATTING);
            for (com.aspose.words.Node node : (Iterable<com.aspose.words.Node>) nodes) {
                com.aspose.words.Node importNode = importer.importNode(node, true);
                this.myDocumentBuilder.insertNode(importNode);
                // dstDoc.getFirstSection().getBody().appendChild(importNode);
            }
        }
        /**
         * Inserts content of the external document after the specified node.
         * Section breaks and section formatting of the inserted document are
         * ignored.
         *
         * @param insertAfterNode
         *            Node in the destination document after which the content
         *            should be inserted. This node should be a block level node
         *            (paragraph or table).
         * @param srcDoc
         *            The document to insert.
         */
        public static void insertDocument(com.aspose.words.Node insertAfterNode,
                                          Document srcDoc) throws Exception {
            // Make sure that the node is either a paragraph or table.
            if ((insertAfterNode.getNodeType() != NodeType.PARAGRAPH)
                    & (insertAfterNode.getNodeType() != NodeType.TABLE))
                throw new IllegalArgumentException(
                        "The destination node should be either a paragraph or table.");
            // We will be inserting into the parent of the destination paragraph.
            CompositeNode dstStory = insertAfterNode.getParentNode();
            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(srcDoc,
                    insertAfterNode.getDocument(),
                    ImportFormatMode.USE_DESTINATION_STYLES);
            // Loop through all sections in the source document.
            for (Section srcSection : srcDoc.getSections()) {
                // Loop through all block level nodes (paragraphs and tables) in the
                // body of the section.
                for (com.aspose.words.Node srcNode : (Iterable<com.aspose.words.Node>) srcSection
                        .getBody()) {
                    // Let's skip the node if it is a last empty paragraph in a
                    // section.
                    if (srcNode.getNodeType() == (NodeType.PARAGRAPH)) {
                        Paragraph para = (Paragraph) srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }
                    // This creates a clone of the node, suitable for insertion into
                    // the destination document.
                    com.aspose.words.Node newNode = importer.importNode(srcNode,
                            true);
                    // Insert new node after the reference node.
                    dstStory.insertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }
        /*** ASPOSE ADD-IN ***/
        /**
         * Extracts a range of nodes from a document found between specified markers
         * and returns a copy of those nodes. Content can be extracted between
         * inline nodes, block level nodes, and also special nodes such as Comment
         * or Boomarks. Any combination of different marker types can used.
         *
         * @param startNode
         *            The node which defines where to start the extraction from the
         *            document. This node can be block or inline level of a body.
         * @param endNode
         *            The node which defines where to stop the extraction from the
         *            document. This node can be block or inline level of body.
         * @param isInclusive
         *            Should the marker nodes be included.
         */
        public static ArrayList extractContent(com.aspose.words.Node startNode,
                                               com.aspose.words.Node endNode, boolean isInclusive)
                throws Exception {
            // First check that the nodes passed to this method are valid for use.
            verifyParameterNodes(startNode, endNode);
            // Create a list to store the extracted nodes.
            ArrayList<com.aspose.words.Node> nodes = new ArrayList<com.aspose.words.Node>();
            // Keep a record of the original nodes passed to this method so we can
            // split marker nodes if needed.
            com.aspose.words.Node originalStartNode = startNode;
            com.aspose.words.Node originalEndNode = endNode;
            // Extract content based on block level nodes (paragraphs and tables).
            // Traverse through parent nodes to find them.
            // We will split the content of first and last nodes depending if the
            // marker nodes are inline
            while (startNode.getParentNode().getNodeType() != NodeType.BODY)
                startNode = startNode.getParentNode();
            while (endNode.getParentNode().getNodeType() != NodeType.BODY)
                endNode = endNode.getParentNode();
            boolean isExtracting = true;
            boolean isStartingNode = true;
            boolean isEndingNode;
            // The current node we are extracting from the document.
            com.aspose.words.Node currNode = startNode;
            // Begin extracting content. Process all block level nodes and
            // specifically split the first and last nodes when needed so paragraph
            // formatting is retained.
            // Method is little more complex than a regular extractor as we need to
            // factor in extracting using inline nodes, fields, bookmarks etc as to
            // make it really useful.
            while (isExtracting) {
                // Clone the current node and its children to obtain a copy.
                CompositeNode cloneNode = (CompositeNode) currNode.deepClone(true);
                isEndingNode = currNode.equals(endNode);
                if (isStartingNode || isEndingNode) {
                    // We need to process each marker separately so pass it off to a
                    // separate method instead.
                    if (isStartingNode) {
                        processMarker(cloneNode, nodes, originalStartNode,
                                isInclusive, isStartingNode, isEndingNode);
                        isStartingNode = false;
                    }
                    // Conditional needs to be separate as the block level start and
                    // end markers maybe the same node.
                    if (isEndingNode) {
                        processMarker(cloneNode, nodes, originalEndNode,
                                isInclusive, isStartingNode, isEndingNode);
                        isExtracting = false;
                    }
                } else
                    // Node is not a start or end marker, simply add the copy to the
                    // list.
                    nodes.add(cloneNode);
                // Move to the next node and extract it. If next node is null that
                // means the rest of the content is found in a different section.
                if (currNode.getNextSibling() == null && isExtracting) {
                    // Move to the next section.
                    Section nextSection = (Section) currNode.getAncestor(
                            NodeType.SECTION).getNextSibling();
                    currNode = nextSection.getBody().getFirstChild();
                } else {
                    // Move to the next node in the body.
                    currNode = currNode.getNextSibling();
                }
            }
            // Return the nodes between the node markers.
            return nodes;
        }
        /**
         * Checks the input parameters are correct and can be used. Throws an
         * exception if there is any problem.
         */
        protected static void verifyParameterNodes(com.aspose.words.Node startNode,
                                                   com.aspose.words.Node endNode) throws Exception {
            // The order in which these checks are done is important.
            if (startNode == null)
                throw new IllegalArgumentException("Start node cannot be null");
            if (endNode == null)
                throw new IllegalArgumentException("End node cannot be null");
            if (!startNode.getDocument().equals(endNode.getDocument()))
                throw new IllegalArgumentException(
                        "Start node and end node must belong to the same document");
            if (startNode.getAncestor(NodeType.BODY) == null
                    || endNode.getAncestor(NodeType.BODY) == null)
                throw new IllegalArgumentException(
                        "Start node and end node must be a child or descendant of a body");
            // Check the end node is after the start node in the DOM tree
            // First check if they are in different sections, then if they're not
            // check their position in the body of the same section they are in.
            Section startSection = (Section) startNode
                    .getAncestor(NodeType.SECTION);
            Section endSection = (Section) endNode.getAncestor(NodeType.SECTION);
            int startIndex = startSection.getParentNode().indexOf(startSection);
            int endIndex = endSection.getParentNode().indexOf(endSection);
            if (startIndex == endIndex) {
                if (startSection.getBody().indexOf(startNode) > endSection
                        .getBody().indexOf(endNode))
                    throw new IllegalArgumentException(
                            "The end node must be after the start node in the body");
            } else if (startIndex > endIndex)
                throw new IllegalArgumentException(
                        "The section of end node must be after the section start node");
        }
        /**
         * Checks if a node passed is an inline node.
         */
        protected static boolean isInline(com.aspose.words.Node node)
                throws Exception {
            // Test if the node is desendant of a Paragraph or Table node and also
            // is not a paragraph or a table a paragraph inside a comment class
            // which is decesant of a pararaph is possible.
            return ((node.getAncestor(NodeType.PARAGRAPH) != null || node
                    .getAncestor(NodeType.TABLE) != null) && !(node.getNodeType() == NodeType.PARAGRAPH || node
                    .getNodeType() == NodeType.TABLE));
        }
        /**
         * Removes the content before or after the marker in the cloned node
         * depending on the type of marker.
         */
        protected static void processMarker(CompositeNode cloneNode, ArrayList nodes,
                                            com.aspose.words.Node node, boolean isInclusive,
                                            boolean isStartMarker, boolean isEndMarker) throws Exception {
            // If we are dealing with a block level node just see if it should be
            // included and add it to the list.
            if (!isInline(node)) {
                // Don't add the node twice if the markers are the same node
                if (!(isStartMarker && isEndMarker)) {
                    if (isInclusive)
                        nodes.add(cloneNode);
                }
                return;
            }
            // If a marker is a FieldStart node check if it's to be included or not.
            // We assume for simplicity that the FieldStart and FieldEnd appear in
            // the same paragraph.
            if (node.getNodeType() == NodeType.FIELD_START) {
                // If the marker is a start node and is not be included then skip to
                // the end of the field.
                // If the marker is an end node and it is to be included then move
                // to the end field so the field will not be removed.
                if ((isStartMarker && !isInclusive)
                        || (!isStartMarker && isInclusive)) {
                    while (node.getNextSibling() != null
                            && node.getNodeType() != NodeType.FIELD_END)
                        node = node.getNextSibling();
                }
            }
            // If either marker is part of a comment then to include the comment
            // itself we need to move the pointer forward to the Comment
            // node found after the CommentRangeEnd node.
            if (node.getNodeType() == NodeType.COMMENT_RANGE_END) {
                while (node.getNextSibling() != null
                        && node.getNodeType() != NodeType.COMMENT)
                    node = node.getNextSibling();
            }
            // Find the corresponding node in our cloned node by index and return
            // it.
            // If the start and end node are the same some child nodes might already
            // have been removed. Subtract the
            // difference to get the right index.
            int indexDiff = node.getParentNode().getChildNodes().getCount()
                    - cloneNode.getChildNodes().getCount();
            // Child node count identical.
            if (indexDiff == 0)
                node = cloneNode.getChildNodes().get(
                        node.getParentNode().indexOf(node));
            else
                node = cloneNode.getChildNodes().get(
                        node.getParentNode().indexOf(node) - indexDiff);
            // Remove the nodes up to/from the marker.
            boolean isSkip;
            boolean isProcessing = true;
            boolean isRemoving = isStartMarker;
            com.aspose.words.Node nextNode = cloneNode.getFirstChild();
            while (isProcessing && nextNode != null) {
                com.aspose.words.Node currentNode = nextNode;
                isSkip = false;
                if (currentNode.equals(node)) {
                    if (isStartMarker) {
                        isProcessing = false;
                        if (isInclusive)
                            isRemoving = false;
                    } else {
                        isRemoving = true;
                        if (isInclusive)
                            isSkip = true;
                    }
                }
                nextNode = nextNode.getNextSibling();
                if (isRemoving && !isSkip)
                    currentNode.remove();
            }
            // After processing the composite node may become empty. If it has don't
            // include it.
            if (!(isStartMarker && isEndMarker)) {
                if (cloneNode.hasChildNodes())
                    nodes.add(cloneNode);
            }
        }
    public class RemoveLineReplaceHandler implements IReplacingCallback
    {
         public int replacing(ReplacingArgs args) throws Exception
         {
             Node para = args.getMatchNode().getParentNode();
                     para.remove();
             return ReplaceAction.SKIP;
         }
    }
}