package DocsExamples.Programming_with_Documents.Contents_Management;

// ********* THIS FILE IS AUTO PORTED *********

import java.util.ArrayList;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Section;
import com.aspose.words.Paragraph;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.ms.System.msString;
import com.aspose.words.NodeImporter;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.CompositeNode;
import com.aspose.ms.System.Diagnostics.Debug;


class ExtractContentHelper
{
    //ExStart:CommonExtractContent
    public static ArrayList<Node> extractContent(Node startNode, Node endNode, boolean isInclusive)
    {
        // First, check that the nodes passed to this method are valid for use.
        verifyParameterNodes(startNode, endNode);

        // Create a list to store the extracted nodes.
        ArrayList<Node> nodes = new ArrayList<Node>();

        // If either marker is part of a comment, including the comment itself, we need to move the pointer
        // forward to the Comment Node found after the CommentRangeEnd node.
        if (endNode.getNodeType() == NodeType.COMMENT_RANGE_END && isInclusive)
        {
            Node node = findNextNode(NodeType.COMMENT, endNode.getNextSibling());
            if (node != null)
                endNode = node;
        }

        // Keep a record of the original nodes passed to this method to split marker nodes if needed.
        Node originalStartNode = startNode;
        Node originalEndNode = endNode;

        // Extract content based on block-level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        // We will split the first and last nodes' content, depending if the marker nodes are inline.
        startNode = getAncestorInBody(startNode);
        endNode = getAncestorInBody(endNode);

        boolean isExtracting = true;
        boolean isStartingNode = true;
        // The current node we are extracting from the document.
        Node currNode = startNode;

        // Begin extracting content. Process all block-level nodes and specifically split the first
        // and last nodes when needed, so paragraph formatting is retained.
        // Method is a little more complicated than a regular extractor as we need to factor
        // in extracting using inline nodes, fields, bookmarks, etc. to make it useful.
        while (isExtracting)
        {
            // Clone the current node and its children to obtain a copy.
            Node cloneNode = currNode.deepClone(true);
            boolean isEndingNode = currNode.equals(endNode);

            if (isStartingNode || isEndingNode)
            {
                // We need to process each marker separately, so pass it off to a separate method instead.
                // End should be processed at first to keep node indexes.
                if (isEndingNode)
                {
                    // !isStartingNode: don't add the node twice if the markers are the same node.
                    processMarker(cloneNode, nodes, originalEndNode, currNode, isInclusive,
                        false, !isStartingNode, false);
                    isExtracting = false;
                }

                // Conditional needs to be separate as the block level start and end markers, maybe the same node.
                if (isStartingNode)
                {
                    processMarker(cloneNode, nodes, originalStartNode, currNode, isInclusive,
                        true, true, false);
                    isStartingNode = false;
                }
            }
            else
                // Node is not a start or end marker, simply add the copy to the list.
                nodes.add(cloneNode);

            // Move to the next node and extract it. If the next node is null,
            // the rest of the content is found in a different section.
            if (currNode.getNextSibling() == null && isExtracting)
            {
                // Move to the next section.
                Section nextSection = (Section) currNode.getAncestor(NodeType.SECTION).getNextSibling();
                currNode = nextSection.getBody().getFirstChild();
            }
            else
            {
                // Move to the next node in the body.
                currNode = currNode.getNextSibling();
            }
        }

        // For compatibility with mode with inline bookmarks, add the next paragraph (empty).
        if (isInclusive && originalEndNode == endNode && !originalEndNode.isComposite())
            includeNextParagraph(endNode, nodes);

        // Return the nodes between the node markers.
        return nodes;
    }
    //ExEnd:CommonExtractContent

    public static ArrayList<Paragraph> paragraphsByStyleName(Document doc, String styleName)
    {
        // Create an array to collect paragraphs of the specified style.
        ArrayList<Paragraph> paragraphsWithStyle = new ArrayList<Paragraph>();
        
        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);
        
        // Look through all paragraphs to find those with the specified style.
        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs)
        {
            if (msString.equals(paragraph.getParagraphFormat().getStyle().getName(), styleName))
                paragraphsWithStyle.add(paragraph);
        }

        return paragraphsWithStyle;
    }

    //ExStart:CommonGenerateDocument
    public static Document generateDocument(Document srcDoc, ArrayList<Node> nodes) throws Exception
    {
        Document dstDoc = new Document();
        // Remove the first paragraph from the empty document.
        dstDoc.getFirstSection().getBody().removeAllChildren();

        // Import each node from the list into the new document. Keep the original formatting of the node.
        NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        for (Node node : nodes)
        {
            Node importNode = importer.importNode(node, true);
            dstDoc.getFirstSection().getBody().appendChild(importNode);
        }

        return dstDoc;
    }
    //ExEnd:CommonGenerateDocument

    //ExStart:CommonExtractContentHelperMethods
    private static void verifyParameterNodes(Node startNode, Node endNode)
    {
        // The order in which these checks are done is important.
        if (startNode == null)
            throw new IllegalArgumentException("Start node cannot be null");
        if (endNode == null)
            throw new IllegalArgumentException("End node cannot be null");

        if (!startNode.getDocument().equals(endNode.getDocument()))
            throw new IllegalArgumentException("Start node and end node must belong to the same document");

        if (startNode.getAncestor(NodeType.BODY) == null || endNode.getAncestor(NodeType.BODY) == null)
            throw new IllegalArgumentException("Start node and end node must be a child or descendant of a body");

        // Check the end node is after the start node in the DOM tree.
        // First, check if they are in different sections, then if they're not,
        // check their position in the body of the same section.
        Section startSection = (Section) startNode.getAncestor(NodeType.SECTION);
        Section endSection = (Section) endNode.getAncestor(NodeType.SECTION);

        int startIndex = startSection.getParentNode().indexOf(startSection);
        int endIndex = endSection.getParentNode().indexOf(endSection);

        if (startIndex == endIndex)
        {
            if (startSection.getBody().indexOf(getAncestorInBody(startNode)) >
                endSection.getBody().indexOf(getAncestorInBody(endNode)))
                throw new IllegalArgumentException("The end node must be after the start node in the body");
        }
        else if (startIndex > endIndex)
            throw new IllegalArgumentException("The section of end node must be after the section start node");
    }

    private static Node findNextNode(/*NodeType*/int nodeType, Node fromNode)
    {
        if (fromNode == null || fromNode.getNodeType() == nodeType)
            return fromNode;

        if (fromNode.isComposite())
        {
            Node node = findNextNode(nodeType, ((CompositeNode) fromNode).getFirstChild());
            if (node != null)
                return node;
        }

        return findNextNode(nodeType, fromNode.getNextSibling());
    }

    private boolean isInline(Node node)
    {
        // Test if the node is a descendant of a Paragraph or Table node and is not a paragraph
        // or a table a paragraph inside a comment class that is decent of a paragraph is possible.
        return ((node.getAncestor(NodeType.PARAGRAPH) != null || node.getAncestor(NodeType.TABLE) != null) &&
                !(node.getNodeType() == NodeType.PARAGRAPH || node.getNodeType() == NodeType.TABLE));
    }

    private static void processMarker(Node cloneNode, ArrayList<Node> nodes, Node node, Node blockLevelAncestor,
        boolean isInclusive, boolean isStartMarker, boolean canAdd, boolean forceAdd)
    {
        // If we are dealing with a block-level node, see if it should be included and add it to the list.
        if (node == blockLevelAncestor)
        {
            if (canAdd && isInclusive)
                nodes.add(cloneNode);
            return;
        }

        // cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor
        // is the node's ancestor that means it is a composite node.
        com.aspose.ms.System.Diagnostics.Debug.assert(cloneNode.isComposite());

        // If a marker is a FieldStart node check if it's to be included or not.
        // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        if (node.getNodeType() == NodeType.FIELD_START)
        {
            // If the marker is a start node and is not included, skip to the end of the field.
            // If the marker is an end node and is to be included, then move to the end field so the field will not be removed.
            if (isStartMarker && !isInclusive || !isStartMarker && isInclusive)
            {
                while (node.getNextSibling() != null && node.getNodeType() != NodeType.FIELD_END)
                    node = node.getNextSibling();
            }
        }

        // Support a case if the marker node is on the third level of the document body or lower.
        ArrayList<Node> nodeBranch = fillSelfAndParents(node, blockLevelAncestor);

        // Process the corresponding node in our cloned node by index.
        Node currentCloneNode = cloneNode;
        for (int i = nodeBranch.size() - 1; i >= 0; i--)
        {
            Node currentNode = nodeBranch.get(i);
            int nodeIndex = currentNode.getParentNode().indexOf(currentNode);
            currentCloneNode = ((CompositeNode) currentCloneNode).getChildNodes().get(nodeIndex);

            removeNodesOutsideOfRange(currentCloneNode, isInclusive || (i > 0), isStartMarker);
        }

        // After processing, the composite node may become empty if it has doesn't include it.
        if (canAdd &&
            (forceAdd || ((CompositeNode) cloneNode).hasChildNodes()))
            nodes.add(cloneNode);
    }

    private static void removeNodesOutsideOfRange(Node markerNode, boolean isInclusive, boolean isStartMarker)
    {
        boolean isProcessing = true;
        boolean isRemoving = isStartMarker;
        Node nextNode = markerNode.getParentNode().getFirstChild();

        while (isProcessing && nextNode != null)
        {
            Node currentNode = nextNode;
            boolean isSkip = false;

            if (currentNode.equals(markerNode))
            {
                if (isStartMarker)
                {
                    isProcessing = false;
                    if (isInclusive)
                        isRemoving = false;
                }
                else
                {
                    isRemoving = true;
                    if (isInclusive)
                        isSkip = true;
                }
            }

            nextNode = nextNode.getNextSibling();
            if (isRemoving && !isSkip)
                currentNode.remove();
        }
    }

    private static ArrayList<Node> fillSelfAndParents(Node node, Node tillNode)
    {
        ArrayList<Node> list = new ArrayList<Node>();
        Node currentNode = node;

        while (currentNode != tillNode)
        {
            list.add(currentNode);
            currentNode = currentNode.getParentNode();
        }

        return list;
    }

    private static void includeNextParagraph(Node node, ArrayList<Node> nodes)
    {
        Paragraph paragraph = (Paragraph) findNextNode(NodeType.PARAGRAPH, node.getNextSibling());
        if (paragraph != null)
        {
            // Move to the first child to include paragraphs without content.
            Node markerNode = paragraph.hasChildNodes() ? paragraph.getFirstChild() : paragraph;
            Node rootNode = getAncestorInBody(paragraph);

            processMarker(rootNode.deepClone(true), nodes, markerNode, rootNode,
                markerNode == paragraph, false, true, true);
        }
    }

    private static Node getAncestorInBody(Node startNode)
    {
        while (startNode.getParentNode().getNodeType() != NodeType.BODY)
            startNode = startNode.getParentNode();
        return startNode;
    }
    //ExEnd:CommonExtractContentHelperMethods
}
