<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class ExtractContent {

    private static $gDataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithdocument/extractcontent/data/";

    public static function main() {

        ExtractContent::extractContentBetweenParagraphs();
        ExtractContent::extractContentBetweenBlockLevelNodes();
        ExtractContent::extractContentBetweenParagraphStyles();
        ExtractContent::extractContentBetweenRuns();
        ExtractContent::extractContentUsingField();
        ExtractContent::extractContentBetweenBookmark();
        ExtractContent::extractContentBetweenCommentRange();

    }

    public static function extractContentBetweenParagraphs(){

        //ExStart
        //ExId:ExtractBetweenNodes_BetweenParagraphs
        //ExSummary:Shows how to extract the content between specific paragraphs using the ExtractContent method above.
        // Load in the document
        $doc = new Java("com.aspose.words.Document", ExtractContent::$gDataDir . "TestFile.doc");

        // Gather the nodes. The GetChild method uses 0-based index
        $nodeType = Java("com.aspose.words.NodeType");
        $startPara = $doc->getFirstSection()->getChild($nodeType->PARAGRAPH, 6, true);
        $endPara =   $doc->getFirstSection()->getChild($nodeType->PARAGRAPH, 10, true);
        // Extract the content between these nodes in the document. Include these markers in the extraction.

        $extractedNodes = ExtractContent::ExtractContents($startPara, $endPara, true);

        // Insert the content into a new separate document and save it to disk.
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodes);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.Paragraphs Out.doc");
        //ExEnd

    }

    public static function extractContentBetweenBlockLevelNodes(){

        //ExStart
        //ExId:ExtractBetweenNodes_BetweenNodes
        //ExSummary:Shows how to extract the content between a paragraph and table using the ExtractContent method.
        // Load in the document
        $doc = new Java("com.aspose.words.Document", ExtractContent::$gDataDir . "TestFile.doc");

        $nodeType = Java("com.aspose.words.NodeType");
        $startPara = $doc->getLastSection()->getChild($nodeType->PARAGRAPH, 2, true);
        $endTable = $doc->getLastSection()->getChild($nodeType->TABLE, 0, true);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        $extractedNodes = ExtractContent::ExtractContents($startPara, $endTable, true);

        // Lets reverse the array to make inserting the content back into the document easier.
        $collections = new Java("java.util.Collections");
        $collections->reverse($extractedNodes);

        while (java_values($extractedNodes->size()) > 0)
        {
            // Insert the last node from the reversed list
            $endTable->getParentNode()->insertAfter($extractedNodes->get(0), $endTable);
            // Remove this node from the list after insertion.
            $extractedNodes->remove(0);
        }

        // Save the generated document to disk.
        $doc->save(ExtractContent::$gDataDir . "TestFile.DuplicatedContent Out.doc");
        //ExEnd


    }

    public static function extractContentBetweenParagraphStyles(){

        //ExStart
        //ExId:ExtractBetweenNodes_BetweenStyles
        //ExSummary:Shows how to extract content between paragraphs with specific styles using the ExtractContent method.
        // Load in the document
        $doc = new Java("com.aspose.words.Document" , ExtractContent::$gDataDir . "TestFile.doc");

        // Gather a list of the paragraphs using the respective heading styles.
        $parasStyleHeading1 = ExtractContent::paragraphsByStyleName($doc, "Heading 1");
        $parasStyleHeading3 = ExtractContent::paragraphsByStyleName($doc, "Heading 3");

        // Use the first instance of the paragraphs with those styles.
        $startPara1 = $parasStyleHeading1->get(0);
        $endPara1 = $parasStyleHeading3->get(0);

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        $extractedNodes = ExtractContent::ExtractContents($startPara1, $endPara1, false);
        // Insert the content into a new separate document and save it to disk.
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodes);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.Styles Out.doc");
        //ExEnd

    }

    public static function extractContentBetweenRuns(){

        //ExStart
        //ExId:ExtractBetweenNodes_BetweenRuns
        //ExSummary:Shows how to extract content between specific runs of the same paragraph using the ExtractContent method.
        // Load in the document
        $doc = new Java("com.aspose.words.Document" , ExtractContent::$gDataDir . "TestFile.doc");

        // Retrieve a paragraph from the first section.
        $nodeType = Java("com.aspose.words.NodeType");
        $para = $doc->getChild($nodeType->PARAGRAPH, 7, true);

        // Use some runs for extraction.
        $startRun = $para->getRuns()->get(1);
        $endRun = $para->getRuns()->get(4);

        // Extract the content between these nodes in the document. Include these markers in the extraction.
        $extractedNodes = ExtractContent::ExtractContents($startRun, $endRun, true);

        // Get the node from the list. There should only be one paragraph returned in the list.
        $node = $extractedNodes->get(0);
        // Print the text of this node to the console.
        $SaveFormat = Java("com.aspose.words.SaveFormat");
        echo $node->toString($SaveFormat->TEXT);
        //ExEnd
    }

    public static function extractContentUsingField(){

        //ExStart
        //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
        //ExId:ExtractBetweenNodes_UsingField
        //ExSummary:Shows how to extract content between a specific field and paragraph in the document using the ExtractContent method.
        // Load in the document
        $doc = new Java("com.aspose.words.Document", ExtractContent::$gDataDir . "TestFile.doc");

        // Use a document builder to retrieve the field start of a merge field.
        $builder = new Java("com.aspose.words.DocumentBuilder", $doc);

        // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        $builder->moveToMergeField("Fullname", false, false);

        // The builder cursor should be positioned at the start of the field.
        $nodeType = Java("com.aspose.words.NodeType");
        $startField = $builder->getCurrentNode();
        $endPara = $doc->getFirstSection()->getChild($nodeType->PARAGRAPH, 5, true);

        // Extract the content between these nodes in the document. Don't include these markers in the extraction.
        $extractedNodes = ExtractContent::ExtractContents($startField, $endPara, false);

        // Insert the content into a new separate document and save it to disk.
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodes);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.Fields Out.pdf");
        //ExEnd
    }

    public static function extractContentBetweenBookmark(){

        //ExStart
        //ExId:ExtractBetweenNodes_BetweenBookmark
        //ExSummary:Shows how to extract the content referenced a bookmark using the ExtractContent method.
        // Load in the document
        $doc = new Java("com.aspose.words.Document" , ExtractContent::$gDataDir . "TestFile.doc");

        // Retrieve the bookmark from the document.
        $bookmark = $doc->getRange()->getBookmarks()->get("Bookmark1");

        // We use the BookmarkStart and BookmarkEnd nodes as markers.
        $bookmarkStart = $bookmark->getBookmarkStart();
        $bookmarkEnd = $bookmark->getBookmarkEnd();

        // Firstly extract the content between these nodes including the bookmark.
        $extractedNodesInclusive = ExtractContent::ExtractContents($bookmarkStart, $bookmarkEnd, true);
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodesInclusive);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.BookmarkInclusive Out.doc");

        // Secondly extract the content between these nodes this time without including the bookmark.
        $extractedNodesExclusive = ExtractContent::ExtractContents($bookmarkStart, $bookmarkEnd, false);
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodesExclusive);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.BookmarkExclusive Out.doc");
        //ExEnd

    }

    public static function extractContentBetweenCommentRange(){

        //ExStart
        //ExId:ExtractBetweenNodes_BetweenComment
        //ExSummary:Shows how to extract content referenced by a comment using the ExtractContent method.
        // Load in the document
        $doc = new Java("com.aspose.words.Document" , ExtractContent::$gDataDir . "TestFile.doc");

        // This is a quick way of getting both comment nodes.
        // Your code should have a proper method of retrieving each corresponding start and end node.
        $nodeType = Java("com.aspose.words.NodeType");
        $commentStart = $doc->getChild($nodeType->COMMENT_RANGE_START, 0, true);
        $commentEnd = $doc->getChild($nodeType->COMMENT_RANGE_END, 0, true);

        // Firstly extract the content between these nodes including the comment as well.
        $extractedNodesInclusive = ExtractContent::ExtractContents($commentStart, $commentEnd, true);
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodesInclusive);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.CommentInclusive Out.doc");

        // Secondly extract the content between these nodes without the comment.
        $extractedNodesExclusive = ExtractContent::ExtractContents($commentStart, $commentEnd, false);
        $dstDoc = ExtractContent::generateDocument($doc, $extractedNodesExclusive);
        $dstDoc->save(ExtractContent::$gDataDir . "TestFile.CommentExclusive Out.doc");
        //ExEnd

    }

    //ExStart
    //ExId:ExtractBetweenNodes_ExtractContent
    //ExSummary:This is a method which extracts blocks of content from a document between specified nodes.
    /**
     * Extracts a range of nodes from a document found between specified markers and returns a copy of those nodes. Content can be extracted
     * between inline nodes, block level nodes, and also special nodes such as Comment or Boomarks. Any combination of different marker types can used.
     *
     * @param startNode The node which defines where to start the extraction from the document. This node can be block or inline level of a body.
     * @param endNode The node which defines where to stop the extraction from the document. This node can be block or inline level of body.
     * @param isInclusive Should the marker nodes be included.
     */
    public static function ExtractContents($startNode, $endNode, $isInclusive) {


        // First check that the nodes passed to this method are valid for use.
        ExtractContent::verifyParameterNodes($startNode, $endNode);

        // Create a list to store the extracted nodes.
        $nodes = new Java("java.util.ArrayList");

        // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.

        $originalStartNode = $startNode;
        $originalEndNode = $endNode;

        // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        // We will split the content of first and last nodes depending if the marker nodes are inline
        $nodeType = Java("com.aspose.words.NodeType");
        //echo java_values($nodeType->BODY);
        while ( java_values($startNode->getParentNode()->getNodeType()) != java_values($nodeType->BODY)) {
            $startNode = $startNode->getParentNode();
        }



        while ( java_values($endNode->getParentNode()->getNodeType()) != java_values($nodeType->BODY)) {
            $endNode = $endNode->getParentNode();
        }


        $isExtracting = true;
        $isStartingNode = true;
        $isEndingNode = '';
        // The current node we are extracting from the document.
        $currNode = $startNode;

        // Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
        // Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.

        while ($isExtracting)
        {

            // Clone the current node and its children to obtain a copy.
            $cloneNode = $currNode->deepClone(true);
            $isEndingNode = $currNode->equals(java_values($endNode));

            if(java_values($isStartingNode) || java_values($isEndingNode))
            {
                // We need to process each marker separately so pass it off to a separate method instead.
                if (java_values($isStartingNode))
                {
                    ExtractContent::processMarker($cloneNode, $nodes, $originalStartNode, $isInclusive, $isStartingNode, $isEndingNode);
                    $isStartingNode = false;
                }

                // Conditional needs to be separate as the block level start and end markers maybe the same node.
                if (java_values($isEndingNode))
                {
                    ExtractContent::processMarker($cloneNode, $nodes, $originalEndNode, $isInclusive, $isStartingNode, $isEndingNode);
                    $isExtracting = false;
                }
            }
            else {
                // Node is not a start or end marker, simply add the copy to the list.
                $nodes->add($cloneNode); //array_push($nodes,java_values($cloneNode));

            }

            // Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
            if (java_values($currNode->getNextSibling()) == null && $isExtracting)
            {
                // Move to the next section.
                $nodeType = Java("com.aspose.words.NodeType");
                $nextSection = $currNode->getAncestor($nodeType->SECTION)->getNextSibling();
                $currNode = $nextSection->getBody()->getFirstChild();
            }
            else
            {
                // Move to the next node in the body.
                $currNode = $currNode->getNextSibling();
            }
        }

        // Return the nodes between the node markers.

        return $nodes;

    }

    //ExEnd

    //ExStart
    //ExId:ExtractBetweenNodes_Helpers
    //ExSummary:The helper methods used by the ExtractContent method.
    /**
     * Checks the input parameters are correct and can be used. Throws an exception if there is any problem.
     */

    public static function verifyParameterNodes($startNode, $endNode) {

        // The order in which these checks are done is important.
        if (java_values($startNode) == null)
            throw new Exception("Start node cannot be null");
        if (java_values($endNode) == null)
            throw new Exception("End node cannot be null");

        if (! java_values($startNode->getDocument()->equals($endNode->getDocument())))
            throw new Exception("Start node and end node must belong to the same document");

        $nodeType = Java("com.aspose.words.NodeType");
        if ( java_values($startNode->getAncestor($nodeType->BODY)) == null || java_values($endNode->getAncestor($nodeType->BODY)) == null)
            throw new Exception("Start node and end node must be a child or descendant of a body");
        // Check the end node is after the start node in the DOM tree
        // First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        $startSection = $startNode->getAncestor($nodeType->SECTION);
        $endSection = $endNode->getAncestor($nodeType->SECTION);

        $startIndex = java_values($startSection->getParentNode()->indexOf($startSection));
        $endIndex = java_values($endSection->getParentNode()->indexOf($endSection));

        if ($startIndex == $endIndex)
        {
            if ( java_values($startSection->getBody()->indexOf($startNode)) > java_values($endSection->getBody()->indexOf($endNode)))
                throw new Exception("The end node must be after the start node in the body");
        }
        else if ($startIndex > $endIndex)
            throw new Exception("The section of end node must be after the section start node");

    }

    public static function generateDocument($srcDoc, $nodes) {

        // Create a blank document.
        $dstDoc = new Java("com.aspose.words.Document");
        // Remove the first paragraph from the empty document.
        $dstDoc->getFirstSection()->getBody()->removeAllChildren();

        // Import each node from the list into the new document. Keep the original formatting of the node.
        $importFormatMode = Java("com.aspose.words.ImportFormatMode");
        $importer = new Java("com.aspose.words.NodeImporter", $srcDoc, $dstDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);

        foreach ($nodes as $node)
        {
            $importNode = $importer->importNode($node, true);
            $dstDoc->getFirstSection()->getBody()->appendChild($importNode);
        }

        // Return the generated document.
        return $dstDoc;
    }

    public static function processMarker($cloneNode, $nodes, $node, $isInclusive, $isStartMarker, $isEndMarker) {


        // If we are dealing with a block level node just see if it should be included and add it to the list.


        if(! java_values(ExtractContent::isInline($node)))
        {


            // Don't add the node twice if the markers are the same node
            if(!($isStartMarker && $isEndMarker))
            {
                if ($isInclusive)
                    $nodes->add($cloneNode); // array_push($nodes,$cloneNode); //nodes.add(cloneNode);
            }
            return;
        }

        // If a marker is a FieldStart node check if it's to be included or not.
        // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        $nodeType = Java("com.aspose.words.NodeType");
        if (java_values($node->getNodeType()) == java_values($nodeType->FIELD_START))
        {
            // If the marker is a start node and is not be included then skip to the end of the field.
            // If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            if (($isStartMarker && !$isInclusive) || (!$isStartMarker && $isInclusive))
            {
                while (java_values($node->getNextSibling()) != null && java_values($node->getNodeType()) != java_values($nodeType->FIELD_END))
                    $node = $node->getNextSibling();

            }
        }

        // If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
        // node found after the CommentRangeEnd node.
        if (java_values($node->getNodeType()) == java_values($nodeType->COMMENT_RANGE_END))
        {
            while (java_values($node->getNextSibling()) != null && java_values($node->getNodeType()) != java_values($nodeType->COMMENT))
                $node = $node->getNextSibling();

        }

        // Find the corresponding node in our cloned node by index and return it.
        // If the start and end node are the same some child nodes might already have been removed. Subtract the
        // difference to get the right index.
        $indexDiff = java_values($node->getParentNode()->getChildNodes()->getCount()) - java_values($cloneNode->getChildNodes()->getCount());

        // Child node count identical.
        if ($indexDiff == 0)
            $node = $cloneNode->getChildNodes()->get($node->getParentNode()->indexOf($node));
        else
            $node = $cloneNode->getChildNodes()->get($node->getParentNode()->indexOf($node) - $indexDiff);

        // Remove the nodes up to/from the marker.
        $isSkip = '';
        $isProcessing = true;
        $isRemoving = $isStartMarker;
        $nextNode = $cloneNode->getFirstChild();
        while ($isProcessing && $nextNode != null)
        {

            $currentNode = $nextNode;
            $isSkip = false;

            if (java_values($currentNode->equals($node)))
            {
                if (java_values($isStartMarker))
                {
                    $isProcessing = false;
                    if (java_values($isInclusive))
                        $isRemoving = false;
                }
                else
                {
                    $isRemoving = true;
                    if (java_values($isInclusive))
                        $isSkip = true;
                }
            }

            $nextNode = $nextNode->getNextSibling();
            if ($isRemoving && !$isSkip)
                $currentNode->remove();
        }

        // After processing the composite node may become empty. If it has don't include it.
        if (!($isStartMarker && $isEndMarker))
        {
            if ($cloneNode->hasChildNodes())
                $nodes->add($cloneNode); // array_push($nodes,$cloneNode); //nodes.add(cloneNode);
        }

    }

    public static function isInline($node) {

        // Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
        $nodeType = Java("com.aspose.words.NodeType");
        return ((java_values($node->getAncestor($nodeType->PARAGRAPH)) != null || java_values($node->getAncestor($nodeType->TABLE)) != null) && !(java_values($node->getNodeType()) == (java_values($nodeType->PARAGRAPH) || java_values($node->getNodeType()) == java_values($nodeType->TABLE))));
    }

    public static function paragraphsByStyleName($doc, $styleName) {

        // Create an array to collect paragraphs of the specified style.
        $paragraphsWithStyle = new Java("java.util.ArrayList");

        // Get all paragraphs from the document.
        $nodeType = Java("com.aspose.words.NodeType");
        $paragraphs = $doc->getChildNodes($nodeType->PARAGRAPH, true);
        $paragraphs_count = $paragraphs->getCount();
        $paragraphs_count = java_values($paragraphs_count);

        // Look through all paragraphs to find those with the specified style.
        $i = 0;
        while($i < $paragraphs_count){

            $paragraphs = $doc->getChildNodes($nodeType->PARAGRAPH, true);
            $paragraph = $paragraphs->get($i);

            if (java_values($paragraph->getParagraphFormat()->getStyle()->getName()->equals($styleName))){

                $paragraphsWithStyle->add($paragraph);
            }



            $i++;
        }

        return $paragraphsWithStyle;

    }



}