<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */


class CopyBookmarkedText {

    public static function main() {

        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithbookmarks/copybookmarkedtext/data/";

        // Load the source document.
        $srcDoc = new Java("com.aspose.words.Document", $dataDir . "Template.doc");

        // This is the bookmark whose content we want to copy.
        $srcBookmark = $srcDoc->getRange()->getBookmarks()->get("ntf010145060");

        // We will be adding to this document.
        $dstDoc = new Java("com.aspose.words.Document");

        // Let's say we will be appending to the end of the body of the last section.
        $dstNode = $dstDoc->getLastSection()->getBody();

        // It is a good idea to use this import context object because multiple nodes are being imported.
        // If you import multiple times without a single context, it will result in many styles created.

        $importFormatMode = java('com.aspose.words.ImportFormatMode');
        $importer = new Java("com.aspose.words.NodeImporter", $srcDoc, $dstDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);

        // Do it once.
        CopyBookmarkedText::appendBookmarkedText($importer, $srcBookmark, $dstNode);

        // Do it one more time for fun.
        CopyBookmarkedText::appendBookmarkedText($importer, $srcBookmark, $dstNode);

        // Save the finished document.
        $dstDoc->save($dataDir . "Template Out.doc");

    }

    /**
     * Copies content of the bookmark and adds it to the end of the specified node.
     * The destination node can be in a different document.
     *
     * @param importer Maintains the import context.
     * @param srcBookmark The input bookmark.
     * @param dstNode Must be a node that can contain paragraphs (such as a Story).
     */

    private static function appendBookmarkedText($importer,$srcBookmark,$dstNode ) {

        $startPara = $srcBookmark->getBookmarkStart()->getParentNode();

            // This is the paragraph that contains the end of the bookmark.
        $endPara = $srcBookmark->getBookmarkEnd()->getParentNode();

        if ((java_values($startPara) == null) || (java_values($endPara) == null))
            throw new Exception("Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet.");

            // Limit ourselves to a reasonably simple scenario.
        $spara =  java_values($startPara->getParentNode());
        $epara = java_values($endPara->getParentNode());
        if (trim($spara) != trim($epara)){
            throw new Exception("Start and end paragraphs have different parents, cannot handle this scenario yet.");
        }

            // We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
            // therefore the node at which we stop is one after the end paragraph.

        $endNode = $endPara->getNextSibling();


            // This is the loop to go through all paragraph-level nodes in the bookmark.
        $curNode = $startPara;
        $cNode = java_values($curNode);
        $eNode = java_values($endNode);
        //echo $cNode . "<BR>1" . $eNode; exit;
        while(trim($cNode) != trim($eNode) ) {
            // This creates a copy of the current node and imports it (makes it valid) in the context
            // of the destination document. Importing means adjusting styles and list identifiers correctly.
            $newNode = $importer->importNode(java_values($curNode), true);

            $curNode = $curNode->getNextSibling();

            $cNode = java_values($curNode);
            $dstNode->appendChild(java_values($newNode));

        }

    }


}