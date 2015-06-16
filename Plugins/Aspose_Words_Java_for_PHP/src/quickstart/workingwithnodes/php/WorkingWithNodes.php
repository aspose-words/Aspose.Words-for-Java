<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
class WorkingWithNodes {

    public static function main(){

        // Create a new document.
        $doc = new Java("com.aspose.words.Document");

        // Creates and adds a paragraph node to the document.
        $para = new Java("com.aspose.words.Paragraph",$doc);

        // Typed access to the last section of the document.
        $section = $doc->getLastSection();
        $section->getBody()->appendChild($para);

        // Next print the node type of one of the nodes in the document.
        $nodeType = $doc->getFirstSection()->getBody()->getNodeType();

        $node = new Java("com.aspose.words.Node");

        echo "NodeType: " . $node->nodeTypeToString($nodeType);
    }

}