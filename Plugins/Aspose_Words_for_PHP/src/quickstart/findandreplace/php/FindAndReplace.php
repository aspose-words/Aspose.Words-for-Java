<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */


class FindAndReplace {

    public static function replaceText() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/findandreplace/data/";

        // Open the document.
        $doc = new Java("com.aspose.words.Document",$dataDir."ReplaceSimple.doc");
        // Check the text of the document
        echo "Original document text: " . $doc->getRange()->getText() . "<BR>";
        // Replace the text in the document.
        $doc->getRange()->replace("_CustomerName_", "James Bond", false, false);
        // Check the replacement was made.
        echo "Document text after replace: " . $doc->getRange()->getText();
        // Save the modified document.
        $doc->save($dataDir . "ReplaceSimple Out.doc");
    }

}