<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class LoadAndSaveToDisk {

    public  static function saveToDisk() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/loadandsavetodisk/data/";
        // Load the document from the absolute path on disk.
        $doc = new Java("com.aspose.words.Document", $dataDir . "Document.doc");
        // Save the document as DOCX document.");
        $doc->save($dataDir . "Document Out.docx");

    }

}