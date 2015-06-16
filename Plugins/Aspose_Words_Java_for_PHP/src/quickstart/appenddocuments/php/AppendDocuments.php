<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class AppendDocuments {

    public static function AppendDocs() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/appenddocuments/data/";

        // Load the destination and source documents from disk.
        $dstDocObject = new Java("com.aspose.words.Document",$dataDir."TestFile.Destination.doc");
        $srcDocObject = new Java("com.aspose.words.Document",$dataDir."TestFile.Source.doc");

        $importFormatModeObject = new java('com.aspose.words.ImportFormatMode');
        $sourceFormating = $importFormatModeObject->KEEP_SOURCE_FORMATTING;


        // Append the source document to the destination document while keeping the original formatting of the source document.
        $dstDocObject->appendDocument(java_values($srcDocObject), java_values($sourceFormating));
        $dstDocObject->save($dataDir . "TestFile_Out.docx");


    }

}