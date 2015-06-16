<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class HelloWorld {

    public static function printHelloWorld() {


        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/helloworld/data/";

        // Create a blank document.
        $documentObject = new Java("com.aspose.words.Document");

        // DocumentBuilder provides members to easily add content to a document.
        $builderBoject = new Java("com.aspose.words.DocumentBuilder",$documentObject);
        // Write a new paragraph in the document with the text "Hello World!"


        $builderBoject->writeln("Hello World!");
        // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
        // Aspose.Words supports saving any document in many more formats.
        $documentObject->save($dataDir . "HelloWorld.docx");

    }

}