<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class LoadAndSaveToStream {

    public static function saveToStream(){

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/loadandsavetostream/data/";
        // Open the stream. Read only access is enough for Aspose.Words to load a document.
        $stream = new Java("java.io.FileInputStream",$dataDir . "Document.doc");


        // Load the entire document into memory.
        $doc = new Java("com.aspose.words.Document", $stream);
        // You can close the stream now, it is no longer needed because the document is in memory.
        $stream->close();
        // ... do something with the document
        // Convert the document to a different format and save to stream.
        $dstStream = new Java("java.io.ByteArrayOutputStream");
        $SaveFormat = new Java("com.aspose.words.SaveFormat");
        $doc->save($dstStream, $SaveFormat->RTF);
        $output = new Java("java.io.FileOutputStream", $dataDir . "Document Out.rtf");
        $output->write($dstStream->toByteArray());
        $output->close();

    }

}