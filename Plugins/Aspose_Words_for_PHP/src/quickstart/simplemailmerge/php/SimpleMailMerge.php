<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class SimpleMailMerge {

    public static function mailmerge() {


        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/simplemailmerge/data/";

        $doc = new Java("com.aspose.words.Document", $dataDir. "Template.doc");
        // Fill the fields in the document with user data.
        $doc->getMailMerge()->execute(
            array("FullName", "Company", "Address", "Address2", "City"),
            array("James Bond", "MI5 Headquarters", "Milbank", "", "London")
        );

        // Saves the document to disk.
        $doc->save($dataDir . "MailMerge Result Out.docx");

    }

}