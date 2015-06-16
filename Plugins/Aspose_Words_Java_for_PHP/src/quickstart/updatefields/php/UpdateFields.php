<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class UpdateFields {

    public static function update(){

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/quickstart/updatefields/data/";
        // Demonstrates how to insert fields and update them using Aspose.Words.
        // First create a blank document.
        $doc = new Java("com.aspose.words.Document");
        // Use the document builder to insert some content and fields.
        $builder = new Java("com.aspose.words.DocumentBuilder",$doc);
        // Insert a table of contents at the beginning of the document.
        $builder->insertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        $builder->writeln();
        // Insert some other fields.
        $builder->write("Page: ");
        $builder->insertField("PAGE");
        $builder->write(" of ");
        $builder->insertField("NUMPAGES");
        $builder->writeln();
        $builder->write("Date: ");
        $builder->insertField("DATE");
        // Start the actual document content on the second page.
        $breakType = new Java("com.aspose.words.BreakType");

        $builder->insertBreak($breakType->SECTION_BREAK_NEW_PAGE);
        // Build a document with complex structure by applying different heading styles thus creating TOC entries.

        $styleIdentifier = new Java("com.aspose.words.StyleIdentifier");
        $builder->getParagraphFormat()->setStyleIdentifier($styleIdentifier->HEADING_1);
        $builder->writeln("Heading 1");
        $builder->getParagraphFormat()->setStyleIdentifier($styleIdentifier->HEADING_2);
        $builder->writeln("Heading 1.1");
        $builder->writeln("Heading 1.2");
        $builder->getParagraphFormat()->setStyleIdentifier($styleIdentifier->HEADING_1);
        $builder->writeln("Heading 2");
        $builder->writeln("Heading 3");
        // Move to the next page.
        $builder->insertBreak($breakType->PAGE_BREAK);
        $builder->getParagraphFormat()->setStyleIdentifier($styleIdentifier->HEADING_2);
        $builder->writeln("Heading 3.1");
        $builder->getParagraphFormat()->setStyleIdentifier($styleIdentifier->HEADING_3);
        $builder->writeln("Heading 3.1.1");
        $builder->writeln("Heading 3.1.2");
        $builder->writeln("Heading 3.1.3");
        $builder->getParagraphFormat()->setStyleIdentifier($styleIdentifier->HEADING_2);
        $builder->writeln("Heading 3.2");
        $builder->writeln("Heading 3.3");
        echo "Updating all fields in the document.";
        // Call the method below to update the TOC.
        $doc->updateFields();
        $doc->save($dataDir . "Document Field Update Out.docx");

    }

}