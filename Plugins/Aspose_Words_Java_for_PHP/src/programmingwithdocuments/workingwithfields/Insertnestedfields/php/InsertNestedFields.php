<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class InsertNestedFields {

    public static function main() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithfields/insertnestedfields/data/";

        $doc = new Java("com.aspose.words.Document"); // Document();
        $builder = new Java("com.aspose.words.DocumentBuilder", $doc); // DocumentBuilder(doc);

        // Insert few page breaks (just for testing)
        $breakType = Java("com.aspose.words.BreakType");
        for ($i = 0; $i < 5; $i++)
            $builder->insertBreak($breakType->PAGE_BREAK);

        // Move DocumentBuilder cursor into the primary footer.
        $headerFooterType = Java("com.aspose.words.HeaderFooterType");
        $builder->moveToHeaderFooter($headerFooterType->FOOTER_PRIMARY);

        // We want to insert a field like this:
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        $field = $builder->insertField("IF ");
        $builder->moveTo($field->getSeparator());
        $builder->insertField("PAGE");
        $builder->write(" <> ");
        $builder->insertField("NUMPAGES");
        $builder->write(" \"See Next Page\" \"Last Page\" ");

        // Finally update the outer field to recalcaluate the final value. Doing this will automatically update
        // the inner fields at the same time.
        $field->update();

        $doc->save($dataDir . "InsertNestedFields Out.docx");
    }
}