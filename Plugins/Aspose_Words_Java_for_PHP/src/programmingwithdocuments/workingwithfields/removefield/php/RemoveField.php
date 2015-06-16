<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */


class RemoveField {

    public static function main() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithfields/removefield/data/";

        $doc = new Java("com.aspose.words.Document" , $dataDir . "Field.RemoveField.doc");

        //ExStart
        //ExFor:Field.Remove
        //ExId:DocumentBuilder_RemoveField
        //ExSummary:Removes a field from the document.
        $field = $doc->getRange()->getFields()->get(0);
        // Calling this method completely removes the field from the document.
        $field->remove();
        //ExEnd
    }

}