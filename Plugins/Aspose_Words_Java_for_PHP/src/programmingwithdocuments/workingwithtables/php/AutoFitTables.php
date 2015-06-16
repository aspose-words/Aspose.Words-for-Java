<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */


class AutoFitTables {

    public static function main() {
        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithtables/data/";

        // Demonstrate autofitting a table to the window.
        AutoFitTables::autoFitTableToWindow($dataDir);

        // Demonstrate autofitting a table to its contents.
        AutoFitTables::autoFitTableToContents($dataDir);

        // Demonstrate autofitting a table to fixed column widths.
        AutoFitTables::autoFitTableToFixedColumnWidths($dataDir);
    }

    public static function autoFitTableToWindow($dataDir) {

        //ExStart
        //ExFor:Table.AutoFit
        //ExFor:AutoFitBehavior
        //ExId:FitTableToPageWidth
        //ExSummary:Autofits a table to fit the page width.
        // Open the document


        $doc = new java("com.aspose.words.Document",$dataDir . "TestFile.doc");
        $nodeType = java("com.aspose.words.NodeType");
		$table = $doc->getChild($nodeType->TABLE, 0, true);

        // Autofit the first table to the page width.
        $autoFitBehavior = new Java("com.aspose.words.AutoFitBehavior");
        $table->autoFit($autoFitBehavior->AUTO_FIT_TO_WINDOW);

        // Save the document to disk.
        $doc->save($dataDir . "TestFile.AutoFitToWindow Out.doc");
        //ExEnd
        $preferredWidthType = new Java("com.aspose.words.PreferredWidthType");

        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getPreferredWidth()->getType()) == java_values($preferredWidthType->PERCENT)) {
            echo "PreferredWidth type is not percent <br />";
        }

        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getPreferredWidth()->getValue()) == 100) {
            echo "PreferredWidth value is different than 100 <br />";
        }

    }

    public static function autoFitTableToContents($dataDir) {

        //ExStart
        //ExFor:Table.AutoFit
        //ExFor:AutoFitBehavior
        //ExId:FitTableToContents
        //ExSummary:Autofits a table in the document to its contents.
        // Open the document
        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");
        $nodeType = new Java("com.aspose.words.NodeType");
        $table = $doc->getChild($nodeType->TABLE, 0, true);

		  // Auto fit the table to the cell contents
        $autoFitBehavior = new Java("com.aspose.words.AutoFitBehavior");
        $table->autoFit($autoFitBehavior->AUTO_FIT_TO_CONTENTS);

		  // Save the document to disk.
        $doc->save($dataDir . "TestFile.AutoFitToContents Out.doc");
        //ExEnd

        $preferredWidthType = new Java("com.aspose.words.PreferredWidthType");
        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getPreferredWidth()->getType()) == java_values($preferredWidthType->AUTO)) {

            echo "PreferredWidth type is not auto <br />";
        }

        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getFirstRow()->getFirstCell()->getCellFormat()->getPreferredWidth()->getType()) == java_values($preferredWidthType->AUTO)) {

            echo "PrefferedWidth on cell is not auto <br />";
        }

        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getFirstRow()->getFirstCell()->getCellFormat()->getPreferredWidth()->getValue()) == 0) {

            echo "PreferredWidth value is not 0 <br />";
        }


    }

    public static function autoFitTableToFixedColumnWidths($dataDir) {

        //ExStart
        //ExFor:Table.AutoFit
        //ExFor:AutoFitBehavior
        //ExId:DisableAutoFitAndUseFixedWidths
        //ExSummary:Disables autofitting and enables fixed widths for the specified table.
        // Open the document
        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");
        $nodeType = new Java("com.aspose.words.NodeType");
        $table = $doc->getChild($nodeType->TABLE, 0, true);

		 // Disable autofitting on this table.
        $autoFitBehavior = new Java("com.aspose.words.AutoFitBehavior");
        $table->autoFit($autoFitBehavior->AUTO_FIT_TO_CONTENTS);

		 // Save the document to disk.
		$doc->save($dataDir . "TestFile.FixedWidth Out.doc");
		 //ExEnd

        $preferredWidthType = new Java("com.aspose.words.PreferredWidthType");
        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getPreferredWidth()->getType()) == java_values($preferredWidthType->AUTO)) {

            echo "PreferredWidth type is not auto <br />";
        }

        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getPreferredWidth()->getValue()) == 0) {

            echo "PreferredWidth value is not 0 <br />";
        }

        if(java_values($doc->getFirstSection()->getBody()->getTables()->get(0)->getFirstRow()->getFirstCell()->getCellFormat()->getWidth()) == 0) {

            echo "Cell width is not correct. <br />";
        }

    }

}