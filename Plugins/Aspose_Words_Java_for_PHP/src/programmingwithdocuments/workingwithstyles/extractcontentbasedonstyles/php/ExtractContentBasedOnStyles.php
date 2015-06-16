<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */


class ExtractContentBasedOnStyles {

    public static function main() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithstyles/extractcontentbasedonstyles/data/";

        //ExStart
        //ExId:ExtractContentBasedOnStyles_Main
        //ExSummary:Run queries and display results.
        // Open the document.
        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");

        // Define style names as they are specified in the Word document.
        $para_style = "Heading 1";
        $run_style = "Intense Emphasis";

        // Collect paragraphs with defined styles.
        // Show the number of collected paragraphs and display the text of this paragraphs.
        $paragraphs = ExtractContentBasedOnStyles::paragraphsByStyleName($doc, $para_style);
        $para_size = $paragraphs->size();
        $para_size = java_values($para_size);
        echo "Paragraphs with \{$para_style}\ styles $para_size  : <br>";


        //echo "<PRE>"; echo java_inspect($paragraphs); exit;
        $paragraphs = $paragraphs->toArray();


        $saveFormat = Java("com.aspose.words.SaveFormat");

        foreach ($paragraphs as $paragraph) {
            echo $paragraph->toString($saveFormat->TEXT) . "<BR>";
        }


        // Collect runs with defined styles.
        // Show the number of collected runs and display the text of this runs.
        $runs = ExtractContentBasedOnStyles::runsByStyleName($doc, $run_style);

        $runs_size = $runs->size();
        $runs_size = java_values($runs_size);

        echo "<BR> Runs with \{$run_style}\ styles $runs_size <BR>";
        $runs = $runs->toArray();
        foreach ($runs as $run) {

            echo $run->getRange()->getText() . "<BR>";
        }

    }

    public static function paragraphsByStyleName($doc, $styleName) {

        // Create an array to collect paragraphs of the specified style.
        $paragraphsWithStyle = new Java("java.util.ArrayList");
        // Get all paragraphs from the document.
        $nodeType = Java("com.aspose.words.NodeType");
        $paragraphs = $doc->getChildNodes($nodeType->PARAGRAPH, true);
        $paragraphs = $paragraphs->toArray();

        //echo "<PRE>"; echo java_inspect($paragraphs); exit;
        // Look through all paragraphs to find those with the specified style.
        foreach ($paragraphs as $paragraph)
        {
            if (java_values($paragraph->getParagraphFormat()->getStyle()->getName()->equals($styleName)))
                $paragraphsWithStyle->add($paragraph);
        }


        return $paragraphsWithStyle;
    }

    public static function runsByStyleName($doc, $styleName) {

        // Create an array to collect runs of the specified style.
        $runsWithStyle = new Java("java.util.ArrayList");
        // Get all runs from the document.
        $nodeType = Java("com.aspose.words.NodeType");
        $runs = $doc->getChildNodes($nodeType->RUN, true);
        // Look through all runs to find those with the specified style.
        $runs = $runs->toArray();
        foreach ($runs as $run)
        {
            if (java_values($run->getFont()->getStyle()->getName()->equals($styleName)))
                $runsWithStyle->add($run);
        }
        return $runsWithStyle;
    }
}
