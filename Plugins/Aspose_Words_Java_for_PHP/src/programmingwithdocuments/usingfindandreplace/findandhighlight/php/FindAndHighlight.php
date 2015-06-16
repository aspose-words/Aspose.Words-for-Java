<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use com\aspose\words\IReplacingCallback as IReplacingCallback;
use java\lang\Boolean as Boolean;



class FindAndHighlight {

    public static function main() {
        //$color = Java("java.awt.Color");
        //echo java_values($color->YELLOW); exit;



        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/usingfindandreplace/findandhighlight/data/";

        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");

        // We want the "your document" phrase to be highlighted.

        $pattern = new Java("java.util.regex.Pattern");

        $regex = $pattern->compile("your document", $pattern->CASE_INSENSITIVE);

        $replaceEvaluation = new ReplaceEvaluatorFindAndHighlight();

        $range = $doc->getRange();


        $range->replace(($regex), java_values($replaceEvaluation));

        // Save the output document.
        $doc->save($dataDir . "TestFile Out.doc");
    }

}

class ReplaceEvaluatorFindAndHighlight extends IReplacingCallback  {

    /**
     * This method is called by the Aspose.Words find and replace engine for each match.
     * This method highlights the match string, even if it spans multiple runs.
     */

    public static function replacing($e) {


        // This is a Run node that contains either the beginning or the complete match.
        $currentNode = $e->getMatchNode();

        // The first (and may be the only) run can contain text before the match,
        // in this case it is necessary to split the run.
        if (java_values($e->getMatchOffset()) > 0)
            $currentNode = ReplaceEvaluatorFindAndHighlight::splitRun($currentNode, $e->getMatchOffset());

        // This array is used to store all nodes of the match for further highlighting.
        $runs = new Java("java.util.ArrayList");

        // Find all runs that contain parts of the match string.
        $remainingLength = $e->getMatch()->group()->length();
        while (
            (java_values($remainingLength) > 0) &&
            (java_values($currentNode) != null) &&
            (java_values($currentNode->getText()->length()) <= java_values($remainingLength)))
        {
            $runs->add($currentNode);
            $remainingLength = java_values($remainingLength) - java_values($currentNode->getText()->length());

            // Select the next Run node.
            // Have to loop because there could be other nodes such as BookmarkStart etc.
            do
            {
                $currentNode = $currentNode->getNextSibling();
                $nodeType = Java("com.aspose.words.NodeType");
            }
            while ((java_values($currentNode) != null) && (java_values($currentNode->getNodeType()) != java_values($nodeType->RUN)));
        }

        // Split the last run that contains the match if there is any text left.
        if ((java_values($currentNode) != null) && (java_values($remainingLength) > 0))
        {
            ReplaceEvaluatorFindAndHighlight::splitRun($currentNode, $remainingLength);
            $runs->add($currentNode);
        }

        // Now highlight all runs in the sequence.
        $runs = $runs->toArray();
        $color = Java("java.awt.Color");


        foreach ($runs as $run){

            $run->getFont()->setHighlightColor('yellow');
        }


        // Signal to the replace engine to do nothing because we have already done all what we wanted.
        $replaceAction = Java("com.aspose.words.ReplaceAction");
        return $replaceAction->SKIP;
    }

    public static function splitRun ($run, $position) {

        $afterRun = $run->deepClone(true);
        $afterRun->setText($run->getText()->substring(position));
        $run->setText($run->getText()->substring((0), (0) . ($position)));
        $run->getParentNode()->insertAfter($afterRun, $run);
        return $afterRun;
    }
}