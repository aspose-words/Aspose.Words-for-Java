<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */


class ProcessComments {

    public static function main(){

        // A sample infrastructure.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithcomments/processcomments/data/";

        // Open the document.
        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");

        //ExStart
        //ExId:ProcessComments_Main
        //ExSummary: The demo-code that illustrates the methods for the comments extraction and removal.
        // Extract the information about the comments of all the authors.
        $comments = ProcessComments::extractComments($doc);


        foreach ($comments as $comment ) {

            echo java_values($comment) . "<br>";
        }


        // Remove comments by the "pm" author.
        ProcessComments::removeComments($doc, "pm");


        echo "Comments from \"pm\" are removed! <br>";

        // Extract the information about the comments of the "ks" author.
        $comments = ProcessComments::extractComments($doc, "ks");

        foreach ($comments as $comment ) {

            echo java_values($comment) . "<br>";
        }


        // Remove all comments.
        ProcessComments::removeComments($doc);

        echo "All comments are removed! <br>";

        // Save the document.
        $doc->save($dataDir . "Test File Out.doc");
        //ExEnd

    }

    public static function extractComments() {

        $args = func_get_args();
        $doc = $args[0];

        $collectedComments = new Java("java.util.ArrayList");
        // Collect all comments in the document
        $nodeType = Java("com.aspose.words.NodeType");
        $comments = $doc->getChildNodes($nodeType->COMMENT, true)->toArray();

        //echo "<PRE>"; echo java_inspect($comments); exit;
        // Look through all comments and gather information about them.
        $saveFormat = Java("com.aspose.words.SaveFormat");

        foreach ($comments as $comment)
        {
            if(isset($args[1]) && !empty($args[1])) {
                $authorName = $args[1];
                if (java_values($comment->getAuthor()->equals(authorName)))
                    $collectedComments->add($comment->getAuthor() . " " . $comment->getDateTime() . " " . $comment->toString($saveFormat->TEXT));
            } else {

                $collectedComments->add($comment->getAuthor() . " " . $comment->getDateTime() . " " . $comment->toString($saveFormat->TEXT));
            }

        }
        return $collectedComments;

    }

    public static function removeComments() {

        $args = func_get_args();
        $doc = $args[0];

        if(isset($args[1]) && !empty($args[1])) {
            $authorName = $args[1];
        }


        // Collect all comments in the document
        $nodeType = Java("com.aspose.words.NodeType");

        $comments = $doc->getChildNodes($nodeType->COMMENT, true);

        $comments_count = $comments->getCount();

        // Look through all comments and remove those written by the authorName author.
        $i = $comments_count;

        $i = $i - 1;

        while($i >= 0) {

            $comment = $comments->get($i);
            //echo "<PRE>"; echo java_inspect($comment); exit;
            if(isset($authorName)){
                if (java_values($comment->getAuthor()->equals($authorName)))
                    $comment->remove();
            } else {
                $comment->remove();
            }

            $i--;
        }

    }



}