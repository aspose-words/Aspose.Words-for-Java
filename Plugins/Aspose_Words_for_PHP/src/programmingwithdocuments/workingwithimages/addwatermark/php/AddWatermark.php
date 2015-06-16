<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
class AddWatermark {
    public function main() {
        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/workingwithimages/addwatermark/data/";
        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");
        $this->insertWatermarkText($doc , "CONFIDENTIAL");
        $doc->save($dataDir . "TestFile Out.doc");
    }
    private function insertWatermarkText($doc, $watermarkText){
        // Create a watermark shape. This will be a WordArt shape.
        // You are free to try other shape types as watermarks.
        $shapeType = new Java("com.aspose.words.ShapeType");
        $watermark = new Java("com.aspose.words.Shape", $doc,  $shapeType->TEXT_PLAIN_TEXT);
        // Set up the text of the $watermark->
        $watermark->getTextPath()->setText($watermarkText);
        $watermark->getTextPath()->setFontFamily("Arial");
        $watermark->setWidth(500);
        $watermark->setHeight(100);
        // Text will be directed from the bottom-left to the top-right corner.
        $watermark->setRotation(-40);
        // Remove the following two lines if you need a solid black text.
        $color = new Java("java.awt.Color");
        $watermark->getFill()->setColor($color->GRAY); // Try LightGray to get more Word-style watermark
        $watermark->setStrokeColor($color->GRAY); // Try LightGray to get more Word-style watermark
        // Place the watermark in the page center.
        $relativeHorizontalPosition = new Java("com.aspose.words.RelativeHorizontalPosition");
        $watermark->setRelativeHorizontalPosition($relativeHorizontalPosition->PAGE);
        $watermark->setRelativeVerticalPosition($relativeHorizontalPosition->PAGE);
        $wrapType = new Java("com.aspose.words.WrapType");
        $watermark->setWrapType($wrapType->NONE);
        $verticalAlignment = new Java("com.aspose.words.VerticalAlignment");
        $watermark->setVerticalAlignment($verticalAlignment->CENTER);
        $horizontalAlignment = new Java("com.aspose.words.HorizontalAlignment");
        $watermark->setHorizontalAlignment($horizontalAlignment->CENTER);
        // Create a new paragraph and append the watermark to this paragraph.
        $watermarkPara = new Java("com.aspose.words.Paragraph", $doc);
        $watermarkPara->appendChild($watermark);
        $sects = $doc->getSections()->toArray();
        // Insert the watermark into all headers of each document section.
        foreach ($sects as $sect)
        {
            $headerFooterType = new Java("com.aspose.words.HeaderFooterType");
            // There could be up to three different headers in each section, since we want
            // the watermark to appear on all pages, insert into all headers.
            $this->insertWatermarkIntoHeader($watermarkPara, $sect, $headerFooterType->HEADER_PRIMARY);
            $this->insertWatermarkIntoHeader($watermarkPara, $sect, $headerFooterType->HEADER_FIRST);
            $this->insertWatermarkIntoHeader($watermarkPara, $sect, $headerFooterType->HEADER_EVEN);
        }
    }
    private function insertWatermarkIntoHeader($watermarkPara, $sect, $headerType) {
        $header = $sect->getHeadersFooters()->getByHeaderFooterType($headerType);
        if (java_values($header) == null)
        {
            // There is no header of the specified type in the current section, create it.
            $header = new Java("com.aspose.words.HeaderFooter",$sect->getDocument(),$headerType);
            $sect->getHeadersFooters()->add($header);
        }
        // Insert a clone of the watermark into the header.
        $header->appendChild($watermarkPara->deepClone(true));
    }
}