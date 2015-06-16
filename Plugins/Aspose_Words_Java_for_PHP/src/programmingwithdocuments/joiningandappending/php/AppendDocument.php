<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
class AppendDocument {
    // The path to the documents directory.
    private static $gDataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/programmingwithdocuments/joiningandappending/data/";
    const BOOKMARK_PREFIX = "_SubDocumentEnd";
    // Field name of the NUMPAGES field.
    const NUM_PAGES_FIELD_NAME = "NUMPAGES";
    // Field name of the PAGEREF field.
    const PAGE_REF_FIELD_NAME = "PAGEREF";
    public static function main() {
        // Run each of the sample code snippets.
        AppendDocument::SimpleAppendDocument();
        AppendDocument::KeepSourceFormatting();
        AppendDocument::UseDestinationStyles();
        AppendDocument::JoinContinuous();
        AppendDocument::JoinNewPage();
        AppendDocument::RestartPageNumbering();
        AppendDocument::LinkHeadersFooters();
        AppendDocument::UnlinkHeadersFooters();
        AppendDocument::RemoveSourceHeadersFooters();
        AppendDocument::DifferentPageSetup();
        AppendDocument::ConvertNumPageFields();
        //AppendDocument::ListUseDestinationStyles();
        //AppendDocument::ListKeepSourceFormatting();
        //AppendDocument::KeepSourceTogether();
        //AppendDocument::BaseDocument();
        //AppendDocument::UpdatePageLayout();
    }
    public static function SimpleAppendDocument() {
        $dstDoc = new Java("com.aspose.words.Document",AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document",AppendDocument::$gDataDir . "TestFile.Source.doc");
        //ExStart
        //ExId:AppendDocument_SimpleAppend
        //ExSummary:Shows how to append a document to the end of another document using no additional options.
        // Append the source document to the destination document using no extra options.
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        //ExEnd
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.SimpleAppendDocument Out.docx");
    }
    public static function KeepSourceFormatting() {
        //ExStart
        //ExId:AppendDocument_KeepSourceFormatting
        //ExSummary:Shows how to append a document to another document while keeping the original formatting.
        // Load the documents to join.
        $dstDoc = new Java("com.aspose.words.Document",AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document",AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Keep the formatting from the source document when appending it to the destination document.
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        // Save the joined document to disk.
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.KeepSourceFormatting Out.docx");
        //ExEnd
    }
    public static function UseDestinationStyles() {
        //ExStart
        //ExId:AppendDocument_UseDestinationStyles
        //ExSummary:Shows how to append a document to another document using the formatting of the destination document.
        // Load the documents to join.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document" ,AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Append the source document using the styles of the destination document.
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->USE_DESTINATION_STYLES);
        // Save the joined document to disk.
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.UseDestinationStyles Out.doc");
        //ExEnd
    }
    public static function JoinContinuous() {
        //ExStart
        //ExId:AppendDocument_JoinContinuous
        //ExSummary:Shows how to append a document to another document so the content flows continuously.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Make the document appear straight after the destination documents content.
        $sectionStart = new Java("com.aspose.words.SectionStart");
        $srcDoc->getFirstSection()->getPageSetup()->setSectionStart($sectionStart->CONTINUOUS);
        // Append the source document using the original styles found in the source document.
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.JoinContinuous Out.doc");
        //ExEnd
    }
    public static function JoinNewPage() {
        //ExStart
        //ExId:AppendDocument_JoinNewPage
        //ExSummary:Shows how to append a document to another document so it starts on a new page.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Set the appended document to start on a new page.
        $sectionStart = new Java("com.aspose.words.SectionStart");
        $srcDoc->getFirstSection()->getPageSetup()->setSectionStart($sectionStart->NEW_PAGE);
        // Append the source document using the original styles found in the source document.
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.JoinNewPage Out.doc");
        //ExEnd
    }
    public static function RestartPageNumbering() {
        //ExStart
        //ExId:AppendDocument_RestartPageNumbering
        //ExSummary:Shows how to append a document to another document with page numbering restarted.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Set the appended document to appear on the next page.
        $sectionStart = new Java("com.aspose.words.SectionStart");
        $srcDoc->getFirstSection()->getPageSetup()->setSectionStart($sectionStart->NEW_PAGE);
        // Restart the page numbering for the document to be appended.
        $srcDoc->getFirstSection()->getPageSetup()->setRestartPageNumbering(true);
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.RestartPageNumbering Out.doc");
        //ExEnd
    }
    public static function LinkHeadersFooters() {
        //ExStart
        //ExFor:HeaderFooterCollection.LinkToPrevious(Boolean)
        //ExId:AppendDocument_LinkHeadersFooters
        //ExSummary:Shows how to append a document to another document and continue headers and footers from the destination document.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Set the appended document to appear on a new page.
        $sectionStart = new Java("com.aspose.words.SectionStart");
        $srcDoc->getFirstSection()->getPageSetup()->setSectionStart($sectionStart->NEW_PAGE);
        // Link the headers and footers in the source document to the previous section.
        // This will override any headers or footers already found in the source document.
        $srcDoc->getFirstSection()->getHeadersFooters()->linkToPrevious(true);
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.LinkHeadersFooters Out.doc");
        //ExEnd
    }
    public static function UnlinkHeadersFooters() {
        //ExStart
        //ExId:AppendDocument_UnlinkHeadersFooters
        //ExSummary:Shows how to append a document to another document so headers and footers do not continue from the destination document.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
        // Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        // from the destination document.
        $srcDoc->getFirstSection()->getHeadersFooters()->linkToPrevious(false);
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.UnlinkHeadersFooters Out.doc");
        //ExEnd
    }
    public static function RemoveSourceHeadersFooters() {
        //ExStart
        //ExId:AppendDocument_RemoveSourceHeadersFooters
        //ExSummary:Shows how to remove headers and footers from a document before appending it to another document.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Remove the headers and footers from each of the sections in the source document.
        $sections = $srcDoc->getSections()->toArray();
        foreach ($sections as $section)
        {
            $section->clearHeadersFooters();
        }
        // Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting
        // for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination
        // document. This should set to false to avoid this behaviour.
        $srcDoc->getFirstSection()->getHeadersFooters()->linkToPrevious(false);
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.RemoveSourceHeadersFooters Out.doc");
        //ExEnd
    }
    public static function DifferentPageSetup() {
        //ExStart
        //ExId:AppendDocument_DifferentPageSetup
        //ExSummary:Shows how to append a document to another document continuously which has different page settings.
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.SourcePageSetup.doc");
        // Set the source document to continue straight after the end of the destination document.
        // If some page setup settings are different then this may not work and the source document will appear
        // on a new page.
        $sectionStart = new Java("com.aspose.words.SectionStart");
        $srcDoc->getFirstSection()->getPageSetup()->setSectionStart($sectionStart->CONTINUOUS);
        // To ensure this does not happen when the source document has different page setup settings make sure the
        // settings are identical between the last section of the destination document.
        // If there are further continuous sections that follow on in the source document then this will need to be
        // repeated for those sections as well.
        $srcDoc->getFirstSection()->getPageSetup()->setPageWidth($dstDoc->getLastSection()->getPageSetup()->getPageWidth());
        $srcDoc->getFirstSection()->getPageSetup()->setPageHeight($dstDoc->getLastSection()->getPageSetup()->getPageHeight());
        $srcDoc->getFirstSection()->getPageSetup()->setOrientation($dstDoc->getLastSection()->getPageSetup()->getOrientation());
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.DifferentPageSetup Out.doc");
        //ExEnd
    }
    public static function ConvertNumPageFields() {
        $dstDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Destination.doc");
        $srcDoc = new Java("com.aspose.words.Document", AppendDocument::$gDataDir . "TestFile.Source.doc");
        // Restart the page numbering on the start of the source document.
        $srcDoc->getFirstSection()->getPageSetup()->setRestartPageNumbering(true);
        $srcDoc->getFirstSection()->getPageSetup()->setPageStartingNumber(1);
        // Append the source document to the end of the destination document.
        $importFormatMode = new Java("com.aspose.words.ImportFormatMode");
        $dstDoc->appendDocument($srcDoc, $importFormatMode->KEEP_SOURCE_FORMATTING);
        // After joining the documents the NUMPAGE fields will now display the total number of pages which
        // is undesired behaviour. Call this method to fix them by replacing them with PAGEREF fields.
        AppendDocument::convertNumPageFieldsToPageRef($dstDoc);
        // This needs to be called in order to update the new fields with page numbers.
        $dstDoc->updatePageLayout();
        $dstDoc->save(AppendDocument::$gDataDir . "TestFile.ConvertNumPageFields Out.doc");
    }
    /**
     * Replaces all NUMPAGES fields in the document with PAGEREF fields. The replacement field displays the total number
     * of pages in the sub document instead of the total pages in the document.
     *
     * @param doc The combined document to process.
     */
    public static function convertNumPageFieldsToPageRef($doc) {
        // This is the prefix for each bookmark which signals where page numbering restarts.
        // The underscore "_" at the start inserts this bookmark as hidden in MS Word.
        // Create a new DocumentBuilder which is used to insert the bookmarks and replacement fields.
        $builder = new Java("com.aspose.words.DocumentBuilder",$doc);
        // Defines the number of page restarts that have been encountered and therefore the number of "sub" documents
        // found within this document.
        $subDocumentCount = 0;
        // Iterate through all sections in the document.
        $sections = $doc->getSections()->toArray();
        foreach ($sections as $section)
        {
            // This section has it's page numbering restarted so we will treat this as the start of a sub document.
            // Any PAGENUM fields in this inner document must be converted to special PAGEREF fields to correct numbering.
            if ($section->getPageSetup()->getRestartPageNumbering())
            {
                // Don't do anything if this is the first section in the document. This part of the code will insert the bookmark marking
                // the end of the previous sub document so therefore it is not applicable for first section in the document.
                if (!java_values($section->equals($doc->getFirstSection())))
                {
                    // Get the previous section and the last node within the body of that section.
                    $prevSection = $section->getPreviousSibling();
                    $lastNode = $prevSection->getBody()->getLastChild();
                    // Use the DocumentBuilder to move to this node and insert the bookmark there.
                    // This bookmark represents the end of the sub document.
                    $builder->moveTo($lastNode);
                    $builder->startBookmark(AppendDocument::BOOKMARK_PREFIX . $subDocumentCount);
                    $builder->endBookmark(AppendDocument::BOOKMARK_PREFIX . $subDocumentCount);
                    // Increase the subdocument count to insert the correct bookmarks.
                    $subDocumentCount++;
                }
            }
            // The last section simply needs the ending bookmark to signal that it is the end of the current sub document.
            if (java_values($section->equals($doc->getLastSection())))
            {
                // Insert the bookmark at the end of the body of the last section.
                // Don't increase the count this time as we are just marking the end of the document.
                $lastNode = $doc->getLastSection()->getBody()->getLastChild();
                $builder->moveTo($lastNode);
                $builder->startBookmark(AppendDocument::BOOKMARK_PREFIX . $subDocumentCount);
                $builder->endBookmark(AppendDocument::BOOKMARK_PREFIX . $subDocumentCount);
            }
            // Iterate through each NUMPAGES field in the section and replace the field with a PAGEREF field referring to the bookmark of the current subdocument
            // This bookmark is positioned at the end of the sub document but does not exist yet. It is inserted when a section with restart page numbering or the last
            // section is encountered.
            $nodeType = new Java("com.aspose.words.NodeType");
            $nodes = $section->getChildNodes($nodeType->FIELD_START, true)->toArray();
            foreach ($nodes as $node)
            {
                $fieldStart = $node;
                $fieldType = new Java("com.aspose.words.FieldType");
                if (java_values($fieldStart->getFieldType()) == java_values($fieldType->FIELD_NUM_PAGES))
                {
                    // Get the field code.
                    $fieldCode = AppendDocument::getFieldCode($fieldStart);
                    // Since the NUMPAGES field does not take any additional parameters we can assume the remaining part of the field
                    // code after the fieldname are the switches. We will use these to help recreate the NUMPAGES field as a PAGEREF field.
                    $fieldCode = java_values($fieldCode);
                    $fieldSwitches = str_replace(AppendDocument::NUM_PAGES_FIELD_NAME,'',$fieldCode);
                    $fieldSwitches = trim($fieldSwitches);
                    // Inserting the new field directly at the FieldStart node of the original field will cause the new field to
                    // not pick up the formatting of the original field. To counter this insert the field just before the original field
                    $previousNode = $fieldStart->getPreviousSibling();
                    // If a previous run cannot be found then we are forced to use the FieldStart node.
                    if (java_values($previousNode) == null)
                        $previousNode = $fieldStart;
                    // Insert a PAGEREF field at the same position as the field.
                    $builder->moveTo($previousNode);
                    $message = java_values(AppendDocument::PAGE_REF_FIELD_NAME) . ' ' . java_values(AppendDocument::BOOKMARK_PREFIX) . java_values($subDocumentCount) . ' ' .  java_values($fieldSwitches);
                    $newField = $builder->insertField(java_values($message));
                    // The field will be inserted before the referenced node. Move the node before the field instead.
                    $previousNode->getParentNode()->insertBefore($previousNode, $newField->getStart());
                    // Remove the original NUMPAGES field from the document.
                    AppendDocument::removeField($fieldStart);
                }
            }
        }
    }
    //ExEnd
    //ExStart
    //ExFor:FieldStart
    //ExFor:FieldSeparator
    //ExFor:FieldEnd
    //ExId:AppendDocument_HelperFunctions
    //ExSummary:Provides some helper functions by the methods above
    /**
     * Retrieves the field code from a field.
     *
     * @param fieldStart The field start of the field which to gather the field code from.
     */
    private static function getFieldCode($fieldStart) {
        $builder = new Java("java.lang.StringBuilder");
        $nodeType = new Java("com.aspose.words.NodeType");
        for ($node = $fieldStart; java_values($node) != null && java_values($node->getNodeType()) != java_values($nodeType->FIELD_SEPARATOR) &&
        java_values($node->getNodeType()) != java_values($nodeType->FIELD_END) ; $node = $node->nextPreOrder($node->getDocument()))
        {
            $nodeType = new Java("com.aspose.words.NodeType");
            // Use text only of Run nodes to avoid duplication.
            if (java_values($node->getNodeType()) == java_values($nodeType->RUN))
                $builder->append($node->getText());
        }
        return $builder->toString();
    }
    private static function removeField($fieldStart) {
        $currentNode = $fieldStart;
        $isRemoving = true;
        while (java_values($currentNode) != null && $isRemoving)
        {
            $nodeType = new Java("com.aspose.words.NodeType");
            if (java_values($currentNode->getNodeType()) == java_values($nodeType->FIELD_END))
                $isRemoving = false;
            $nextNode = $currentNode->nextPreOrder($currentNode->getDocument());
            $currentNode->remove();
            $currentNode = $nextNode;
        }
    }
}