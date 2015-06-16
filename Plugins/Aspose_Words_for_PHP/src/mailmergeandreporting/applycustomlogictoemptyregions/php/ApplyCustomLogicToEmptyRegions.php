<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use com\aspose\words\IFieldMergingCallback as IFieldMergingCallback;
use com\aspose\words\DataRelation as DataRelation;

class ApplyCustomLogicToEmptyRegions {

    public static function main() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/mailmergeandreporting/applycustomlogictoemptyregions/data/";

        //ExStart
        //ExId:CustomHandleRegionsMain
        //ExSummary:Shows how to handle unmerged regions after mail merge with user defined code.
        // Open the document.
        $doc = new Java("com.aspose.words.Document", $dataDir . "TestFile.doc");

        // Create a data source which has some data missing.
        // This will result in some regions that are merged and some that remain after executing mail merge.
        $data = ApplyCustomLogicToEmptyRegions::getDataSource();

        // Make sure that we have not set the removal of any unused regions as we will handle them manually.
        // We achieve this by removing the RemoveUnusedRegions flag from the cleanup options by using the AND and NOT bitwise operators.
        $mailMergeCleanupOptions = Java("com.aspose.words.MailMergeCleanupOptions");
        $doc->getMailMerge()->setCleanupOptions($doc->getMailMerge()->getCleanupOptions() & ~$mailMergeCleanupOptions->REMOVE_UNUSED_REGIONS);

        // Execute mail merge. Some regions will be merged with data, others left unmerged.
        $doc->getMailMerge()->executeWithRegions($data);

        // The regions which contained data now would of been merged. Any regions which had no data and were
        // not merged will still remain in the document.
        $mergedDoc = $doc->deepClone(); //ExSkip
        // Apply logic to each unused region left in the document using the logic set out in the handler.
        // The handler class must implement the IFieldMergingCallback interface.
        ApplyCustomLogicToEmptyRegions::executeCustomLogicOnEmptyRegions($doc, new EmptyRegionsHandler());
        die("KKK");

        // Save the output document to disk.
        $doc->save($dataDir . "TestFile.CustomLogicEmptyRegions1 Out.doc");
        //ExEnd

        // Reload the original merged document.
        $doc = $mergedDoc->deepClone();

        // Apply different logic to unused regions this time.
        ApplyCustomLogicToEmptyRegions::executeCustomLogicOnEmptyRegions($doc, new EmptyRegionsHandler_MergeTable());

        $doc->save($dataDir . "TestFile.CustomLogicEmptyRegions2 Out.doc");

        // Reload the original merged document.
        $doc = $mergedDoc->deepClone();

        //ExStart
        //ExId:HandleContactDetailsRegion
        //ExSummary:Shows how to specify only the ContactDetails region to be handled through the handler class.
        // Only handle the ContactDetails region in our handler.
        $regions = new Java("java.util.ArrayList");
        $regions->add("ContactDetails");
        ApplyCustomLogicToEmptyRegions::executeCustomLogicOnEmptyRegions($doc, new EmptyRegionsHandler(), $regions);
        //ExEnd

        $doc->save($dataDir . "TestFile.CustomLogicEmptyRegions3 Out.doc");
    }

    /**
     * Returns the data used to merge the TestFile document.
     * This dataset purposely contains only rows for the StoreDetails region and only a select few for the child region.
     */

    public static function getDataSource() {

        // Create empty disconnected Java result sets.
        $storeDetailsResultSet = ApplyCustomLogicToEmptyRegions::createCachedRowSet(array("ID", "Name", "Address", "City", "Country"));
        $contactDetailsResultSet = ApplyCustomLogicToEmptyRegions::createCachedRowSet(array("ID", "Name", "Number"));

        // Create new Aspose.Words DataSet and DataTable objects to be used for mail merge.
        $data = new Java("com.aspose.words.DataSet");
        $storeDetails = new Java("com.aspose.words.DataTable" , $storeDetailsResultSet, "StoreDetails");
        $contactDetails = new Java("com.aspose.words.DataTable" , $contactDetailsResultSet, "ContactDetails");

        // Add the data to the tables.
        ApplyCustomLogicToEmptyRegions::addRow($storeDetailsResultSet, array("0", "Hungry Coyote Import Store", "2732 Baker Blvd", "Eugene", "USA"));
        ApplyCustomLogicToEmptyRegions::addRow($storeDetailsResultSet, array("1", "Great Lakes Food Market", "City Center Plaza, 516 Main St.", "San Francisco", "USA"));

        // Add data to the child table only for the first record.
        ApplyCustomLogicToEmptyRegions::addRow($contactDetailsResultSet, array("0", "Thomas Hardy", "(206) 555-9857 ext 237"));
        ApplyCustomLogicToEmptyRegions::addRow($contactDetailsResultSet, array("0", "Elizabeth Brown", "(206) 555-9857 ext 764"));

        // Include the tables in the DataSet.
        $data->getTables()->add($storeDetails);
        $data->getTables()->add($contactDetails);

        // Setup the relation between the parent table (StoreDetails) and the child table (ContactDetails).
        $data->getRelations()->add(new DataRelation("StoreDetailsToContactDetails",
            $storeDetails,
            $contactDetails,
            array("ID"),
            array("ID")));

        return $data;


    }

    public static function createCachedRowSet($columnNames) {

        $metaData = new RowSetMetaDataImpl();
        metaData.setColumnCount(columnNames.length);
        for (int i = 0; i < columnNames.length; i++)
        {
            metaData.setColumnName(i + 1, columnNames[i]);
            metaData.setColumnType(i + 1, java.sql.Types.VARCHAR);
        }

        CachedRowSetImpl rowSet = new CachedRowSetImpl();
        rowSet.setMetaData(metaData);

        return rowSet;

    }

    //ExStart
    //ExId:ExecuteCustomLogicOnEmptyRegionsMethod
    //ExSummary:Shows how to execute custom logic on unused regions using the specified handler.
    /**
     * Applies logic defined in the passed handler class to all unused regions in the document. This allows to manually control
     * how unused regions are handled in the document.
     *
     * @param doc The document containing unused regions.
     * @param handler The handler which implements the IFieldMergingCallback interface and defines the logic to be applied to each unmerged region.
     */


    public static function executeCustomLogicOnEmptyRegions() {

    }

}


class EmptyRegionsHandler extends IFieldMergingCallback {



    public function fieldMerging(FieldMergingArgs $args) {

        // Change the text of each field of the ContactDetails region individually.
        if (java_values($args->getTableName()) == "ContactDetails")
        {
            // Set the text of the field based off the field name.
            if (java_values($args->getFieldName()) == "Name")
                $args->setText("(No details found)");
            else if (java_values($args->getFieldName()) == "Number")
                $args->setText("(N/A)");
        }

        // Remove the entire table of the Suppliers region. Also check if the previous paragraph
        // before the table is a heading paragraph and if so remove that too.
        if (java_values($args->getTableName()) == "Suppliers")
        {
            $nodeType = Java("com.aspose.words.NodeType");
            $table = $args->getField()->getStart()->getAncestor($nodeType->TABLE);

                // Check if the table has been removed from the document already.
                if (java_values($table->getParentNode()) != null)
                {
                    // Try to find the paragraph which precedes the table before the table is removed from the document.
                    if (java_values($table->getPreviousSibling()) != null && java_values($table->getPreviousSibling()->getNodeType()) == $nodeType->PARAGRAPH)
                    {
                        $previousPara = $table->getPreviousSibling();
                        if ($this->isHeadingParagraph($previousPara))
                            $previousPara->remove();
                    }

                    $table->remove();
                }
            }

    }

    public function isHeadingParagraph($para) {

        $styleIdentifier = Java("com.aspose.words.StyleIdentifier");

        return ( java_values($para->getParagraphFormat()->getStyleIdentifier()) >= java_values($styleIdentifier->HEADING_1) && java_values($para->getParagraphFormat()->getStyleIdentifier()) <= java_values($styleIdentifier->HEADING_9));

    }

    public function imageFieldMerging($args) {

        // Do nothing
    }


}

class EmptyRegionsHandler_MergeTable extends IFieldMergingCallback {

    public function fieldMerging($args) {

        //ExStart
        //ExId:ContactDetailsCodeVariation
        //ExSummary:Shows how to replace an unused region with a message and remove extra paragraphs.
        // Store the parent paragraph of the current field for easy access.
        $parentParagraph = $args->getField()->getStart()->getParentParagraph();

        // Define the logic to be used when the ContactDetails region is encountered.
        // The region is removed and replaced with a single line of text stating that there are no records.
        if ("ContactDetails" == java_values($args->getTableName()))
        {
            // Called for the first field encountered in a region. This can be used to execute logic on the first field
            // in the region without needing to hard code the field name. Often the base logic is applied to the first field and
            // different logic for other fields. The rest of the fields in the region will have a null FieldValue.
            if ("FirstField" == java_values($args->getFieldValue()))
            {
                // Remove the "Name:" tag from the start of the paragraph
                $parentParagraph->getRange()->replace("Name:", "", false, false);
                // Set the text of the first field to display a message stating that there are no records.
                $args->setText("No records to display");
            }
            else
            {
                // We have already inserted our message in the paragraph belonging to the first field. The other paragraphs in the region
                // will still remain so we want to remove these. A check is added to ensure that the paragraph has not already been removed.
                // which may happen if more than one field is included in a paragraph.
                if (java_values($parentParagraph->getParentNode()) != null)
                    $parentParagraph->remove();
            }
        }
        //ExEnd

        //ExStart
        //ExFor:Cell.IsFirstCell
        //ExId:SuppliersCodeVariation
        //ExSummary:Shows how to merge all the parent cells of an unused region and display a message within the table.
        // Replace the unused region in the table with a "no records" message and merge all cells into one.
        if ("Suppliers" == java_values($args->getTableName()))
        {
            if ("FirstField" == java_values($args->getFieldValue()))
            {
                // We will use the first paragraph to display our message. Make it centered within the table. The other fields in other cells
                // within the table will be merged and won't be displayed so we don't need to do anything else with them.
                $paragraphAlignment = Java("com.aspose.words.ParagraphAlignment");
                $parentParagraph->getParagraphFormat()->setAlignment($paragraphAlignment->CENTER);
                $args->setText("No records to display");
            }

            // Merge the cells of the table together.
            $nodeType = Java("com.aspose.words.NodeType");
            $cell = $parentParagraph->getAncestor($nodeType->CELL);
                if (java_values($cell) != null)
                {
                    $cellMerge = Java("com.aspose.words.CellMerge");
                    if (java_values($cell->isFirstCell()))
                        $cell->getCellFormat()->setHorizontalMerge($cellMerge->FIRST); // If this cell is the first cell in the table then the merge is started using "CellMerge.First".
                    else
                        $cell->getCellFormat()->setHorizontalMerge($cellMerge->PREVIOUS); // Otherwise the merge is continued using "CellMerge.Previous".
                }
            }
        //ExEnd

    }

    public function imageFieldMerging($args) {

        // Do nothing
    }


}