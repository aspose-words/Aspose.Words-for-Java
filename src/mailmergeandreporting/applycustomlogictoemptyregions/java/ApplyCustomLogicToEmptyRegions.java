/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package mailmergeandreporting.applycustomlogictoemptyregions.java;

import javax.sql.rowset.RowSetMetaDataImpl;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;
import com.sun.rowset.CachedRowSetImpl;


public class ApplyCustomLogicToEmptyRegions
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/mailmergeandreporting/applycustomlogictoemptyregions/data/";

        //ExStart
        //ExId:CustomHandleRegionsMain
        //ExSummary:Shows how to handle unmerged regions after mail merge with user defined code.
        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        // Create a data source which has some data missing.
        // This will result in some regions that are merged and some that remain after executing mail merge.
        DataSet data = getDataSource();

        // Make sure that we have not set the removal of any unused regions as we will handle them manually.
        // We achieve this by removing the RemoveUnusedRegions flag from the cleanup options by using the AND and NOT bitwise operators.
        doc.getMailMerge().setCleanupOptions(doc.getMailMerge().getCleanupOptions() & ~MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);
        
        // Execute mail merge. Some regions will be merged with data, others left unmerged.
        doc.getMailMerge().executeWithRegions(data);

        // The regions which contained data now would of been merged. Any regions which had no data and were
        // not merged will still remain in the document.
        Document mergedDoc = doc.deepClone(); //ExSkip
        // Apply logic to each unused region left in the document using the logic set out in the handler.
        // The handler class must implement the IFieldMergingCallback interface.
        executeCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler());

        // Save the output document to disk.
        doc.save(dataDir + "TestFile.CustomLogicEmptyRegions1 Out.doc");
        //ExEnd

        // Reload the original merged document.
        doc = mergedDoc.deepClone();

        // Apply different logic to unused regions this time.
        executeCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler_MergeTable());

        doc.save(dataDir + "TestFile.CustomLogicEmptyRegions2 Out.doc");

        // Reload the original merged document.
        doc = mergedDoc.deepClone();

        //ExStart
        //ExId:HandleContactDetailsRegion
        //ExSummary:Shows how to specify only the ContactDetails region to be handled through the handler class.
        // Only handle the ContactDetails region in our handler.
        ArrayList<String> regions = new ArrayList<String>();
        regions.add("ContactDetails");
        executeCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler(), regions);
        //ExEnd

        doc.save(dataDir + "TestFile.CustomLogicEmptyRegions3 Out.doc");
    }

    //ExStart
    //ExId:CreateDataSourceFromDocumentRegionsMethod
    //ExSummary:Defines the method used to manually handle unmerged regions.
    /**
     * Returns a DataSet object containing a DataTable for the unmerged regions in the specified document.
     * If regionsList is null all regions found within the document are included. If an ArrayList instance is present
     * the only the regions specified in the list that are found in the document are added.
     */
    private static DataSet createDataSourceFromDocumentRegions(Document doc, ArrayList regionsList) throws Exception
    {
        final String TABLE_START_MARKER = "TableStart:";
        DataSet dataSet = new DataSet();
        String tableName = null;

        for (String fieldName : doc.getMailMerge().getFieldNames())
        {
            if (fieldName.contains(TABLE_START_MARKER))
            {
                tableName = fieldName.substring(TABLE_START_MARKER.length());
            }
            else if (tableName != null)
            {
                // Only add the table as a new DataTable if it doesn't already exists in the DataSet.
                if (dataSet.getTables().get(tableName) == null)
                {
                    ResultSet resultSet = createCachedRowSet(new String[] {fieldName});

                    // We only need to add the first field for the handler to be called for the fields in the region.
                    if (regionsList == null || regionsList.contains(tableName))
                    {
                        addRow(resultSet, new String[] {"FirstField"});
                    }

                    dataSet.getTables().add(new DataTable(resultSet, tableName));
                }
                tableName = null;
            }
        }

        return dataSet;
    }
    //ExEnd

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
    public static void executeCustomLogicOnEmptyRegions(Document doc, IFieldMergingCallback handler) throws Exception
    {
        executeCustomLogicOnEmptyRegions(doc, handler, null); // Pass null to handle all regions found in the document.
    }

    /**
     * Applies logic defined in the passed handler class to specific unused regions in the document as defined in regionsList. This allows to manually control
     * how unused regions are handled in the document.
     *
     * @param doc The document containing unused regions.
     * @param handler The handler which implements the IFieldMergingCallback interface and defines the logic to be applied to each unmerged region.
     * @param regionsList A list of strings corresponding to the region names that are to be handled by the supplied handler class. Other regions encountered will not be handled and are removed automatically.
     */
    public static void executeCustomLogicOnEmptyRegions(Document doc, IFieldMergingCallback handler, ArrayList regionsList) throws Exception
    {
        // Certain regions can be skipped from applying logic to by not adding the table name inside the CreateEmptyDataSource method.
        // Enable this cleanup option so any regions which are not handled by the user's logic are removed automatically.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);

        // Set the user's handler which is called for each unmerged region.
        doc.getMailMerge().setFieldMergingCallback(handler);

        // Execute mail merge using the dummy dataset. The dummy data source contains the table names of
        // each unmerged region in the document (excluding ones that the user may have specified to be skipped). This will allow the handler
        // to be called for each field in the unmerged regions.
        doc.getMailMerge().executeWithRegions(createDataSourceFromDocumentRegions(doc, regionsList));
    }

        /**
	     * A helper method that creates an empty Java disconnected ResultSet with the specified columns.
	     */
	    private static ResultSet createCachedRowSet(String[] columnNames) throws Exception
	    {
	        RowSetMetaDataImpl metaData = new RowSetMetaDataImpl();
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

	    /**
	     * A helper method that adds a new row with the specified values to a disconnected ResultSet.
	     */
	    private static void addRow(ResultSet resultSet, String[] values) throws Exception
	    {
	        resultSet.moveToInsertRow();

	        for (int i = 0; i < values.length; i++)
	            resultSet.updateString(i + 1, values[i]);

	        resultSet.insertRow();

	        // This "dance" is needed to add rows to the end of the result set properly.
	        // If I do something else then rows are either added at the front or the result
	        // set throws an exception about a deleted row during mail merge.
	        resultSet.moveToCurrentRow();
	        resultSet.last();
    }
    //ExEnd

    //ExStart
    //ExFor:FieldMergingArgsBase.TableName
    //ExId:EmptyRegionsHandler
    //ExSummary:Shows how to define custom logic in a handler implementing IFieldMergingCallback that is executed for unmerged regions in the document.
    public static class EmptyRegionsHandler implements IFieldMergingCallback
    {
        /**
         * Called for each field belonging to an unmerged region in the document.
         */
        public void fieldMerging(FieldMergingArgs args) throws Exception
        {
            // Change the text of each field of the ContactDetails region individually.
            if ("ContactDetails".equals(args.getTableName()))
            {
                // Set the text of the field based off the field name.
                if ("Name".equals(args.getFieldName()))
                    args.setText("(No details found)");
                else if ("Number".equals(args.getFieldName()))
                    args.setText("(N/A)");
            }

            // Remove the entire table of the Suppliers region. Also check if the previous paragraph
            // before the table is a heading paragraph and if so remove that too.
            if ("Suppliers".equals(args.getTableName()))
            {
                Table table = (Table)args.getField().getStart().getAncestor(NodeType.TABLE);

                // Check if the table has been removed from the document already.
                if (table.getParentNode() != null)
                {
                    // Try to find the paragraph which precedes the table before the table is removed from the document.
                    if (table.getPreviousSibling() != null && table.getPreviousSibling().getNodeType() == NodeType.PARAGRAPH)
                    {
                        Paragraph previousPara = (Paragraph)table.getPreviousSibling();
                        if (isHeadingParagraph(previousPara))
                            previousPara.remove();
                    }

                    table.remove();
                }
            }
        }

        /**
         * Returns true if the paragraph uses any Heading style e.g Heading 1 to Heading 9
         */
        private boolean isHeadingParagraph(Paragraph para) throws Exception
        {
            return (para.getParagraphFormat().getStyleIdentifier() >= StyleIdentifier.HEADING_1 && para.getParagraphFormat().getStyleIdentifier() <= StyleIdentifier.HEADING_9);
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception
        {
            // Do Nothing
        }
    }
    //ExEnd

    public static class EmptyRegionsHandler_MergeTable implements IFieldMergingCallback
    {
        /**
         * Called for each field belonging to an unmerged region in the document.
         */
        public void fieldMerging(FieldMergingArgs args) throws Exception
        {
            //ExStart
            //ExId:ContactDetailsCodeVariation
            //ExSummary:Shows how to replace an unused region with a message and remove extra paragraphs.
            // Store the parent paragraph of the current field for easy access.
            Paragraph parentParagraph = args.getField().getStart().getParentParagraph();

            // Define the logic to be used when the ContactDetails region is encountered.
            // The region is removed and replaced with a single line of text stating that there are no records.
            if ("ContactDetails".equals(args.getTableName()))
            {
                // Called for the first field encountered in a region. This can be used to execute logic on the first field
                // in the region without needing to hard code the field name. Often the base logic is applied to the first field and
                // different logic for other fields. The rest of the fields in the region will have a null FieldValue.
                if ("FirstField".equals(args.getFieldValue()))
                {
                    // Remove the "Name:" tag from the start of the paragraph
                    parentParagraph.getRange().replace("Name:", "", false, false);
                    // Set the text of the first field to display a message stating that there are no records.
                    args.setText("No records to display");
                }
                else
                {
                    // We have already inserted our message in the paragraph belonging to the first field. The other paragraphs in the region
                    // will still remain so we want to remove these. A check is added to ensure that the paragraph has not already been removed.
                    // which may happen if more than one field is included in a paragraph.
                    if (parentParagraph.getParentNode() != null)
                        parentParagraph.remove();
                }
            }
            //ExEnd

            //ExStart
            //ExFor:Cell.IsFirstCell
            //ExId:SuppliersCodeVariation
            //ExSummary:Shows how to merge all the parent cells of an unused region and display a message within the table.
            // Replace the unused region in the table with a "no records" message and merge all cells into one.
            if ("Suppliers".equals(args.getTableName()))
            {
                if ("FirstField".equals(args.getFieldValue()))
                {
                    // We will use the first paragraph to display our message. Make it centered within the table. The other fields in other cells
                    // within the table will be merged and won't be displayed so we don't need to do anything else with them.
                    parentParagraph.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
                    args.setText("No records to display");
                }

                // Merge the cells of the table together.
                Cell cell = (Cell)parentParagraph.getAncestor(NodeType.CELL);
                if (cell != null)
                {
                   if (cell.isFirstCell())
                       cell.getCellFormat().setHorizontalMerge(CellMerge.FIRST); // If this cell is the first cell in the table then the merge is started using "CellMerge.First".
                   else
                       cell.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS); // Otherwise the merge is continued using "CellMerge.Previous".
                }
            }
            //ExEnd
        }

        public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception
        {
            // Do Nothing
        }
    }

    /**
     * Returns the data used to merge the TestFile document.
     * This dataset purposely contains only rows for the StoreDetails region and only a select few for the child region.
     */
    private static DataSet getDataSource() throws Exception
    {
        // Create empty disconnected Java result sets.
        ResultSet storeDetailsResultSet = createCachedRowSet(new String[]{"ID", "Name", "Address", "City", "Country"});
        ResultSet contactDetailsResultSet = createCachedRowSet(new String[]{"ID", "Name", "Number"});

        // Create new Aspose.Words DataSet and DataTable objects to be used for mail merge.
        DataSet data = new DataSet();
        DataTable storeDetails = new DataTable(storeDetailsResultSet, "StoreDetails");
        DataTable contactDetails = new DataTable(contactDetailsResultSet, "ContactDetails");

        // Add the data to the tables.
        addRow(storeDetailsResultSet, new String[] {"0", "Hungry Coyote Import Store", "2732 Baker Blvd", "Eugene", "USA"});
        addRow(storeDetailsResultSet, new String[] {"1", "Great Lakes Food Market", "City Center Plaza, 516 Main St.", "San Francisco", "USA"});

        // Add data to the child table only for the first record.
        addRow(contactDetailsResultSet, new String[] {"0", "Thomas Hardy", "(206) 555-9857 ext 237"});
        addRow(contactDetailsResultSet, new String[] {"0", "Elizabeth Brown", "(206) 555-9857 ext 764"});

        // Include the tables in the DataSet.
        data.getTables().add(storeDetails);
        data.getTables().add(contactDetails);

        // Setup the relation between the parent table (StoreDetails) and the child table (ContactDetails).
        data.getRelations().add(new DataRelation("StoreDetailsToContactDetails",
                storeDetails,
                contactDetails,
                new String[] {"ID"},
                new String[] {"ID"}));

        return data;
    }
}