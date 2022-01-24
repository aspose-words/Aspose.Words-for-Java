package DocsExamples.Mail_Merge_And_Reporting.Complex_examples_and_helpers;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.MailMergeCleanupOptions;
import java.util.ArrayList;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.Table;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Cell;
import com.aspose.words.CellMerge;
import com.aspose.words.net.System.Data.DataRelation;

@Test
public class ApplyCustomLogicToEmptyRegions extends DocsExamplesBase
{
    @Test
    public void executeWithRegionsNestedCustom() throws Exception
    {
        //ExStart:ApplyCustomLogicToEmptyRegions
        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind suppliers.docx");

        // Create a data source which has some data missing.
        // This will result in some regions are merged, and some remain after executing mail merge
        DataSet data = getDataSource();

        // Ensure that we have not set the removal of any unused regions as we will handle them manually.
        // We achieve this by removing the RemoveUnusedRegions flag from the cleanup options using the AND and NOT bitwise operators.
        doc.getMailMerge().setCleanupOptions(doc.getMailMerge().getCleanupOptions() & ~MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);

        doc.getMailMerge().executeWithRegions(data);

        // Regions without data and not merged will remain in the document.
        Document mergedDoc = doc.deepClone(); //ExSkip
        
        // Apply logic to each unused region left in the document.
        executeCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler());

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNestedCustom_1.docx");

        doc = mergedDoc.deepClone();

        // Apply different logic to unused regions this time.
        executeCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandlerMergeTable());

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNestedCustom_2.docx");
        //ExEnd:ApplyCustomLogicToEmptyRegions
        
        doc = mergedDoc.deepClone();
        
        //ExStart:ContactDetails 
        ArrayList<String> regions = new ArrayList<String>();
        regions.add("ContactDetails");

        // Only handle the ContactDetails region in our handler.
        executeCustomLogicOnEmptyRegions(doc, new EmptyRegionsHandler(), regions);
        //ExEnd:ContactDetails

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNestedCustom_3.docx");
    }

    //ExStart:CreateDataSourceFromDocumentRegions
    /// <summary>
    /// Returns a DataSet object containing a DataTable for the unmerged regions in the specified document.
    /// If regionsList is null all regions found within the document are included. If an List instance is present,
    /// the only regions specified in the list found in the document are added.
    /// </summary>
    private DataSet createDataSourceFromDocumentRegions(Document doc, ArrayList<String> regionsList) throws Exception
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
                // Add the table name as a new DataTable if it doesn't already exist in the DataSet.
                if (dataSet.getTables().get(tableName) == null)
                {
                    DataTable table = new DataTable(tableName);
                    table.getColumns().add(fieldName);

                    // We only need to add the first field for the handler to be called for the region's fields.
                    if (regionsList == null || regionsList.contains(tableName))
                    {
                        table.getRows().add("FirstField");
                    }

                    dataSet.getTables().add(table);
                }

                tableName = null;
            }
        }

        return dataSet;
    }
    //ExEnd:CreateDataSourceFromDocumentRegions

    //ExStart:ExecuteCustomLogicOnEmptyRegions
    /// <summary>
    /// Applies logic defined in the passed handler class to all unused regions in the document.
    /// This allows controlling how unused regions are handled in the document manually.
    /// </summary>
    /// <param name="doc">The document containing unused regions.</param>
    /// <param name="handler">The handler which implements the IFieldMergingCallback interface
    /// and defines the logic to be applied to each unmerged region.</param>
    private void executeCustomLogicOnEmptyRegions(Document doc, IFieldMergingCallback handler) throws Exception
    {
        // Pass null to handle all regions found in the document.
        executeCustomLogicOnEmptyRegions(doc, handler, null); 
    }

    /// <summary>
    /// Applies logic defined in the passed handler class to specific unused regions in the document as defined in regionsList.
    /// This allows controlling how unused regions are handled in the document manually.
    /// </summary>
    /// <param name="doc">The document containing unused regions.</param>
    /// <param name="handler">The handler which implements the IFieldMergingCallback interface and defines
    /// the logic to be applied to each unmerged region.</param>
    /// <param name="regionsList">A list of strings corresponding to the region names that are to be handled
    /// by the supplied handler class. Other regions encountered will not be handled and are removed automatically.</param>
    private void executeCustomLogicOnEmptyRegions(Document doc, IFieldMergingCallback handler,
        ArrayList<String> regionsList) throws Exception
    {
        // Certain regions can be skipped from applying logic to by not adding
        // the table name inside the CreateEmptyDataSource method. Enable this cleanup option, so any regions
        // which are not handled by the user's logic are removed automatically.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);

        // Set the user's handler, which is called for each unmerged region.
        doc.getMailMerge().setFieldMergingCallback(handler);

        // Execute mail merge using the dummy dataset. The dummy data source contains each unmerged region's table names
        // in the document (excluding ones that the user may have specified to be skipped).
        // This will allow the handler to be called for each field in the unmerged regions.
        doc.getMailMerge().executeWithRegions(createDataSourceFromDocumentRegions(doc, regionsList));
    }
    //ExEnd:ExecuteCustomLogicOnEmptyRegions

    //ExStart:EmptyRegionsHandler 
    private static class EmptyRegionsHandler implements IFieldMergingCallback
    {
        /// <summary>
        /// Called for each field belonging to an unmerged region in the document.
        /// </summary>
        public void fieldMerging(FieldMergingArgs args)
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

            // Remove the entire table of the Suppliers region. Also, check if the previous paragraph
            // before the table is a heading paragraph and remove that.
            if ("Suppliers".equals(args.getTableName()))
            {
                Table table = (Table) args.getField().getStart().getAncestor(NodeType.TABLE);

                // Check if the table has been removed from the document already.
                if (table.getParentNode() != null)
                {
                    // Try to find the paragraph which precedes the table before the table is removed from the document.
                    if (table.getPreviousSibling() != null && table.getPreviousSibling().getNodeType() == NodeType.PARAGRAPH)
                    {
                        Paragraph previousPara = (Paragraph) table.getPreviousSibling();
                        if (isHeadingParagraph(previousPara))
                            previousPara.remove();
                    }

                    table.remove();
                }
            }
        }

        /// <summary>
        /// Returns true if the paragraph uses any Heading style, e.g., Heading 1 to Heading 9.
        /// </summary>
        private boolean isHeadingParagraph(Paragraph para)
        {
            return para.getParagraphFormat().getStyleIdentifier() >= StyleIdentifier.HEADING_1 &&
                   para.getParagraphFormat().getStyleIdentifier() <= StyleIdentifier.HEADING_9;
        }

        public void imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do nothing.
        }
    }
    //ExEnd:EmptyRegionsHandler 

    public static class EmptyRegionsHandlerMergeTable implements IFieldMergingCallback
    {
        /// <summary>
        /// Called for each field belonging to an unmerged region in the document.
        /// </summary>
        public void fieldMerging(FieldMergingArgs args) throws Exception
        {
            //ExStart:RemoveExtraParagraphs
            // Store the parent paragraph of the current field for easy access.
            Paragraph parentParagraph = args.getField().getStart().getParentParagraph();

            // Define the logic to be used when the ContactDetails region is encountered.
            // The region is removed and replaced with a single line of text stating that there are no records.
            if ("ContactDetails".equals(args.getTableName()))
            {
                // Called for the first field encountered in a region. This can be used to execute logic on the first field
                // in the region without needing to hard code the field name. Often the base logic is applied to the first field and 
                // different logic for other fields. The rest of the fields in the region will have a null FieldValue.
                if ("FirstField".equals((String) args.getFieldValue()))
                {
                    FindReplaceOptions options = new FindReplaceOptions();
                    // Remove the "Name:" tag from the start of the paragraph.
                    parentParagraph.getRange().replace("Name:", "", options);
                    // Set the text of the first field to display a message stating that there are no records.
                    args.setText("No records to display");
                }
                else
                {
                    // We have already inserted our message in the paragraph belonging to the first field.
                    // The other paragraphs in the region will remain, so we want to remove these.
                    // A check is added to ensure that the paragraph has not been removed,
                    // which may happen if more than one field is included in a paragraph.
                    if (parentParagraph.getParentNode() != null)
                        parentParagraph.remove();
                }
            }
            //ExEnd:RemoveExtraParagraphs

            //ExStart:MergeAllCells
            // Replace the unused region in the table with a "no records" message and merge all cells into one.
            if ("Suppliers".equals(args.getTableName()))
            {
                if ("FirstField".equals((String) args.getFieldValue()))
                {
                    // We will use the first paragraph to display our message. Make it centered within the table.
                    // The other fields in other cells within the table will be merged and won't be displayed,
                    // so we don't need to do anything else with them.
                    parentParagraph.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
                    args.setText("No records to display");
                }

                // Merge the cells of the table.
                Cell cell = (Cell) parentParagraph.getAncestor(NodeType.CELL);
                if (cell != null)
                {
                    cell.getCellFormat().setHorizontalMerge(cell.isFirstCell() ? CellMerge.FIRST : CellMerge.PREVIOUS);
                }
            }
            //ExEnd:MergeAllCells
        }

        public void imageFieldMerging(ImageFieldMergingArgs args)
        {
            // Do Nothing
        }
    }

    /// <summary>
    /// Returns the data used to merge the document.
    /// This dataset purposely contains only rows for the StoreDetails region and only a select few for the child region. 
    /// </summary>
    private DataSet getDataSource()
    {
        DataSet data = new DataSet();
        DataTable storeDetails = new DataTable("StoreDetails");
        DataTable contactDetails = new DataTable("ContactDetails");

        contactDetails.getColumns().add("ID");
        contactDetails.getColumns().add("Name");
        contactDetails.getColumns().add("Number");

        storeDetails.getColumns().add("ID");
        storeDetails.getColumns().add("Name");
        storeDetails.getColumns().add("Address");
        storeDetails.getColumns().add("City");
        storeDetails.getColumns().add("Country");

        storeDetails.getRows().add("0", "Hungry Coyote Import Store", "2732 Baker Blvd", "Eugene", "USA");
        storeDetails.getRows().add("1", "Great Lakes Food Market", "City Center Plaza, 516 Main St.", "San Francisco",
            "USA");

        contactDetails.getRows().add("0", "Thomas Hardy", "(206) 555-9857 ext 237");
        contactDetails.getRows().add("0", "Elizabeth Brown", "(206) 555-9857 ext 764");

        data.getTables().add(storeDetails);
        data.getTables().add(contactDetails);

        data.getRelations().add(storeDetails.getColumns().get("ID"), contactDetails.getColumns().get("ID"));

        return data;
    }

    private /*final*/ DataTable orderTable = null;
    private /*final*/ DataTable itemTable = null;

    private void disableForeignKeyConstraints(DataSet dataSet)
    {
        //ExStart:DisableForeignKeyConstraints
        dataSet.getRelations().add(new DataRelation("OrderToItem", orderTable.getColumns().get("Order_Id"),
            itemTable.getColumns().get("Order_Id"), false));
        //ExEnd:DisableForeignKeyConstraints
    }

    private void createDataRelation(DataSet dataSet)
    {
        //ExStart:CreateDataRelation
        dataSet.getRelations().add(new DataRelation("OrderToItem", orderTable.getColumns().get("Order_Id"),
            itemTable.getColumns().get("Order_Id")));
        //ExEnd:CreateDataRelation
    }
}
