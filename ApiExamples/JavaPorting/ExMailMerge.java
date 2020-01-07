// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.net.System.Data.DataView;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import java.util.ArrayList;
import com.aspose.words.MailMergeRegionInfo;
import com.aspose.words.FieldQuote;
import com.aspose.words.FieldType;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.FieldMergeField;
import com.aspose.words.MappedDataFieldCollection;
import java.util.Iterator;
import java.util.Map;
import com.aspose.ms.System.msConsole;
import com.aspose.words.FieldAddressBlock;
import com.aspose.words.FieldGreetingLine;
import com.aspose.words.Field;
import com.aspose.words.IMailMergeCallback;
import com.aspose.words.net.System.Data.DataRow;
import org.testng.annotations.DataProvider;


@Test
public class ExMailMerge extends ApiExampleBase
{

    @Test
    public void executeDataTable() throws Exception
    {
        //ExStart
        //ExFor:Document
        //ExFor:MailMerge
        //ExFor:MailMerge.Execute(DataTable)
        //ExFor:MailMerge.Execute(DataRow)
        //ExFor:Document.MailMerge
        //ExSummary:Executes mail merge from an ADO.NET DataTable.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteDataTable.doc");

        // This example creates a table, but you would normally load table from a database
        DataTable table = new DataTable("Test");
        table.getColumns().add("CustomerName");
        table.getColumns().add("Address");
        table.getRows().add(new Object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        table.getRows().add(new Object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

        // Field values from the table are inserted into the mail merge fields found in the document
        doc.getMailMerge().execute(table);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.doc");

        // Open a fresh copy of our document to perform another mail merge
        doc = new Document(getMyDir() + "MailMerge.ExecuteDataTable.doc");

        // We can also source values for a mail merge from a single row in the table
        doc.getMailMerge().execute(table.getRows().get(1));

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.OneRow.doc");
        //ExEnd
    }

    @Test
    public void executeDataView() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(DataView)
        //ExSummary:Shows how to process a DataTable's data with a DataView before using it in a mail merge.
        // Create a new document and populate it with merge fields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Congratulations ");
        builder.insertField(" MERGEFIELD Name");
        builder.write(" for passing with a grade of ");
        builder.insertField(" MERGEFIELD Grade");

        // Create a data table that merge data will be sourced from 
        DataTable table = new DataTable("ExamResults");
        table.getColumns().add("Name");
        table.getColumns().add("Grade");
        table.getRows().add(new Object[] { "John Doe", "67" });
        table.getRows().add(new Object[] { "Jane Doe", "81" });
        table.getRows().add(new Object[] { "John Cardholder", "47" });
        table.getRows().add(new Object[] { "Joe Bloggs", "75" });

        // If we execute the mail merge on the table, a page will be created for each row in the order that it appears in the table
        // If we want to sort/filter rows without changing the table, we can use a data view
        DataView view = new DataView(table);
        view.setSort("Grade DESC");
        view.setRowFilter("Grade >= 50");

        // This mail merge will be executed on a view where the rows are sorted by the "Grade" column
        // and rows where the Grade values are below 50 are filtered out
        doc.getMailMerge().execute(view);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataView.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(DataSet)
    //ExSummary:Shows how to create a nested mail merge with regions with data from a data set with two related tables.
    @Test
    public void executeWithRegionsNested() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a MERGEFIELD with a value of "TableStart:Customers"
        // Normally, MERGEFIELDs specify the name of the column that they take row data from
        // "TableStart:Customers" however means that we are starting a mail merge region which belongs to a table called "Customers"
        // This will start the outer region and an "TableEnd:Customers" MERGEFIELD will signify its end 
        builder.insertField(" MERGEFIELD TableStart:Customers");

        // Data from rows of the "CustomerName" column of the "Customers" table will go in this MERGEFIELD
        builder.write("Orders for ");
        builder.insertField(" MERGEFIELD CustomerName");
        builder.write(":");

        // Create column headers for a table which will contain values from the second inner region
        builder.startTable();
        builder.insertCell();
        builder.write("Item");
        builder.insertCell();
        builder.write("Quantity");
        builder.endRow();

        // We have a second data table called "Orders", which has a many-to-one relationship with "Customers",
        // related by a "CustomerID" column
        // We will start this inner mail merge region over which the "Orders" table will preside,
        // which will iterate over the "Orders" table once for each merge of the outer "Customers" region,
        // picking up rows with the same CustomerID value
        builder.insertCell();
        builder.insertField(" MERGEFIELD TableStart:Orders");
        builder.insertField(" MERGEFIELD ItemName");
        builder.insertCell();
        builder.insertField(" MERGEFIELD Quantity");

        // End the inner region
        // One stipulation of using regions and tables is that the opening and closing of a mail merge region must
        // only happen over one row of a document's table  
        builder.insertField(" MERGEFIELD TableEnd:Orders");
        builder.endTable();

        // End the outer region
        builder.insertField(" MERGEFIELD TableEnd:Customers");

        DataSet customersAndOrders = createDataSet();
        doc.getMailMerge().executeWithRegions(customersAndOrders);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNested.docx");
    }

    /// <summary>
    /// Generates a data set which has two data tables named "Customers" and "Orders",
    /// with a one-to-many relationship between the former and latter on the "CustomerID" column.
    /// </summary>
    private static DataSet createDataSet()
    {
        // Create the outer mail merge
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
        tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

        // Create the table for the inner merge
        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
        tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
        tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

        // Add both tables to a data set
        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);

        // The "CustomerID" column, also the primary key of the customers table is the foreign key for the Orders table
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        return dataSet;
    }
    //ExEnd

    @Test
    public void executeWithRegionsConcurrent() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.ExecuteWithRegions(DataTable)
        //ExFor:MailMerge.ExecuteWithRegions(DataView)
        //ExSummary:Shows how to use regions to execute two separate mail merges in one document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // If we want to perform two consecutive mail merges on one document while taking data from two tables
        // that are related to each other in any way, we can separate the mail merges with regions
        // A mail merge region starts and ends with "TableStart:[RegionName]" and "TableEnd:[RegionName]" MERGEFIELDs
        // These regions are separate for unrelated data, while they can be nested for hierarchical data
        builder.writeln("\tCities: ");
        builder.insertField(" MERGEFIELD TableStart:Cities");
        builder.insertField(" MERGEFIELD Name");
        builder.insertField(" MERGEFIELD TableEnd:Cities");
        builder.insertParagraph();

        // Both MERGEFIELDs refer to a same column name, but values for each will come from different data tables
        builder.writeln("\tFruit: ");
        builder.insertField(" MERGEFIELD TableStart:Fruit");
        builder.insertField(" MERGEFIELD Name");
        builder.insertField(" MERGEFIELD TableEnd:Fruit");

        // Create two data tables that aren't linked or related in any way which we still want in the same document
        DataTable tableCities = new DataTable("Cities");
        tableCities.getColumns().add("Name");
        tableCities.getRows().add(new Object[] { "Washington" });
        tableCities.getRows().add(new Object[] { "London" });
        tableCities.getRows().add(new Object[] { "New York" });

        DataTable tableFruit = new DataTable("Fruit");
        tableFruit.getColumns().add("Name");
        tableFruit.getRows().add(new Object[] { "Cherry"});
        tableFruit.getRows().add(new Object[] { "Apple" });
        tableFruit.getRows().add(new Object[] { "Watermelon" });
        tableFruit.getRows().add(new Object[] { "Banana" });

        // We will need to run one mail merge per table
        // This mail merge will populate the MERGEFIELDs in the "Cities" range, while leaving the fields in "Fruit" empty
        doc.getMailMerge().executeWithRegions(tableCities);

        // Run a second merge for the "Fruit" table
        // We can use a DataView to sort or filter values of a DataTable before it is merged
        DataView dv = new DataView(tableFruit);
        dv.setSort("Name ASC");
        doc.getMailMerge().executeWithRegions(dv);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsConcurrent.docx");
        //ExEnd
    }

    @Test
    public void mailMergeRegionInfo() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.GetFieldNamesForRegion(System.String)
        //ExFor:MailMerge.GetFieldNamesForRegion(System.String,System.Int32)
        //ExFor:MailMerge.GetRegionsByName(System.String)
        //ExFor:MailMerge.RegionEndTag
        //ExFor:MailMerge.RegionStartTag
        //ExSummary:Shows how to create, list and read mail merge regions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // These tags, which go inside MERGEFIELDs, denote the strings that signify the starts and ends of mail merge regions 
        msAssert.areEqual("TableStart", doc.getMailMerge().getRegionStartTag());
        msAssert.areEqual("TableEnd", doc.getMailMerge().getRegionEndTag());

        // By using these tags, we will start and end a "MailMergeRegion1", which will contain MERGEFIELDs for two columns
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column2");
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // We can keep track of merge regions and their columns by looking at these collections
        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");
        msAssert.areEqual(1, regions.size());
        msAssert.areEqual("MailMergeRegion1", regions.get(0).getName());

        String[] mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1");
        msAssert.areEqual("Column1", mergeFieldNames[0]);
        msAssert.areEqual("Column2", mergeFieldNames[1]);

        // Insert a region with the same name as an existing region, which will make it a duplicate
        builder.insertParagraph(); // A single row/paragraph cannot be shared by multiple regions
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.insertField(" MERGEFIELD Column3");
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // Regions that share the same name are still accounted for and can be accessed by index
        regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");
        msAssert.areEqual(2, regions.size());

        mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1", 1);
        msAssert.areEqual("Column3", mergeFieldNames[0]);
        //ExEnd
    }

    //ExStart
    //ExFor:MailMerge.MergeDuplicateRegions
    //ExSummary:Shows how to work with duplicate mail merge regions.
    @Test (dataProvider = "mergeDuplicateRegionsDataProvider") //ExSkip
    public void mergeDuplicateRegions(boolean isMergeDuplicateRegions) throws Exception
    {
        // Create a document and table that we will merge
        Document doc = createSourceDocMergeDuplicateRegions();
        DataTable dataTable = createSourceTableMergeDuplicateRegions();

        // If this property is false, the first region will be merged
        // while the MERGEFIELDs of the second one will be left in the pre-merge state
        // To get both regions merged we would have to execute the mail merge twice, on a table of the same name
        // If this is set to true, both regions will be affected by the merge
        doc.getMailMerge().setMergeDuplicateRegions(isMergeDuplicateRegions);

        doc.getMailMerge().executeWithRegions(dataTable);
        doc.save(getArtifactsDir() + "MailMerge.MergeDuplicateRegions.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "mergeDuplicateRegionsDataProvider")
	public static Object[][] mergeDuplicateRegionsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    /// <summary>
    /// Return a document that contains two duplicate mail merge regions (sharing the same name in the "TableStart/End" tags).
    /// </summary>
    private static Document createSourceDocMergeDuplicateRegions() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD TableStart:MergeRegion");
        builder.insertField(" MERGEFIELD Column1");
        builder.insertField(" MERGEFIELD TableEnd:MergeRegion");
        builder.insertParagraph();

        builder.insertField(" MERGEFIELD TableStart:MergeRegion");
        builder.insertField(" MERGEFIELD Column2");
        builder.insertField(" MERGEFIELD TableEnd:MergeRegion");

        return doc;
    }

    /// <summary>
    /// Create a data table with one row and two columns.
    /// </summary>
    private static DataTable createSourceTableMergeDuplicateRegions()
    {
        DataTable dataTable = new DataTable("MergeRegion");
        dataTable.getColumns().add("Column1");
        dataTable.getColumns().add("Column2");
        dataTable.getRows().add(new Object[] { "Value 1", "Value 2" });

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.PreserveUnusedTags
    //ExSummary:Shows how to preserve the appearance of alternative mail merge tags that go unused during a mail merge. 
    @Test (dataProvider = "preserveUnusedTagsDataProvider") //ExSkip
    public void preserveUnusedTags(boolean preserveUnusedTags) throws Exception
    {
        // Create a document and table that we will merge
        Document doc = createSourceDocWithAlternativeMergeFields();
        DataTable dataTable = createSourceTablePreserveUnusedTags();

        // By default, alternative merge tags that can't receive data because the data source has no columns with their name
        // are converted to and left on display as MERGEFIELDs after the mail merge
        // We can preserve their original appearance setting this attribute to true
        doc.getMailMerge().setPreserveUnusedTags(preserveUnusedTags);
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "MailMerge.PreserveUnusedTags.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "preserveUnusedTagsDataProvider")
	public static Object[][] preserveUnusedTagsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    /// <summary>
    /// Create a document and add two tags that can accept mail merge data that are not the traditional MERGEFIELDs.
    /// </summary>
    private static Document createSourceDocWithAlternativeMergeFields() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("{{ Column1 }}");
        builder.writeln("{{ Column2 }}");

        // Our tags will only register as destinations for mail merge data if we set this to true
        doc.getMailMerge().setUseNonMergeFields(true);

        return doc;
    }

    /// <summary>
    /// Create a simple data table with one column.
    /// </summary>
    private static DataTable createSourceTablePreserveUnusedTags()
    {
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("Column1");
        dataTable.getRows().add(new Object[] { "Value1" });

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.MergeWholeDocument
    //ExSummary:Shows the relationship between mail merges with regions and field updating.
    @Test (dataProvider = "mergeWholeDocumentDataProvider") //ExSkip
    public void mergeWholeDocument(boolean isMergeWholeDocument) throws Exception
    {
        // Create a document and data table that will both be merged
        Document doc = createSourceDocMergeWholeDocument();
        DataTable dataTable = createSourceTableMergeWholeDocument();

        // A regular mail merge will update all fields in the document as part of the procedure,
        // which will happen if this property is set to true
        // Otherwise, a mail merge with regions will only update fields inside of the designated mail merge region
        doc.getMailMerge().setMergeWholeDocument(isMergeWholeDocument);
        doc.getMailMerge().executeWithRegions(dataTable);

        // If true, all fields in the document will be updated upon merging
        // In this case that property is false, so the first QUOTE field will not be updated and will not show a value,
        // but the second one inside the region designated by the data table name will show the correct value
        doc.save(getArtifactsDir() + "MailMerge.MergeWholeDocument.docx");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "mergeWholeDocumentDataProvider")
	public static Object[][] mergeWholeDocumentDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    /// <summary>
    /// Create a document with a QUOTE field outside and one more inside a mail merge region called "MyTable"
    /// </summary>
    private static Document createSourceDocMergeWholeDocument() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert QUOTE field outside of any mail merge regions
        FieldQuote field = (FieldQuote)builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("This QUOTE field is outside of the \"MyTable\" merge region.");

        // Start "MyTable" merge region
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD TableStart:MyTable");

        // Insert QUOTE field inside "MyTable" merge region
        field = (FieldQuote)builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("This QUOTE field is inside the \"MyTable\" merge region.");
        builder.insertParagraph();

        // Add a MERGEFIELD for a column in the data table, end the "MyTable" region and return the document
        builder.insertField(" MERGEFIELD MyColumn");
        builder.insertField(" MERGEFIELD TableEnd:MyTable");

        return doc;
    }

    /// <summary>
    /// Create a simple data table that will be used in a mail merge.
    /// </summary>
    private static DataTable createSourceTableMergeWholeDocument()
    {
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("MyColumn");
        dataTable.getRows().add(new Object[] { "MyValue" });

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.UseWholeParagraphAsRegion
    //ExSummary:Shows the relationship between mail merge regions and paragraphs.
    @Test //ExSkip
    public void useWholeParagraphAsRegion() throws Exception
    {
        // Create a document with 2 mail merge regions in one paragraph and a table to which can fill one of the regions during a mail merge
        Document doc = createSourceDocWithNestedMergeRegions();
        DataTable dataTable = createSourceTableDataTableForOneRegion();

        // By default, a paragraph can belong to no more than one mail merge region
        // Our document breaks this rule so executing a mail merge with regions now will cause an exception to be thrown
        Assert.assertTrue(doc.getMailMerge().getUseWholeParagraphAsRegion());
        
        // If we set this variable to false, paragraphs and mail merge regions are independent so we can safely run our mail merge
        doc.getMailMerge().setUseWholeParagraphAsRegion(false);
        doc.getMailMerge().executeWithRegions(dataTable);

        // Our first region is populated, while our second is safely displayed as unused all across one paragraph
        doc.save(getArtifactsDir() + "MailMerge.UseWholeParagraphAsRegion.docx");
    }

    /// <summary>
    /// Create a document with two mail merge regions sharing one paragraph.
    /// </summary>
    private static Document createSourceDocWithNestedMergeRegions() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Region 1: ");
        builder.insertField(" MERGEFIELD TableStart:MyTable");
        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column2");
        builder.insertField(" MERGEFIELD TableEnd:MyTable");

        builder.write(", Region 2: ");
        builder.insertField(" MERGEFIELD TableStart:MyOtherTable");
        builder.insertField(" MERGEFIELD TableEnd:MyOtherTable");

        return doc;
    }

    /// <summary>
    /// Create a data table that can populate one region during a mail merge.
    /// </summary>
    private static DataTable createSourceTableDataTableForOneRegion()
    {
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("Column1");
        dataTable.getColumns().add("Column2");
        dataTable.getRows().add(new Object[] { "Value 1", "Value 2" });

        return dataTable;
    }
    //ExEnd

    @Test
    public void trimWhiteSpaces() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.TrimWhitespaces
        //ExSummary:Shows how to trimmed whitespaces from mail merge values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD field", null);

        doc.getMailMerge().setTrimWhitespaces(true);
        doc.getMailMerge().execute(new String[] { "field" }, new Object[] { " first line\rsecond line\rthird line " });

        msAssert.areEqual("first line\rsecond line\rthird line\f", doc.getText());
        //ExEnd
    }

    @Test
    public void mailMergeGetFieldNames() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.GetFieldNames
        //ExSummary:Shows how to get names of all merge fields in a document.
        String[] fieldNames = doc.getMailMerge().getFieldNames();
        //ExEnd
    }

    @Test
    public void deleteFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.DeleteFields
        //ExSummary:Shows how to delete all merge fields from a document without executing mail merge.
        doc.getMailMerge().deleteFields();
        //ExEnd
    }

    @Test
    public void removeContainingFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to instruct the mail merge engine to remove any containing fields from around a merge field during mail merge.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS);
        //ExEnd
    }

    @Test
    public void removeUnusedFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to automatically remove unmerged merge fields during mail merge.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);
        //ExEnd
    }

    @Test
    public void removeEmptyParagraphs() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        //ExEnd
    }

    @Test (enabled = false, description = "WORDSNET-17733", dataProvider = "removeColonBetweenEmptyMergeFieldsDataProvider")
    public void removeColonBetweenEmptyMergeFields(String punctuationMark,
        boolean isCleanupParagraphsWithPunctuationMarks, String resultText) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.CleanupParagraphsWithPunctuationMarks
        //ExSummary:Shows how to remove paragraphs with punctuation marks after mail merge operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldMergeField mergeFieldOption1 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_1");
        mergeFieldOption1.setFieldName("Option_1");

        // Here is the complete list of cleanable punctuation marks:
        // !
        // ,
        // .
        // :
        // ;
        // ?
        // ¡
        // ¿
        builder.write(punctuationMark);

        FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_2");
        mergeFieldOption2.setFieldName("Option_2");

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        // The default value of the option is true which means that the behaviour was changed to mimic MS Word
        // If you rely on the old behavior are able to revert it by setting the option to false
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(isCleanupParagraphsWithPunctuationMarks);

        doc.getMailMerge().execute(new String[] { "Option_1", "Option_2" }, new Object[] { null, null });

        doc.save(getArtifactsDir() + "RemoveColonBetweenEmptyMergeFields.docx");
        //ExEnd

        msAssert.areEqual(resultText, doc.getText());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "removeColonBetweenEmptyMergeFieldsDataProvider")
	public static Object[][] removeColonBetweenEmptyMergeFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"!",  false,  ""},
			{", ",  false,  ""},
			{" . ",  false,  ""},
			{" :",  false,  ""},
			{"  ; ",  false,  ""},
			{" ?  ",  false,  ""},
			{"  ¡  ",  false,  ""},
			{"  ¿  ",  false,  ""},
			{"!",  true,  "!\f"},
			{", ",  true,  ", \f"},
			{" . ",  true,  " . \f"},
			{" :",  true,  " :\f"},
			{"  ; ",  true,  "  ; \f"},
			{" ?  ",  true,  " ?  \f"},
			{"  ¡  ",  true,  "  ¡  \f"},
			{"  ¿  ",  true,  "  ¿  \f"},
		};
	}

    //ExStart
    //ExFor:MailMerge.MappedDataFields
    //ExFor:MappedDataFieldCollection
    //ExFor:MappedDataFieldCollection.Add
    //ExFor:MappedDataFieldCollection.Clear
    //ExFor:MappedDataFieldCollection.ContainsKey(String)
    //ExFor:MappedDataFieldCollection.ContainsValue(String)
    //ExFor:MappedDataFieldCollection.Count
    //ExFor:MappedDataFieldCollection.GetEnumerator
    //ExFor:MappedDataFieldCollection.Item(String)
    //ExFor:MappedDataFieldCollection.Remove(String)
    //ExSummary:Shows how to map data columns and MERGEFIELDs with different names so the data is transferred between them during a mail merge.
    @Test //ExSkip
    public void mappedDataFieldCollection() throws Exception
    {
        // Create a document and table that we will merge
        Document doc = createSourceDocMappedDataFields();
        DataTable dataTable = createSourceTableMappedDataFields();
        
        // We have a column "Column2" in the data table that doesn't have a respective MERGEFIELD in the document
        // Also, we have a MERGEFIELD named "Column3" that does not exist as a column in the data source
        // If data from "Column2" is suitable for the "Column3" MERGEFIELD,
        // we can map that column name to the MERGEFIELD in the "MappedDataFields" key/value pair
        MappedDataFieldCollection mappedDataFields = doc.getMailMerge().getMappedDataFields();

        // A data source column name is linked to a MERGEFIELD name by adding an element like this
        mappedDataFields.add("MergeFieldName", "DataSourceColumnName");

        // So, values from "Column2" will now go into MERGEFIELDs named "Column3" as well as "Column2", if there are any
        mappedDataFields.add("Column3", "Column2");

        // The MERGEFIELD name is the "key" to the respective data source column name "value"
        msAssert.areEqual("DataSourceColumnName", mappedDataFields.get("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsValue("DataSourceColumnName"));

        // Now if we run this mail merge, the "Column3" MERGEFIELDs will take data from "Column2" of the table
        doc.getMailMerge().execute(dataTable);

        // We can count and iterate over the mapped columns/fields
        msAssert.areEqual(2, mappedDataFields.getCount());

        Iterator<Map.Entry<String, String>> enumerator = mappedDataFields.iterator();
        try /*JAVA: was using*/
    	{
            while (enumerator.hasNext())
                msConsole.writeLine(
                    $"Column named {enumerator.Current.Value} is mapped to MERGEFIELDs named {enumerator.Current.Key}");
    	}
        finally { if (enumerator != null) enumerator.close(); }

        // We can also remove some or all of the elements
        mappedDataFields.remove("MergeFieldName");
        Assert.assertFalse(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertFalse(mappedDataFields.containsValue("DataSourceColumnName"));

        mappedDataFields.clear();
        msAssert.areEqual(0, mappedDataFields.getCount());

        // Removing the mapped key/value pairs has no effect on the document because the merge was already done with them in place
        doc.save(getArtifactsDir() + "MailMerge.MappedDataFieldCollection.docx");
    }

    /// <summary>
    /// Create a document with 2 MERGEFIELDs, one of which does not have a corresponding column in the data table.
    /// </summary>
    private static Document createSourceDocMappedDataFields() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two MERGEFIELDs that will accept data from that table
        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column3");

        return doc;
    }

    /// <summary>
    /// Create a data table with 2 columns, one of which does not have a corresponding MERGEFIELD in our source document.
    /// </summary>
    private static DataTable createSourceTableMappedDataFields()
    {
        // Create a data table that will be used in a mail merge
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("Column1");
        dataTable.getColumns().add("Column2");
        dataTable.getRows().add(new Object[] { "Value1", "Value2" });

        return dataTable;
    }
    //ExEnd

    @Test
    public void getFieldNames() throws Exception
    {
        //ExStart
        //ExFor:FieldAddressBlock
        //ExFor:FieldAddressBlock.GetFieldNames
        //ExSummary:Shows how to get mail merge field names used by the field.
        Document doc = new Document(getMyDir() + "MailMerge.GetFieldNames.docx");

        String[] addressFieldsExpect =
        {
            "Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State",
            "Country or Region", "Postal Code"
        };

        FieldAddressBlock addressBlockField = (FieldAddressBlock) doc.getRange().getFields().get(0);
        String[] addressBlockFieldNames = addressBlockField.getFieldNames();
        //ExEnd

        msAssert.areEqual(addressFieldsExpect, addressBlockFieldNames);

        String[] greetingFieldsExpect = { "Courtesy Title", "Last Name" };

        FieldGreetingLine greetingLineField = (FieldGreetingLine) doc.getRange().getFields().get(1);
        String[] greetingLineFieldNames = greetingLineField.getFieldNames();

        msAssert.areEqual(greetingFieldsExpect, greetingLineFieldNames);
    }

    @Test
    public void useNonMergeFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.UseNonMergeFields
        //ExSummary:Shows how to perform mail merge into merge fields and into additional fields types.
        doc.getMailMerge().setUseNonMergeFields(true);
        //ExEnd
    }

    @Test (dataProvider = "mustacheTemplateSyntaxDataProvider")
    public void mustacheTemplateSyntax(boolean restoreTags, String sectionText) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("{{ testfield1 }}");
        builder.write("{{ testfield2 }}");
        builder.write("{{ testfield3 }}");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().setPreserveUnusedTags(restoreTags);

        DataTable table = new DataTable("Test");
        table.getColumns().add("testfield2");
        table.getRows().add("value 1");

        doc.getMailMerge().execute(table);

        String paraText = DocumentHelper.getParagraphText(doc, 0);

        msAssert.areEqual(sectionText, paraText);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "mustacheTemplateSyntaxDataProvider")
	public static Object[][] mustacheTemplateSyntaxDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true,  "{{ testfield1 }}value 1{{ testfield3 }}\f"},
			{false,  "\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f"},
		};
	}

    @Test
    public void testMailMergeGetRegionsHierarchy() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.GetRegionsHierarchy
        //ExFor:MailMergeRegionInfo
        //ExFor:MailMergeRegionInfo.Regions
        //ExFor:MailMergeRegionInfo.Name
        //ExFor:MailMergeRegionInfo.Fields
        //ExFor:MailMergeRegionInfo.StartField
        //ExFor:MailMergeRegionInfo.EndField
        //ExFor:MailMergeRegionInfo.Level
        //ExSummary:Shows how to get MailMergeRegionInfo and work with it.
        Document doc = new Document(getMyDir() + "MailMerge.TestRegionsHierarchy.doc");

        // Returns a full hierarchy of regions (with fields) available in the document
        MailMergeRegionInfo regionInfo = doc.getMailMerge().getRegionsHierarchy();

        // Get top regions in the document
        ArrayList<MailMergeRegionInfo> topRegions = regionInfo.getRegions();
        msAssert.areEqual(2, topRegions.size());
        msAssert.areEqual("Region1", topRegions.get(0).getName());
        msAssert.areEqual("Region2", topRegions.get(1).getName());
        msAssert.areEqual(1, topRegions.get(0).getLevel());
        msAssert.areEqual(1, topRegions.get(1).getLevel());

        // Get nested region in first top region
        ArrayList<MailMergeRegionInfo> nestedRegions = topRegions.get(0).getRegions();
        msAssert.areEqual(2, nestedRegions.size());
        msAssert.areEqual("NestedRegion1", nestedRegions.get(0).getName());
        msAssert.areEqual("NestedRegion2", nestedRegions.get(1).getName());
        msAssert.areEqual(2, nestedRegions.get(0).getLevel());
        msAssert.areEqual(2, nestedRegions.get(1).getLevel());

        // Get field list in first top region
        ArrayList<Field> fieldList = topRegions.get(0).getFields();
        msAssert.areEqual(4, fieldList.size());

        FieldMergeField startFieldMergeField = nestedRegions.get(0).getStartField();
        msAssert.areEqual("TableStart:NestedRegion1", startFieldMergeField.getFieldName());

        FieldMergeField endFieldMergeField = nestedRegions.get(0).getEndField();
        msAssert.areEqual("TableEnd:NestedRegion1", endFieldMergeField.getFieldName());
        //ExEnd
    }

    @Test
    public void testTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.MailMergeCallback
        //ExFor:IMailMergeCallback
        //ExFor:IMailMergeCallback.TagsReplaced
        //ExSummary:Shows how to define custom logic for handling events during mail merge.
        Document document = new Document();
        document.getMailMerge().setUseNonMergeFields(true);

        MailMergeCallbackStub mailMergeCallbackStub = new MailMergeCallbackStub();
        document.getMailMerge().setMailMergeCallback(mailMergeCallbackStub);

        document.getMailMerge().execute(new String[0], new Object[0]);

        msAssert.areEqual(1, mailMergeCallbackStub.getTagsReplacedCounter());
    }

    private static class MailMergeCallbackStub implements IMailMergeCallback
    {
        public void tagsReplaced()
        {
            setTagsReplacedCounter(getTagsReplacedCounter() + 1)/*Property++*/;
        }

        public int getTagsReplacedCounter() { return mTagsReplacedCounter; }; private void setTagsReplacedCounter(int value) { mTagsReplacedCounter = value; };

        private int mTagsReplacedCounter;
    }
    //ExEnd

    @Test (dataProvider = "getRegionsByNameDataProvider")
    public void getRegionsByName(String regionName) throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.RegionsByName.doc");

        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName(regionName);
        msAssert.areEqual(2, regions.size());

        for (MailMergeRegionInfo region : regions) msAssert.areEqual(regionName, region.getName());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "getRegionsByNameDataProvider")
	public static Object[][] getRegionsByNameDataProvider() throws Exception
	{
		return new Object[][]
		{
			{"Region1"},
			{"NestedRegion1"},
		};
	}

    @Test
    public void cleanupOptions() throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.CleanUp.docx");

        DataTable data = getDataTable();

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS);
        doc.getMailMerge().executeWithRegions(data);

        doc.save(getArtifactsDir() + "MailMerge.CleanUp.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "MailMerge.CleanUp.docx", getGoldsDir() + "MailMerge.CleanUp Gold.docx"));
    }

    /// <summary>
    /// Create DataTable and fill it with data.
    /// In real life this DataTable should be filled from a database.
    /// </summary>
    private static DataTable getDataTable()
    {
        DataTable dataTable = new DataTable("StudentCourse");
        dataTable.getColumns().add("CourseName");

        DataRow dataRowEmpty = dataTable.newRow();
        dataTable.getRows().add(dataRowEmpty);
        dataRowEmpty.set(0, "");

        for (int i = 0; i < 10; i++)
        {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Course " + i);
        }

        return dataTable;
    }

    @Test 
    public void unconditionalMergeFieldsAndRegions() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.UnconditionalMergeFieldsAndRegions
        //ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
        Document doc = new Document(getMyDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");

        // Merge fields and merge regions are merged regardless of the parent IF field's condition
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(true);

        // Fill the fields in the document with user data
        doc.getMailMerge().execute(
            new String[] { "FullName" },
            new Object[] { "James Bond" });

        doc.save(getArtifactsDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");
        //ExEnd
    }
}
