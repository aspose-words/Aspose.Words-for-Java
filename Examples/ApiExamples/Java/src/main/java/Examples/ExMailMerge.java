package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.net.System.Data.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.sql.ResultSet;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

public class ExMailMerge extends ApiExampleBase {
    @Test
    public void executeArray() throws Exception {
        //ExStart
        //ExFor:MailMerge.Execute(String[], Object[])
        //ExFor:ContentDisposition
        //ExSummary:Shows how to perform a mail merge, and then save the document to the client browser.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD FullName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Company ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD City ");

        doc.getMailMerge().execute(new String[]{"FullName", "Company", "Address", "City"},
                new Object[]{"James Bond", "MI5 Headquarters", "Milbank", "London"});
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.mailMergeMatchesArray(new String[][]{new String[]{"James Bond", "MI5 Headquarters", "Milbank", "London"}}, doc, true);
    }

    @Test
    public void executeDataReader() throws Exception {
        //ExStart
        //ExFor:MailMerge.Execute(IDataReader)
        //ExSummary:Shows how to run a mail merge using data from a data reader.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Product:\t");
        builder.insertField(" MERGEFIELD ProductName");
        builder.write("\nSupplier:\t");
        builder.insertField(" MERGEFIELD CompanyName");
        builder.writeln();
        builder.insertField(" MERGEFIELD QuantityPerUnit");
        builder.write(" for $");
        builder.insertField(" MERGEFIELD UnitPrice");

        // "DocumentHelper.executeDataTable" is utility function that creates a connection, command,
        // executes the command and return the result in a DataTable.
        ResultSet resultSet = DocumentHelper.executeDataTable(
                "SELECT Products.ProductName, Suppliers.CompanyName, Products.QuantityPerUnit, " +
                        "{fn ROUND(Products.UnitPrice,2)} as UnitPrice " +
                        "FROM Products INNER JOIN Suppliers ON Products.SupplierID = Suppliers.SupplierID");
        DataTable dataTable = new DataTable(resultSet, "OrderDetails");
        IDataReader dataReader = new DataTableReader(dataTable);

        // Now we can take the data from the reader and use it in the mail merge.
        doc.getMailMerge().execute(dataReader);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataReader.docx");
        //ExEnd
    }

    @Test
    public void executeDataTable() throws Exception {
        //ExStart
        //ExFor:Document
        //ExFor:MailMerge
        //ExFor:MailMerge.Execute(DataTable)
        //ExFor:MailMerge.Execute(DataRow)
        //ExFor:Document.MailMerge
        //ExSummary:Shows how to execute a mail merge with data from a DataTable.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField(" MERGEFIELD CustomerName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");

        // This example creates a table, but you would normally load table from a database
        DataTable table = new DataTable("Test");
        table.getColumns().add("CustomerName");
        table.getColumns().add("Address");
        table.getRows().add("Thomas Hardy", "120 Hanover Sq., London");
        table.getRows().add("Paolo Accorti", "Via Monte Bianco 34, Torino");

        // Field values from the table are inserted into the mail merge fields found in the document
        doc.getMailMerge().execute(table);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.docx");

        // Create a copy of our document to perform another mail merge
        doc = new Document();
        builder = new DocumentBuilder(doc);
        builder.insertField(" MERGEFIELD CustomerName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");

        // We can also source values for a mail merge from a single row in the table
        doc.getMailMerge().execute(table.getRows().get(1));

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.OneRow.docx");
        //ExEnd
    }

    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(DataSet)
    //ExSummary:Shows how to create a nested mail merge with regions with data from a data set with two related tables.
    @Test
    public void executeWithRegionsNested() throws Exception {
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
    private static DataSet createDataSet() {
        // Create the outer mail merge
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(1, "John Doe");
        tableCustomers.getRows().add(2, "Jane Doe");

        // Create the table for the inner merge
        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(1, "Hawaiian", 2);
        tableOrders.getRows().add(2, "Pepperoni", 1);
        tableOrders.getRows().add(2, "Chicago", 1);

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
    public void executeWithRegionsConcurrent() throws Exception {
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
        tableCities.getRows().add("Washington");
        tableCities.getRows().add("London");
        tableCities.getRows().add("New York");

        DataTable tableFruit = new DataTable("Fruit");
        tableFruit.getColumns().add("Name");
        tableFruit.getRows().add("Cherry");
        tableFruit.getRows().add("Apple");
        tableFruit.getRows().add("Watermelon");
        tableFruit.getRows().add("Banana");

        // We will need to run one mail merge per table
        // This mail merge will populate the MERGEFIELDs in the "Cities" range, while leaving the fields in "Fruit" empty
        doc.getMailMerge().executeWithRegions(tableCities);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsConcurrent.docx");
        //ExEnd
    }

    @Test
    public void mailMergeRegionInfo() throws Exception {
        //ExStart
        //ExFor:MailMerge.GetFieldNamesForRegion(System.String)
        //ExFor:MailMerge.GetFieldNamesForRegion(System.String,System.Int32)
        //ExFor:MailMerge.GetRegionsByName(System.String)
        //ExFor:MailMerge.RegionEndTag
        //ExFor:MailMerge.RegionStartTag
        //ExFor:MailMergeRegionInfo.ParentRegion
        //ExSummary:Shows how to create, list and read mail merge regions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // These tags, which go inside MERGEFIELDs, denote the strings that signify the starts and ends of mail merge regions
        Assert.assertEquals(doc.getMailMerge().getRegionStartTag(), "TableStart");
        Assert.assertEquals(doc.getMailMerge().getRegionEndTag(), "TableEnd");

        // By using these tags, we will start and end a "MailMergeRegion1", which will contain MERGEFIELDs for two columns
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column2");
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // We can keep track of merge regions and their columns by looking at these collections
        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");
        Assert.assertEquals(regions.size(), 1);
        Assert.assertEquals(regions.get(0).getName(), "MailMergeRegion1");

        String[] mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1");
        Assert.assertEquals(mergeFieldNames[0], "Column1");
        Assert.assertEquals(mergeFieldNames[1], "Column2");

        // Insert a region with the same name inside the existing region, which will make it a parent.
        // Now a "Column2" field will be inside a new region.
        builder.moveToField(regions.get(0).getFields().get(1), false);
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.moveToField(regions.get(0).getFields().get(1), true);
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // Regions that share the same name are still accounted for and can be accessed by index
        regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");
        Assert.assertEquals(regions.size(), 2);
        // Check that the second region now has a parent region.
        Assert.assertEquals("MailMergeRegion1", regions.get(1).getParentRegion().getName());

        mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1", 1);
        Assert.assertEquals(mergeFieldNames[0], "Column2");
        //ExEnd
    }

    //ExStart
    //ExFor:MailMerge.MergeDuplicateRegions
    //ExSummary:Shows how to work with duplicate mail merge regions.
    @Test(dataProvider = "mergeDuplicateRegionsDataProvider") //ExSkip
    public void mergeDuplicateRegions(boolean isMergeDuplicateRegions) throws Exception {
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

    @DataProvider(name = "mergeDuplicateRegionsDataProvider")
    public static Object[][] mergeDuplicateRegionsDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    /// <summary>
    /// Return a document that contains two duplicate mail merge regions (sharing the same name in the "TableStart/End" tags).
    /// </summary>
    private static Document createSourceDocMergeDuplicateRegions() throws Exception {
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
    private static DataTable createSourceTableMergeDuplicateRegions() {
        DataTable dataTable = new DataTable("MergeRegion");
        dataTable.getColumns().add("Column1");
        dataTable.getColumns().add("Column2");
        dataTable.getRows().add("Value 1", "Value 2");

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.PreserveUnusedTags
    //ExSummary:Shows how to preserve the appearance of alternative mail merge tags that go unused during a mail merge. 
    @Test(dataProvider = "preserveUnusedTagsDataProvider") //ExSkip
    public void preserveUnusedTags(boolean doPreserveUnusedTags) throws Exception {
        // Create a document and table that we will merge
        Document doc = createSourceDocWithAlternativeMergeFields();
        DataTable dataTable = createSourceTablePreserveUnusedTags();

        // By default, alternative merge tags that can't receive data because the data source has no columns with their name
        // are converted to and left on display as MERGEFIELDs after the mail merge
        // We can preserve their original appearance setting this attribute to true
        doc.getMailMerge().setPreserveUnusedTags(doPreserveUnusedTags);
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "MailMerge.PreserveUnusedTags.docx");

        Assert.assertEquals(doc.getText().contains("{{ Column2 }}"), doPreserveUnusedTags);
    }

    @DataProvider(name = "preserveUnusedTagsDataProvider")
    public static Object[][] preserveUnusedTagsDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    /// <summary>
    /// Create a document and add two tags that can accept mail merge data that are not the traditional MERGEFIELDs.
    /// </summary>
    private static Document createSourceDocWithAlternativeMergeFields() throws Exception {
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
    private static DataTable createSourceTablePreserveUnusedTags() {
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("Column1");
        dataTable.getRows().add("Value1");

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.MergeWholeDocument
    //ExSummary:Shows the relationship between mail merges with regions and field updating.
    @Test(dataProvider = "mergeWholeDocumentDataProvider") //ExSkip
    public void mergeWholeDocument(boolean doMergeWholeDocument) throws Exception {
        // Create a document and data table that will both be merged
        Document doc = createSourceDocMergeWholeDocument();
        DataTable dataTable = createSourceTableMergeWholeDocument();

        // A regular mail merge will update all fields in the document as part of the procedure,
        // which will happen if this property is set to true
        // Otherwise, a mail merge with regions will only update fields
        // within a mail merge region which matches the name of the DataTable
        doc.getMailMerge().setMergeWholeDocument(doMergeWholeDocument);
        doc.getMailMerge().executeWithRegions(dataTable);

        // If true, all fields in the document will be updated upon merging
        // In this case that property is false, so the first QUOTE field will not be updated and will not show a value,
        // but the second one inside the region designated by the data table name will show the correct value
        doc.save(getArtifactsDir() + "MailMerge.MergeWholeDocument.docx");

        Assert.assertEquals(doMergeWholeDocument, doc.getText().contains("This QUOTE field is outside of the \"MyTable\" merge region."));
    }

    @DataProvider(name = "mergeWholeDocumentDataProvider")
    public static Object[][] mergeWholeDocumentDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    /// <summary>
    /// Create a document with a QUOTE field outside and one more inside a mail merge region called "MyTable"
    /// </summary>
    private static Document createSourceDocMergeWholeDocument() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert QUOTE field outside of any mail merge regions
        FieldQuote field = (FieldQuote) builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("This QUOTE field is outside of the \"MyTable\" merge region.");

        // Start "MyTable" merge region
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD TableStart:MyTable");

        // Insert QUOTE field inside "MyTable" merge region
        field = (FieldQuote) builder.insertField(FieldType.FIELD_QUOTE, true);
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
    private static DataTable createSourceTableMergeWholeDocument() {
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("MyColumn");
        dataTable.getRows().add("MyValue");

        return dataTable;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.UseWholeParagraphAsRegion
    //ExSummary:Shows the relationship between mail merge regions and paragraphs.
    @Test //ExSkip
    public void useWholeParagraphAsRegion() throws Exception {
        // Create a document with 2 mail merge regions in one paragraph and a table to which can fill one of the regions during a mail merge
        Document doc = createSourceDocWithNestedMergeRegions();
        DataTable dataTable = createSourceTableDataTableForOneRegion();

        // By default, a paragraph can belong to no more than one mail merge region
        // Our document breaks this rule so executing a mail merge with regions now will cause an exception to be thrown
        Assert.assertTrue(doc.getMailMerge().getUseWholeParagraphAsRegion());
        Assert.assertThrows(IllegalStateException.class, () -> doc.getMailMerge().executeWithRegions(dataTable));

        // If we set this variable to false, paragraphs and mail merge regions are independent so we can safely run our mail merge
        doc.getMailMerge().setUseWholeParagraphAsRegion(false);
        doc.getMailMerge().executeWithRegions(dataTable);

        // Our first region is populated, while our second is safely displayed as unused all across one paragraph
        doc.save(getArtifactsDir() + "MailMerge.UseWholeParagraphAsRegion.docx");
    }

    /// <summary>
    /// Create a document with two mail merge regions sharing one paragraph.
    /// </summary>
    private static Document createSourceDocWithNestedMergeRegions() throws Exception {
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
    private static DataTable createSourceTableDataTableForOneRegion() {
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("Column1");
        dataTable.getColumns().add("Column2");
        dataTable.getRows().add("Value 1", "Value 2");

        return dataTable;
    }
    //ExEnd

    @Test(dataProvider = "trimWhiteSpacesDataProvider")
    public void trimWhiteSpaces(boolean doTrimWhitespaces) throws Exception {
        //ExStart
        //ExFor:MailMerge.TrimWhitespaces
        //ExSummary:Shows how to trimmed whitespaces from mail merge values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD myMergeField", null);

        doc.getMailMerge().setTrimWhitespaces(doTrimWhitespaces);
        doc.getMailMerge().execute(new String[]{"myMergeField"}, new Object[]{"\t hello world! "});

        if (doTrimWhitespaces)
            Assert.assertEquals("hello world!\f", doc.getText());
        else
            Assert.assertEquals("\t hello world! \f", doc.getText());
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "trimWhiteSpacesDataProvider")
    public static Object[][] trimWhiteSpacesDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void mailMergeGetFieldNames() throws Exception {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.GetFieldNames
        //ExSummary:Shows how to get names of all merge fields in a document.
        String[] fieldNames = doc.getMailMerge().getFieldNames();
        //ExEnd
    }

    @Test
    public void deleteFields() throws Exception {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.DeleteFields
        //ExSummary:Shows how to delete all merge fields from a document without executing mail merge.
        doc.getMailMerge().deleteFields();
        //ExEnd
    }

    @Test
    public void removeContainingFields() throws Exception {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to instruct the mail merge engine to remove any containing fields from around a merge field during mail merge.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS);
        //ExEnd
    }

    @Test
    public void removeUnusedFields() throws Exception {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to automatically remove unmerged merge fields during mail merge.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);
        //ExEnd
    }

    @Test
    public void removeEmptyParagraphs() throws Exception {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        //ExEnd
    }

    @Test(enabled = false, description = "WORDSNET-17733", dataProvider = "removeColonBetweenEmptyMergeFieldsDataProvider")
    public void removeColonBetweenEmptyMergeFields(final String punctuationMark,
                                                   final boolean isCleanupParagraphsWithPunctuationMarks, final String resultText) throws Exception {
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
        // The default value of the option is true which means that the behavior was changed to mimic MS Word
        // If you rely on the old behavior are able to revert it by setting the option to false
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(isCleanupParagraphsWithPunctuationMarks);

        doc.getMailMerge().execute(new String[]{"Option_1", "Option_2"}, new Object[]{null, null});

        doc.save(getArtifactsDir() + "MailMerge.RemoveColonBetweenEmptyMergeFields.docx");
        //ExEnd

        Assert.assertEquals(doc.getText(), resultText);
    }

    @DataProvider(name = "removeColonBetweenEmptyMergeFieldsDataProvider")
    public static Object[][] removeColonBetweenEmptyMergeFieldsDataProvider() {
        return new Object[][]
                {
                        {"!", false, ""},
                        {", ", false, ""},
                        {" . ", false, ""},
                        {" :", false, ""},
                        {"  ; ", false, ""},
                        {" ?  ", false, ""},
                        {"  ¡  ", false, ""},
                        {"  ¿  ", false, ""},
                        {"!", true, "!\f"},
                        {", ", true, ", \f"},
                        {" . ", true, " . \f"},
                        {" :", true, " :\f"},
                        {"  ; ", true, "  ; \f"},
                        {" ?  ", true, " ?  \f"},
                        {"  ¡  ", true, "  ¡  \f"},
                        {"  ¿  ", true, "  ¿  \f"},
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
    public void mappedDataFieldCollection() throws Exception {
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
        Assert.assertEquals(mappedDataFields.get("MergeFieldName"), "DataSourceColumnName");
        Assert.assertTrue(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsValue("DataSourceColumnName"));

        // Now if we run this mail merge, the "Column3" MERGEFIELDs will take data from "Column2" of the table
        doc.getMailMerge().execute(dataTable);

        // We can count and iterate over the mapped columns/fields
        Assert.assertEquals(mappedDataFields.getCount(), 2);

        Iterator<Map.Entry<String, String>> enumerator = mappedDataFields.iterator();
        try {
            while (enumerator.hasNext()) {
                Map.Entry<String, String> dataField = enumerator.next();
                System.out.println(MessageFormat.format("Column named {0} is mapped to MERGEFIELDs named {1}", dataField.getValue(), dataField.getKey()));
            }
        } finally {
            if (enumerator != null) enumerator.remove();
        }

        // We can also remove some or all of the elements
        mappedDataFields.remove("MergeFieldName");
        Assert.assertFalse(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertFalse(mappedDataFields.containsValue("DataSourceColumnName"));

        mappedDataFields.clear();
        Assert.assertEquals(mappedDataFields.getCount(), 0);

        // Removing the mapped key/value pairs has no effect on the document because the merge was already done with them in place
        doc.save(getArtifactsDir() + "MailMerge.MappedDataFieldCollection.docx");
    }

    /// <summary>
    /// Create a document with 2 MERGEFIELDs, one of which does not have a corresponding column in the data table.
    /// </summary>
    private static Document createSourceDocMappedDataFields() throws Exception {
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
    private static DataTable createSourceTableMappedDataFields() {
        // Create a data table that will be used in a mail merge
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("Column1");
        dataTable.getColumns().add("Column2");
        dataTable.getRows().add("Value1", "Value2");

        return dataTable;
    }
    //ExEnd

    @Test
    public void getFieldNames() throws Exception {
        //ExStart
        //ExFor:FieldAddressBlock
        //ExFor:FieldAddressBlock.GetFieldNames
        //ExSummary:Shows how to get mail merge field names used by the field.
        Document doc = new Document(getMyDir() + "Field sample - ADDRESSBLOCK.docx");

        String[] addressFieldsExpect = {"Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State", "Country or Region", "Postal Code"};

        FieldAddressBlock addressBlockField = (FieldAddressBlock) doc.getRange().getFields().get(0);
        String[] addressBlockFieldNames = addressBlockField.getFieldNames();
        //ExEnd

        Assert.assertEquals(addressBlockFieldNames, addressFieldsExpect);

        String[] greetingFieldsExpect = {"Courtesy Title", "Last Name"};

        FieldGreetingLine greetingLineField = (FieldGreetingLine) doc.getRange().getFields().get(1);
        String[] greetingLineFieldNames = greetingLineField.getFieldNames();

        Assert.assertEquals(greetingLineFieldNames, greetingFieldsExpect);
    }

    @Test
    public void useNonMergeFields() throws Exception {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.UseNonMergeFields
        //ExSummary:Shows how to perform mail merge into merge fields and into additional fields types.
        doc.getMailMerge().setUseNonMergeFields(true);
        //ExEnd
    }

    /// <summary>
    /// Without TestCaseSource/TestCase because of some strange behavior when using long data.
    /// </summary>
    @Test
    public void mustacheTemplateSyntaxTrue() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("{{ testfield1 }}");
        builder.write("{{ testfield2 }}");
        builder.write("{{ testfield3 }}");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().setPreserveUnusedTags(true);

        DataTable table = new DataTable("Test");
        table.getColumns().add("testfield2");
        table.getRows().add("value 1");

        doc.getMailMerge().execute(table);

        String paraText = DocumentHelper.getParagraphText(doc, 0);

        Assert.assertEquals("{{ testfield1 }}value 1{{ testfield3 }}\f", paraText);
    }

    @Test
    public void mustacheTemplateSyntaxFalse() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("{{ testfield1 }}");
        builder.write("{{ testfield2 }}");
        builder.write("{{ testfield3 }}");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().setPreserveUnusedTags(false);

        DataTable table = new DataTable("Test");
        table.getColumns().add("testfield2");
        table.getRows().add("value 1");

        doc.getMailMerge().execute(table);

        String paraText = DocumentHelper.getParagraphText(doc, 0);

        Assert.assertEquals("\u0013MERGEFIELD \"testfield1\"\u0014«testfield1»\u0015value 1\u0013MERGEFIELD \"testfield3\"\u0014«testfield3»\u0015\f", paraText);
    }

    @Test
    public void testMailMergeGetRegionsHierarchy() throws Exception {
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
        Document doc = new Document(getMyDir() + "Mail merge regions.docx");

        // Returns a full hierarchy of regions (with fields) available in the document
        MailMergeRegionInfo regionInfo = doc.getMailMerge().getRegionsHierarchy();

        // Get top regions in the document
        ArrayList topRegions = regionInfo.getRegions();
        Assert.assertEquals(topRegions.size(), 2);
        Assert.assertEquals(((MailMergeRegionInfo) topRegions.get(0)).getName(), "Region1");
        Assert.assertEquals(((MailMergeRegionInfo) topRegions.get(1)).getName(), "Region2");
        Assert.assertEquals(((MailMergeRegionInfo) topRegions.get(0)).getLevel(), 1);
        Assert.assertEquals(((MailMergeRegionInfo) topRegions.get(1)).getLevel(), 1);

        // Get nested region in first top region
        ArrayList nestedRegions = ((MailMergeRegionInfo) topRegions.get(0)).getRegions();
        Assert.assertEquals(nestedRegions.size(), 2);
        Assert.assertEquals(((MailMergeRegionInfo) nestedRegions.get(0)).getName(), "NestedRegion1");
        Assert.assertEquals(((MailMergeRegionInfo) nestedRegions.get(1)).getName(), "NestedRegion2");
        Assert.assertEquals(((MailMergeRegionInfo) nestedRegions.get(0)).getLevel(), 2);
        Assert.assertEquals(((MailMergeRegionInfo) nestedRegions.get(1)).getLevel(), 2);

        // Get field list in first top region
        ArrayList fieldList = ((MailMergeRegionInfo) topRegions.get(0)).getFields();
        Assert.assertEquals(fieldList.size(), 4);

        FieldMergeField startFieldMergeField = ((MailMergeRegionInfo) nestedRegions.get(0)).getStartField();
        Assert.assertEquals(startFieldMergeField.getFieldName(), "TableStart:NestedRegion1");

        FieldMergeField endFieldMergeField = ((MailMergeRegionInfo) nestedRegions.get(0)).getEndField();
        Assert.assertEquals(endFieldMergeField.getFieldName(), "TableEnd:NestedRegion1");
        //ExEnd
    }

    @Test
    public void testTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption() throws Exception {
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

        Assert.assertEquals(mailMergeCallbackStub.getTagsReplacedCounter(), 1);
    }

    private static class MailMergeCallbackStub implements IMailMergeCallback {
        public void tagsReplaced() {
            mTagsReplacedCounter++;
        }

        public int getTagsReplacedCounter() {
            return mTagsReplacedCounter;
        }

        private int mTagsReplacedCounter;
    }
    //ExEnd

    @Test
    public void getRegionsByName() throws Exception {
        Document doc = new Document(getMyDir() + "Mail merge regions.docx");

        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("Region1");
        Assert.assertEquals(doc.getMailMerge().getRegionsByName("Region1").size(), 1);

        for (MailMergeRegionInfo region : regions) Assert.assertEquals(region.getName(), "Region1");

        regions = doc.getMailMerge().getRegionsByName("Region2");
        Assert.assertEquals(doc.getMailMerge().getRegionsByName("Region2").size(), 1);

        for (MailMergeRegionInfo region : regions) Assert.assertEquals(region.getName(), "Region2");

        regions = doc.getMailMerge().getRegionsByName("NestedRegion1");
        Assert.assertEquals(doc.getMailMerge().getRegionsByName("NestedRegion1").size(), 2);

        for (MailMergeRegionInfo region : regions) Assert.assertEquals(region.getName(), "NestedRegion1");
    }

    @Test
    public void cleanupOptions() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.startTable();
        builder.insertCell();
        builder.insertField(" MERGEFIELD  TableStart:StudentCourse ");
        builder.insertCell();
        builder.insertField(" MERGEFIELD  CourseName ");
        builder.insertCell();
        builder.insertField(" MERGEFIELD  TableEnd:StudentCourse ");
        builder.endTable();

        DataTable data = getDataTable();

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS);
        doc.getMailMerge().executeWithRegions(data);

        doc.save(getArtifactsDir() + "MailMerge.CleanupOptions.docx");

        Assert.assertTrue(DocumentHelper.compareDocs(getArtifactsDir() + "MailMerge.CleanupOptions.docx", getGoldsDir() + "MailMerge.CleanupOptions Gold.docx"));
    }

    /**
     * Create DataTable and fill it with data.
     * In real life this DataTable should be filled from a database.
     */
    private static DataTable getDataTable() {
        DataTable dataTable = new DataTable("StudentCourse");
        dataTable.getColumns().add("CourseName");

        DataRow dataRowEmpty = dataTable.newRow();
        dataTable.getRows().add(dataRowEmpty);
        dataRowEmpty.set(0, "");

        for (int i = 0; i < 10; i++) {
            DataRow datarow = dataTable.newRow();
            dataTable.getRows().add(datarow);
            datarow.set(0, "Course " + i);
        }

        return dataTable;
    }

    @Test(dataProvider = "unconditionalMergeFieldsAndRegionsDataProvider")
    public void unconditionalMergeFieldsAndRegions(boolean doCountAllMergeFields) throws Exception {
        //ExStart
        //ExFor:MailMerge.UnconditionalMergeFieldsAndRegions
        //ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD nested inside an IF field
        // Since the statement of the IF field is false, the result of the inner MERGEFIELD will not be displayed
        // and the MERGEFIELD will not receive any data during a mail merge
        FieldIf fieldIf = (FieldIf) builder.insertField(" IF 1 = 2 ");
        builder.moveTo(fieldIf.getSeparator());
        builder.insertField(" MERGEFIELD  FullName ");

        // We can still count MERGEFIELDs inside false-statement IF fields if we set this flag to true
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(doCountAllMergeFields);

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FullName");
        dataTable.getRows().add("James Bond");

        // Execute the mail merge
        doc.getMailMerge().execute(dataTable);

        // The result will not be visible in the document because the IF field is false, but the inner MERGEFIELD did indeed receive data
        doc.save(getArtifactsDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");

        if (doCountAllMergeFields)
            Assert.assertEquals("IF 1 = 2 \"James Bond\"", doc.getText().trim());
        else
            Assert.assertEquals("IF 1 = 2 \u0013 MERGEFIELD  FullName \u0014«FullName»", doc.getText().trim());
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "unconditionalMergeFieldsAndRegionsDataProvider")
    public static Object[][] unconditionalMergeFieldsAndRegionsDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "retainFirstSectionStartDataProvider")
    public void retainFirstSectionStart(boolean isRetainFirstSectionStart, int sectionStart, int expected) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD  FullName ");

        doc.getFirstSection().getPageSetup().setSectionStart(sectionStart);
        doc.getMailMerge().setRetainFirstSectionStart(isRetainFirstSectionStart);

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FullName");
        dataTable.getRows().add("James Bond");

        doc.getMailMerge().execute(dataTable);

        for (Section section : doc.getSections())
            Assert.assertEquals(expected, section.getPageSetup().getSectionStart());
    }

    @DataProvider(name = "retainFirstSectionStartDataProvider")
    public static Object[][] retainFirstSectionStartDataProvider() throws Exception {
        return new Object[][]
                {
                        {true, SectionStart.CONTINUOUS, SectionStart.CONTINUOUS},
                        {true, SectionStart.NEW_COLUMN, SectionStart.NEW_COLUMN},
                        {true, SectionStart.NEW_PAGE, SectionStart.NEW_PAGE},
                        {true, SectionStart.EVEN_PAGE, SectionStart.EVEN_PAGE},
                        {true, SectionStart.ODD_PAGE, SectionStart.ODD_PAGE},
                        {false, SectionStart.CONTINUOUS, SectionStart.NEW_PAGE},
                        {false, SectionStart.NEW_COLUMN, SectionStart.NEW_PAGE},
                        {false, SectionStart.NEW_PAGE, SectionStart.NEW_PAGE},
                        {false, SectionStart.EVEN_PAGE, SectionStart.EVEN_PAGE},
                        {false, SectionStart.ODD_PAGE, SectionStart.ODD_PAGE},
                };
    }

    @Test
    public void restartListsAtEachSection() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.RestartListsAtEachSection
        //ExSummary:Shows how to control whether or not list numbering is restarted at each section when mail merge is performed.
        Document doc = new Document(getMyDir() + "Section breaks with numbering.docx");

        doc.getMailMerge().setRestartListsAtEachSection(false);
        doc.getMailMerge().execute(new String[0], new Object[0]);

        doc.save(getArtifactsDir() + "MailMerge.RestartListsAtEachSection.pdf");
        //ExEnd
    }

    @Test
    public void removeLastEmptyParagraph() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.InsertHtml(String, HtmlInsertOptions)
        //ExSummary:Shows how to use options while inserting html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD Name ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD EMAIL ");
        builder.insertParagraph();

        // By default "DocumentBuilder.InsertHtml" inserts a HTML fragment that ends with a block-level HTML element,
        // it normally closes that block-level element and inserts a paragraph break.
        // As a result, a new empty paragraph appears after inserted document.
        // If we specify "HtmlInsertOptions.RemoveLastEmptyParagraph", those extra empty paragraphs will be removed.
        builder.moveToMergeField("NAME");
        builder.insertHtml("<p>John Smith</p>", HtmlInsertOptions.USE_BUILDER_FORMATTING | HtmlInsertOptions.REMOVE_LAST_EMPTY_PARAGRAPH);
        builder.moveToMergeField("EMAIL");
        builder.insertHtml("<p>jsmith@example.com</p>", HtmlInsertOptions.USE_BUILDER_FORMATTING);

        doc.save(getArtifactsDir() + "MailMerge.RemoveLastEmptyParagraph.docx");
        //ExEnd

        Assert.assertEquals(4, doc.getFirstSection().getBody().getParagraphs().getCount());
    }
}
