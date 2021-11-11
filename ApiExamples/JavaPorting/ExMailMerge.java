// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import com.aspose.words.ContentDisposition;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.net.System.Data.DataView;
import com.aspose.words.net.System.Data.DataSet;
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
import com.aspose.words.FieldIf;
import com.aspose.words.SectionStart;
import com.aspose.words.Section;
import com.aspose.ms.System.IO.File;
import com.aspose.words.MailMergeSettings;
import com.aspose.words.MailMergeMainDocumentType;
import com.aspose.words.MailMergeCheckErrors;
import com.aspose.words.MailMergeDataType;
import com.aspose.words.MailMergeDestination;
import com.aspose.words.Odso;
import com.aspose.words.OdsoDataSourceType;
import com.aspose.words.NodeType;
import com.aspose.words.OdsoFieldMapDataCollection;
import com.aspose.words.OdsoFieldMapData;
import com.aspose.words.OdsoFieldMappingType;
import com.aspose.words.OdsoRecipientDataCollection;
import com.aspose.words.OdsoRecipientData;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.ms.System.DateTime;
import com.aspose.words.FieldUpdateCultureSource;
import com.aspose.words.HtmlInsertOptions;
import org.testng.annotations.DataProvider;


@Test
public class ExMailMerge extends ApiExampleBase
{
    @Test
    public void executeArray() throws Exception
    {
        HttpResponse response = null;

        //ExStart
        //ExFor:MailMerge.Execute(String[], Object[])
        //ExFor:ContentDisposition
        //ExFor:Document.Save(HttpResponse,String,ContentDisposition,SaveOptions)
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

        doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "City" },
            new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "London" });

        // Send the document to the client browser.
        Assert.That(() => doc.Save(response, "Artifacts/MailMerge.ExecuteArray.docx", ContentDisposition.INLINE, null),
            Throws.<NullPointerException>TypeOf()); //Thrown because HttpResponse is null in the test.

        // We will need to close this response manually to ensure that we do not add any superfluous content to the document after saving.
        Assert.That(() => response.End(), Throws.<NullPointerException>TypeOf());
        //ExEnd

        doc = DocumentHelper.saveOpen(doc);

        TestUtil.mailMergeMatchesArray(new String[] { new String[] { "James Bond", "MI5 Headquarters", "Milbank", "London" } }, doc, true);
    }

    @Test (groups = "SkipMono")
    public void executeDataReader() throws Exception
    {
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

        // Create a connection string that points to the "Northwind" database file
        // in our local file system, open a connection, and set up an SQL query.
        String connectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + getDatabaseDir() + "Northwind.mdb";
        String query = 
            "SELECT Products.ProductName, Suppliers.CompanyName, Products.QuantityPerUnit, {fn ROUND(Products.UnitPrice,2)} as UnitPrice\n                FROM Products \n                INNER JOIN Suppliers \n                ON Products.SupplierID = Suppliers.SupplierID";

        OdbcConnection connection = new OdbcConnection();
        try /*JAVA: was using*/
        {
            connection.ConnectionString = connectionString;
            connection.Open();

            // Create an SQL command that will source data for our mail merge.
            // The names of the table's columns that this SELECT statement will return
            // will need to correspond to the merge fields we placed above.
            OdbcCommand command = connection.CreateCommand();
            command.CommandText = query;

            // This will run the command and store the data in the reader.
            OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

            // Take the data from the reader and use it in the mail merge.
            doc.getMailMerge().execute(reader);
        }
        finally { if (connection != null) connection.close(); }

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataReader.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "MailMerge.ExecuteDataReader.docx");

        TestUtil.mailMergeMatchesQueryResult(getDatabaseDir() + "Northwind.mdb", query, doc, true);
    }

    //ExStart
    //ExFor:MailMerge.ExecuteADO(Object)
    //ExSummary:Shows how to run a mail merge with data from an ADO dataset.
    @Test (groups = "SkipMono") //ExSkip
    public void executeADO() throws Exception
    {
        Document doc = createSourceDocADOMailMerge();

        // To work with ADO DataSets, we will need to add a reference to the Microsoft ActiveX Data Objects library,
        // which is included in the .NET distribution and stored in "adodb.dll".
        ADODB.Connection connection = new ADODB.Connection();

        // Create a connection string that points to the "Northwind" database file
        // in our local file system and open a connection.
        String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";
        connection.Open(connectionString);

        // Populate our DataSet by running an SQL command on our database.
        // The names of the columns in the result table will need to correspond
        // to the values of the MERGEFIELDS that will accommodate our data.
        final String COMMAND = "SELECT ProductName, QuantityPerUnit, UnitPrice FROM Products";

        ADODB.Recordset recordset = new ADODB.Recordset();
        recordset.Open(COMMAND, connection);

        // Execute the mail merge and save the document.
        doc.getMailMerge().ExecuteADO(recordset);
        doc.save(getArtifactsDir() + "MailMerge.ExecuteADO.docx");
        TestUtil.mailMergeMatchesQueryResult(getDatabaseDir() + "Northwind.mdb", COMMAND, doc, true); //ExSkip
    }

    /// <summary>
    /// Create a blank document and populate it with MERGEFIELDS that will accept data when a mail merge is executed.
    /// </summary>
    private static Document createSourceDocADOMailMerge() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Product:\t");
        builder.insertField(" MERGEFIELD ProductName");
        builder.writeln();
        builder.insertField(" MERGEFIELD QuantityPerUnit");
        builder.write(" for $");
        builder.insertField(" MERGEFIELD UnitPrice");

        return doc;
    }
    //ExEnd

    //ExStart
    //ExFor:MailMerge.ExecuteWithRegionsADO(Object,String)
    //ExSummary:Shows how to run a mail merge with multiple regions, compiled with data from an ADO dataset.
    @Test (groups = "SkipMono") //ExSkip
    public void executeWithRegionsADO() throws Exception
    {
        Document doc = createSourceDocADOMailMergeWithRegions();

        // To work with ADO DataSets, we will need to add a reference to the Microsoft ActiveX Data Objects library,
        // which is included in the .NET distribution and stored in "adodb.dll".
        ADODB.Connection connection = new ADODB.Connection();

        // Create a connection string that points to the "Northwind" database file
        // in our local file system and open a connection.
        String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";
        connection.Open(connectionString);

        // Populate our DataSet by running an SQL command on our database.
        // The names of the columns in the result table will need to correspond
        // to the values of the MERGEFIELDS that will accommodate our data.
        String command = "SELECT FirstName, LastName, City FROM Employees";

        ADODB.Recordset recordset = new ADODB.Recordset();
        recordset.Open(command, connection);

        // Run a mail merge on just the first region, filling its MERGEFIELDS with data from the record set.
        doc.getMailMerge().ExecuteWithRegionsADO(recordset, "MergeRegion1");

        // Close the record set and reopen it with data from another SQL query.
        command = "SELECT * FROM Customers";

        recordset.Close();
        recordset.Open(command, connection);

        // Run a second mail merge on the second region and save the document.
        doc.getMailMerge().ExecuteWithRegionsADO(recordset, "MergeRegion2");

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsADO.docx");
        TestUtil.mailMergeMatchesQueryResultMultiple(getDatabaseDir() + "Northwind.mdb", new String[] { "SELECT FirstName, LastName, City FROM Employees", "SELECT ContactName, Address, City FROM Customers" }, new Document(getArtifactsDir() + "MailMerge.ExecuteWithRegionsADO.docx"), false); //ExSkip
    }

    /// <summary>
    /// Create a document with two mail merge regions.
    /// </summary>
    private static Document createSourceDocADOMailMergeWithRegions() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("\tEmployees: ");
        builder.insertField(" MERGEFIELD TableStart:MergeRegion1");
        builder.insertField(" MERGEFIELD FirstName");
        builder.write(", ");
        builder.insertField(" MERGEFIELD LastName");
        builder.write(", ");
        builder.insertField(" MERGEFIELD City");
        builder.insertField(" MERGEFIELD TableEnd:MergeRegion1");
        builder.insertParagraph();

        builder.writeln("\tCustomers: ");
        builder.insertField(" MERGEFIELD TableStart:MergeRegion2");
        builder.insertField(" MERGEFIELD ContactName");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Address");
        builder.write(", ");
        builder.insertField(" MERGEFIELD City");
        builder.insertField(" MERGEFIELD TableEnd:MergeRegion2");

        return doc;
    }
    //ExEnd

    //ExStart
    //ExFor:Document
    //ExFor:MailMerge
    //ExFor:MailMerge.Execute(DataTable)
    //ExFor:MailMerge.Execute(DataRow)
    //ExFor:Document.MailMerge
    //ExSummary:Shows how to execute a mail merge with data from a DataTable.
    @Test //ExSkip
    public void executeDataTable() throws Exception
    {
        DataTable table = new DataTable("Test");
        table.getColumns().add("CustomerName");
        table.getColumns().add("Address");
        table.getRows().add(new Object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        table.getRows().add(new Object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

        // Below are two ways of using a DataTable as the data source for a mail merge.
        // 1 -  Use the entire table for the mail merge to create one output mail merge document for every row in the table:
        Document doc = createSourceDocExecuteDataTable();

        doc.getMailMerge().execute(table);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.WholeTable.docx");

        // 2 -  Use one row of the table to create one output mail merge document:
        doc = createSourceDocExecuteDataTable();
        
        doc.getMailMerge().execute(table.getRows().get(1));

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataTable.OneRow.docx");
        testADODataTable(new Document(getArtifactsDir() + "MailMerge.ExecuteDataTable.WholeTable.docx"), new Document(getArtifactsDir() + "MailMerge.ExecuteDataTable.OneRow.docx"), table); //ExSkip
    }

    /// <summary>
    /// Creates a mail merge source document.
    /// </summary>
    private static Document createSourceDocExecuteDataTable() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD CustomerName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");

        return doc;
    }
    //ExEnd

    private void testADODataTable(Document docWholeTable, Document docOneRow, DataTable table)
    {
        TestUtil.mailMergeMatchesDataTable(table, docWholeTable, true);

        DataTable rowAsTable = new DataTable();
        rowAsTable.importRow(table.getRows().get(1));

        TestUtil.mailMergeMatchesDataTable(rowAsTable, docOneRow, true);
    }

    @Test
    public void executeDataView() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(DataView)
        //ExSummary:Shows how to edit mail merge data with a DataView.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Congratulations ");
        builder.insertField(" MERGEFIELD Name");
        builder.write(" for passing with a grade of ");
        builder.insertField(" MERGEFIELD Grade");

        // Create a data table that our mail merge will source data from.
        DataTable table = new DataTable("ExamResults");
        table.getColumns().add("Name");
        table.getColumns().add("Grade");
        table.getRows().add(new Object[] { "John Doe", "67" });
        table.getRows().add(new Object[] { "Jane Doe", "81" });
        table.getRows().add(new Object[] { "John Cardholder", "47" });
        table.getRows().add(new Object[] { "Joe Bloggs", "75" });

        // We can use a data view to alter the mail merge data without making changes to the data table itself.
        DataView view = new DataView(table);
        view.setSort("Grade DESC");
        view.setRowFilter("Grade >= 50");

        // Our data view sorts the entries in descending order along the "Grade" column
        // and filters out rows with values of less than 50 on that column.
        // Three out of the four rows fit those criteria so that the output document will contain three merge documents.
        doc.getMailMerge().execute(view);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteDataView.docx");
        //ExEnd

        TestUtil.mailMergeMatchesDataTable(view.toTable(), new Document(getArtifactsDir() + "MailMerge.ExecuteDataView.docx"), true);
    }

    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(DataSet)
    //ExSummary:Shows how to execute a nested mail merge with two merge regions and two data tables.
    @Test
    public void executeWithRegionsNested() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        // Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        // Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        builder.insertField(" MERGEFIELD TableStart:Customers");

        // This MERGEFIELD is inside the mail merge region of the "Customers" table.
        // When we execute the mail merge, this field will receive data from rows in a data source named "Customers".
        builder.write("Orders for ");
        builder.insertField(" MERGEFIELD CustomerName");
        builder.write(":");

        // Create column headers for a table that will contain values from a second inner region.
        builder.startTable();
        builder.insertCell();
        builder.write("Item");
        builder.insertCell();
        builder.write("Quantity");
        builder.endRow();

        // Create a second mail merge region inside the outer region for a table named "Orders".
        // The "Orders" table has a many-to-one relationship with the "Customers" table on the "CustomerID" column.
        builder.insertCell();
        builder.insertField(" MERGEFIELD TableStart:Orders");
        builder.insertField(" MERGEFIELD ItemName");
        builder.insertCell();
        builder.insertField(" MERGEFIELD Quantity");

        // End the inner region, and then end the outer region. The opening and closing of a mail merge region must
        // happen on the same row of a table.
        builder.insertField(" MERGEFIELD TableEnd:Orders");
        builder.endTable();

        builder.insertField(" MERGEFIELD TableEnd:Customers");

        // Create a dataset that contains the two tables with the required names and relationships.
        // Each merge document for each row of the "Customers" table of the outer merge region will perform its mail merge on the "Orders" table.
        // Each merge document will display all rows of the latter table whose "CustomerID" column values match the current "Customers" table row.
        DataSet customersAndOrders = createDataSet();
        doc.getMailMerge().executeWithRegions(customersAndOrders);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNested.docx");
        TestUtil.mailMergeMatchesDataSet(customersAndOrders, new Document(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNested.docx"), false); //ExSkip
    }

    /// <summary>
    /// Generates a data set that has two data tables named "Customers" and "Orders", with a one-to-many relationship on the "CustomerID" column.
    /// </summary>
    private static DataSet createDataSet()
    {
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
        tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
        tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
        tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
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
        // related to each other in any way, we can separate the mail merges with regions.
        // Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        // Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        // Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        // These regions are separate for unrelated data, while they can be nested for hierarchical data.
        builder.writeln("\tCities: ");
        builder.insertField(" MERGEFIELD TableStart:Cities");
        builder.insertField(" MERGEFIELD Name");
        builder.insertField(" MERGEFIELD TableEnd:Cities");
        builder.insertParagraph();

        // Both MERGEFIELDs refer to the same column name, but values for each will come from different data tables.
        builder.writeln("\tFruit: ");
        builder.insertField(" MERGEFIELD TableStart:Fruit");
        builder.insertField(" MERGEFIELD Name");
        builder.insertField(" MERGEFIELD TableEnd:Fruit");

        // Create two unrelated data tables.
        DataTable tableCities = new DataTable("Cities");
        tableCities.getColumns().add("Name");
        tableCities.getRows().add(new Object[] { "Washington" });
        tableCities.getRows().add(new Object[] { "London" });
        tableCities.getRows().add(new Object[] { "New York" });

        DataTable tableFruit = new DataTable("Fruit");
        tableFruit.getColumns().add("Name");
        tableFruit.getRows().add(new Object[] { "Cherry" });
        tableFruit.getRows().add(new Object[] { "Apple" });
        tableFruit.getRows().add(new Object[] { "Watermelon" });
        tableFruit.getRows().add(new Object[] { "Banana" });

        // We will need to run one mail merge per table. The first mail merge will populate the MERGEFIELDs
        // in the "Cities" range while leaving the fields the "Fruit" range unfilled.
        doc.getMailMerge().executeWithRegions(tableCities);

        // Run a second merge for the "Fruit" table, while using a data view
        // to sort the rows in ascending order on the "Name" column before the merge.
        DataView dv = new DataView(tableFruit);
        dv.setSort("Name ASC");
        doc.getMailMerge().executeWithRegions(dv);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsConcurrent.docx");
        //ExEnd

        DataSet dataSet = new DataSet();

        dataSet.getTables().add(tableCities);
        dataSet.getTables().add(tableFruit);

        TestUtil.mailMergeMatchesDataSet(dataSet, new Document(getArtifactsDir() + "MailMerge.ExecuteWithRegionsConcurrent.docx"), false);
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
        //ExFor:MailMergeRegionInfo.ParentRegion
        //ExSummary:Shows how to create, list, and read mail merge regions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // "TableStart" and "TableEnd" tags, which go inside MERGEFIELDs,
        // denote the strings that signify the starts and ends of mail merge regions.
        Assert.assertEquals("TableStart", doc.getMailMerge().getRegionStartTag());
        Assert.assertEquals("TableEnd", doc.getMailMerge().getRegionEndTag());

        // Use these tags to start and end a mail merge region named "MailMergeRegion1",
        // which will contain MERGEFIELDs for two columns.
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column2");
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // We can keep track of merge regions and their columns by looking at these collections.
        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");

        Assert.assertEquals(1, regions.size());
        Assert.assertEquals("MailMergeRegion1", regions.get(0).getName());

        String[] mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1");

        Assert.assertEquals("Column1", mergeFieldNames[0]);
        Assert.assertEquals("Column2", mergeFieldNames[1]);

        // Insert a region with the same name inside the existing region, which will make it a parent.
        // Now a "Column2" field will be inside a new region.
        builder.moveToField(regions.get(0).getFields().get(1), false); 
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.moveToField(regions.get(0).getFields().get(1), true);
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // If we look up the name of duplicate regions using the "GetRegionsByName" method,
        // it will return all such regions in a collection.
        regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");

        Assert.assertEquals(2, regions.size());
        // Check that the second region now has a parent region.
        Assert.assertEquals("MailMergeRegion1", regions.get(1).getParentRegion().getName());

        mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1", 1);

        Assert.assertEquals("Column2", mergeFieldNames[0]);
        //ExEnd
    }

    //ExStart
    //ExFor:MailMerge.MergeDuplicateRegions
    //ExSummary:Shows how to work with duplicate mail merge regions.
    @Test (dataProvider = "mergeDuplicateRegionsDataProvider") //ExSkip
    public void mergeDuplicateRegions(boolean mergeDuplicateRegions) throws Exception
    {
        Document doc = createSourceDocMergeDuplicateRegions();
        DataTable dataTable = createSourceTableMergeDuplicateRegions();

        // If we set the "MergeDuplicateRegions" property to "false", the mail merge will affect the first region,
        // while the MERGEFIELDs of the second one will be left in the pre-merge state.
        // To get both regions merged like that,
        // we would have to execute the mail merge twice on a table of the same name.
        // If we set the "MergeDuplicateRegions" property to "true", the mail merge will affect both regions.
        doc.getMailMerge().setMergeDuplicateRegions(mergeDuplicateRegions);

        doc.getMailMerge().executeWithRegions(dataTable);
        doc.save(getArtifactsDir() + "MailMerge.MergeDuplicateRegions.docx");
        testMergeDuplicateRegions(dataTable, doc, mergeDuplicateRegions); //ExSkip
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
    /// Returns a document that contains two duplicate mail merge regions (sharing the same name in the "TableStart/End" tags).
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
    /// Creates a data table with one row and two columns.
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

    private void testMergeDuplicateRegions(DataTable dataTable, Document doc, boolean isMergeDuplicateRegions)
    {
        if (isMergeDuplicateRegions) 
            TestUtil.mailMergeMatchesDataTable(dataTable, doc, true);
        else
        {
            dataTable.getColumns().remove("Column2");
            TestUtil.mailMergeMatchesDataTable(dataTable, doc, true);
        }
    }

    //ExStart
    //ExFor:MailMerge.PreserveUnusedTags
    //ExFor:MailMerge.UseNonMergeFields
    //ExSummary:Shows how to preserve the appearance of alternative mail merge tags that go unused during a mail merge. 
    @Test (dataProvider = "preserveUnusedTagsDataProvider") //ExSkip
    public void preserveUnusedTags(boolean preserveUnusedTags) throws Exception
    {
        Document doc = createSourceDocWithAlternativeMergeFields();
        DataTable dataTable = createSourceTablePreserveUnusedTags();

        // By default, a mail merge places data from each row of a table into MERGEFIELDs, which name columns in that table. 
        // Our document has no such fields, but it does have plaintext tags enclosed by curly braces.
        // If we set the "PreserveUnusedTags" flag to "true", we could treat these tags as MERGEFIELDs
        // to allow our mail merge to insert data from the data source at those tags.
        // If we set the "PreserveUnusedTags" flag to "false",
        // the mail merge will convert these tags to MERGEFIELDs and leave them unfilled.
        doc.getMailMerge().setPreserveUnusedTags(preserveUnusedTags);
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "MailMerge.PreserveUnusedTags.docx");

        // Our document has a tag for a column named "Column2", which does not exist in the table.
        // If we set the "PreserveUnusedTags" flag to "false", then the mail merge will convert this tag into a MERGEFIELD.
        Assert.assertEquals(doc.getText().contains("{{ Column2 }}"), preserveUnusedTags);

        if (preserveUnusedTags)
            Assert.AreEqual(0, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldMergeField));
        else
            Assert.AreEqual(1, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldMergeField));
        TestUtil.mailMergeMatchesDataTable(dataTable, doc, true); //ExSkip
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
    /// Create a document and add two plaintext tags that may act as MERGEFIELDs during a mail merge.
    /// </summary>
    private static Document createSourceDocWithAlternativeMergeFields() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("{{ Column1 }}");
        builder.writeln("{{ Column2 }}");

        // Our tags will register as destinations for mail merge data only if we set this to true.
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
    //ExSummary:Shows the relationship between mail merges with regions, and field updating.
    @Test (dataProvider = "mergeWholeDocumentDataProvider") //ExSkip
    public void mergeWholeDocument(boolean mergeWholeDocument) throws Exception
    {
        Document doc = createSourceDocMergeWholeDocument();
        DataTable dataTable = createSourceTableMergeWholeDocument();

        // If we set the "MergeWholeDocument" flag to "true",
        // the mail merge with regions will update every field in the document.
        // If we set the "MergeWholeDocument" flag to "false", the mail merge will only update fields
        // within the mail merge region whose name matches the name of the data source table.
        doc.getMailMerge().setMergeWholeDocument(mergeWholeDocument);
        doc.getMailMerge().executeWithRegions(dataTable);

        // The mail merge will only update the QUOTE field outside of the mail merge region
        // if we set the "MergeWholeDocument" flag to "true".
        doc.save(getArtifactsDir() + "MailMerge.MergeWholeDocument.docx");

        Assert.assertTrue(doc.getText().contains("This QUOTE field is inside the \"MyTable\" merge region."));
        Assert.assertEquals(mergeWholeDocument, 
            doc.getText().contains("This QUOTE field is outside of the \"MyTable\" merge region."));
        TestUtil.mailMergeMatchesDataTable(dataTable, doc, true); //ExSkip
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
    /// Create a document with a mail merge region that belongs to a data source named "MyTable".
    /// Insert one QUOTE field inside this region, and one more outside it.
    /// </summary>
    private static Document createSourceDocMergeWholeDocument() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldQuote field = (FieldQuote)builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("This QUOTE field is outside of the \"MyTable\" merge region.");

        builder.insertParagraph();
        builder.insertField(" MERGEFIELD TableStart:MyTable");

        field = (FieldQuote)builder.insertField(FieldType.FIELD_QUOTE, true);
        field.setText("This QUOTE field is inside the \"MyTable\" merge region.");
        builder.insertParagraph();

        builder.insertField(" MERGEFIELD MyColumn");
        builder.insertField(" MERGEFIELD TableEnd:MyTable");

        return doc;
    }

    /// <summary>
    /// Create a data table that will be used in a mail merge.
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
    @Test (dataProvider = "useWholeParagraphAsRegionDataProvider") //ExSkip
    public void useWholeParagraphAsRegion(boolean useWholeParagraphAsRegion) throws Exception
    {
        Document doc = createSourceDocWithNestedMergeRegions();
        DataTable dataTable = createSourceTableDataTableForOneRegion();

        // By default, a paragraph can belong to no more than one mail merge region.
        // The contents of our document do not meet these criteria.
        // If we set the "UseWholeParagraphAsRegion" flag to "true",
        // running a mail merge on this document will throw an exception.
        // If we set the "UseWholeParagraphAsRegion" flag to "false",
        // we will be able to execute a mail merge on this document.
        doc.getMailMerge().setUseWholeParagraphAsRegion(useWholeParagraphAsRegion);

        if (useWholeParagraphAsRegion)
            Assert.<IllegalStateException>Throws(() => doc.getMailMerge().executeWithRegions(dataTable));
        else
            doc.getMailMerge().executeWithRegions(dataTable);

        // The mail merge populates our first region while leaving the second region unused
        // since it is the region that breaks the rule.
        doc.save(getArtifactsDir() + "MailMerge.UseWholeParagraphAsRegion.docx");
        if (!useWholeParagraphAsRegion) //ExSkip
            TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMerge.UseWholeParagraphAsRegion.docx"), true); //ExSkip
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "useWholeParagraphAsRegionDataProvider")
	public static Object[][] useWholeParagraphAsRegionDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
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

    @Test (dataProvider = "trimWhiteSpacesDataProvider")
    public void trimWhiteSpaces(boolean trimWhitespaces) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.TrimWhitespaces
        //ExSummary:Shows how to trim whitespaces from values of a data source while executing a mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD myMergeField", null);

        doc.getMailMerge().setTrimWhitespaces(trimWhitespaces);
        doc.getMailMerge().execute(new String[] { "myMergeField" }, new Object[] { "\t hello world! " });

        Assert.assertEquals(trimWhitespaces ? "hello world!\f" : "\t hello world! \f", doc.getText());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "trimWhiteSpacesDataProvider")
	public static Object[][] trimWhiteSpacesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void mailMergeGetFieldNames() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.GetFieldNames
        //ExSummary:Shows how to get names of all merge fields in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD FirstName ");
        builder.write(" ");
        builder.insertField(" MERGEFIELD LastName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD City ");

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getColumns().add("City");
        dataTable.getRows().add(new Object[] { "John", "Doe", "New York" });
        dataTable.getRows().add(new Object[] { "Joe", "Bloggs", "Washington" });
        
        // For every MERGEFIELD name in the document, ensure that the data table contains a column
        // with the same name, and then execute the mail merge. 
        String[] fieldNames = doc.getMailMerge().getFieldNames();

        Assert.assertEquals(3, fieldNames.length);

        for (String fieldName : fieldNames)
            Assert.assertTrue(dataTable.getColumns().contains(fieldName));

        doc.getMailMerge().execute(dataTable);
        //ExEnd

        TestUtil.mailMergeMatchesDataTable(dataTable, doc, true);
    }

    @Test
    public void deleteFields() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.DeleteFields
        //ExSummary:Shows how to delete all MERGEFIELDs from a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField(" MERGEFIELD FirstName ");
        builder.write(" ");
        builder.insertField(" MERGEFIELD LastName ");
        builder.writeln(",");
        builder.writeln("Greetings!");

        Assert.assertEquals(
            "Dear \u0013 MERGEFIELD FirstName \u0014«FirstName»\u0015 \u0013 MERGEFIELD LastName \u0014«LastName»\u0015,\rGreetings!", 
            doc.getText().trim());

        doc.getMailMerge().deleteFields();

        Assert.assertEquals("Dear  ,\rGreetings!", doc.getText().trim());
        //ExEnd
    }

    @Test (dataProvider = "removeUnusedFieldsDataProvider")
    public void removeUnusedFields(/*MailMergeCleanupOptions*/int mailMergeCleanupOptions) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to automatically remove MERGEFIELDs that go unused during mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a document with MERGEFIELDs for three columns of a mail merge data source table,
        // and then create a table with only two columns whose names match our MERGEFIELDs.
        builder.insertField(" MERGEFIELD FirstName ");
        builder.write(" ");
        builder.insertField(" MERGEFIELD LastName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD City ");

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[] { "John", "Doe" });
        dataTable.getRows().add(new Object[] { "Joe", "Bloggs" });

        // Our third MERGEFIELD references a "City" column, which does not exist in our data source.
        // The mail merge will leave fields such as this intact in their pre-merge state.
        // Setting the "CleanupOptions" property to "RemoveUnusedFields" will remove any MERGEFIELDs
        // that go unused during a mail merge to clean up the merge documents.
        doc.getMailMerge().setCleanupOptions(mailMergeCleanupOptions);
        doc.getMailMerge().execute(dataTable);

        if (mailMergeCleanupOptions == MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS || 
            mailMergeCleanupOptions == MailMergeCleanupOptions.REMOVE_STATIC_FIELDS)
            Assert.assertEquals(0, doc.getRange().getFields().getCount());
        else
            Assert.assertEquals(2, doc.getRange().getFields().getCount());
        //ExEnd

        TestUtil.mailMergeMatchesDataTable(dataTable, doc, true);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "removeUnusedFieldsDataProvider")
	public static Object[][] removeUnusedFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{MailMergeCleanupOptions.NONE},
			{MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS},
			{MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS},
			{MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS},
			{MailMergeCleanupOptions.REMOVE_STATIC_FIELDS},
			{MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS},
			{MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS},
		};
	}

    @Test (dataProvider = "removeEmptyParagraphsDataProvider")
    public void removeEmptyParagraphs(/*MailMergeCleanupOptions*/int mailMergeCleanupOptions) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.CleanupOptions
        //ExFor:MailMergeCleanupOptions
        //ExSummary:Shows how to remove empty paragraphs that a mail merge may create from the merge output document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD TableStart:MyTable");
        builder.insertField(" MERGEFIELD FirstName ");
        builder.write(" ");
        builder.insertField(" MERGEFIELD LastName ");
        builder.insertField(" MERGEFIELD TableEnd:MyTable");

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[] { "John", "Doe" });
        dataTable.getRows().add(new Object[] { "", "" });
        dataTable.getRows().add(new Object[] { "Jane", "Doe" });

        doc.getMailMerge().setCleanupOptions(mailMergeCleanupOptions);
        doc.getMailMerge().executeWithRegions(dataTable);

        if (doc.getMailMerge().getCleanupOptions() == MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS) 
            Assert.assertEquals(
                "John Doe\r" +
                "Jane Doe", doc.getText().trim());
        else
            Assert.assertEquals(
                "John Doe\r" +
                " \r" +
                "Jane Doe", doc.getText().trim());
        //ExEnd

        TestUtil.mailMergeMatchesDataTable(dataTable, doc, false);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "removeEmptyParagraphsDataProvider")
	public static Object[][] removeEmptyParagraphsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{MailMergeCleanupOptions.NONE},
			{MailMergeCleanupOptions.REMOVE_CONTAINING_FIELDS},
			{MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS},
			{MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS},
			{MailMergeCleanupOptions.REMOVE_STATIC_FIELDS},
			{MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS},
			{MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS},
		};
	}

    @Test (enabled = false, description = "WORDSNET-17733", dataProvider = "removeColonBetweenEmptyMergeFieldsDataProvider")
    public void removeColonBetweenEmptyMergeFields(String punctuationMark,
        boolean cleanupParagraphsWithPunctuationMarks, String resultText) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.CleanupParagraphsWithPunctuationMarks
        //ExSummary:Shows how to remove paragraphs with punctuation marks after a mail merge operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        FieldMergeField mergeFieldOption1 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_1");
        mergeFieldOption1.setFieldName("Option_1");

        builder.write(punctuationMark);

        FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_2");
        mergeFieldOption2.setFieldName("Option_2");

        // Configure the "CleanupOptions" property to remove any empty paragraphs that this mail merge would create.
        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);

        // Setting the "CleanupParagraphsWithPunctuationMarks" property to "true" will also count paragraphs
        // with punctuation marks as empty and will get the mail merge operation to remove them as well.
        // Setting the "CleanupParagraphsWithPunctuationMarks" property to "false"
        // will remove empty paragraphs, but not ones with punctuation marks.
        // This is a list of punctuation marks that this property concerns: "!", ",", ".", ":", ";", "?", "¡", "¿".
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(cleanupParagraphsWithPunctuationMarks);

        doc.getMailMerge().execute(new String[] { "Option_1", "Option_2" }, new Object[] { null, null });

        doc.save(getArtifactsDir() + "MailMerge.RemoveColonBetweenEmptyMergeFields.docx");
        //ExEnd

        Assert.assertEquals(resultText, doc.getText());
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
        Document doc = createSourceDocMappedDataFields();
        DataTable dataTable = createSourceTableMappedDataFields();

        // The table has a column named "Column2", but there are no MERGEFIELDs with that name.
        // Also, we have a MERGEFIELD named "Column3", but the data source does not have a column with that name.
        // If data from "Column2" is suitable for the "Column3" MERGEFIELD,
        // we can map that column name to the MERGEFIELD in the "MappedDataFields" key/value pair.
        MappedDataFieldCollection mappedDataFields = doc.getMailMerge().getMappedDataFields();

        // We can link a data source column name to a MERGEFIELD name like this.
        mappedDataFields.add("MergeFieldName", "DataSourceColumnName");

        // Link the data source column named "Column2" to MERGEFIELDs named "Column3".
        mappedDataFields.add("Column3", "Column2");

        // The MERGEFIELD name is the "key" to the respective data source column name "value".
        Assert.assertEquals("DataSourceColumnName", mappedDataFields.get("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsValue("DataSourceColumnName"));

        // Now if we run this mail merge, the "Column3" MERGEFIELDs will take data from "Column2" of the table.
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "MailMerge.MappedDataFieldCollection.docx");

        // We can iterate over the elements in this collection.
        Assert.assertEquals(2, mappedDataFields.getCount());

        Iterator<Map.Entry<String, String>> enumerator = mappedDataFields.iterator();
        try /*JAVA: was using*/
    	{
            while (enumerator.hasNext())
                System.out.println("Column named {enumerator.Current.Value} is mapped to MERGEFIELDs named {enumerator.Current.Key}");
    	}
        finally { if (enumerator != null) enumerator.close(); }

        // We can also remove elements from the collection.
        mappedDataFields.remove("MergeFieldName");

        Assert.assertFalse(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertFalse(mappedDataFields.containsValue("DataSourceColumnName"));

        mappedDataFields.clear();

        Assert.assertEquals(0, mappedDataFields.getCount());
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMerge.MappedDataFieldCollection.docx"), true); //ExSkip
    }

    /// <summary>
    /// Create a document with 2 MERGEFIELDs, one of which does not have a
    /// corresponding column in the data table from the method below.
    /// </summary>
    private static Document createSourceDocMappedDataFields() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column3");

        return doc;
    }

    /// <summary>
    /// Create a data table with 2 columns, one of which does not have a
    /// corresponding MERGEFIELD in the source document from the method above.
    /// </summary>
    private static DataTable createSourceTableMappedDataFields()
    {
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
        //ExSummary:Shows how to get mail merge field names used by a field.
        Document doc = new Document(getMyDir() + "Field sample - ADDRESSBLOCK.docx");

        String[] addressFieldsExpect =
        {
            "Company", "First Name", "Middle Name", "Last Name", "Suffix", "Address 1", "City", "State",
            "Country or Region", "Postal Code"
        };

        FieldAddressBlock addressBlockField = (FieldAddressBlock) doc.getRange().getFields().get(0);
        String[] addressBlockFieldNames = addressBlockField.getFieldNames();
        //ExEnd

        Assert.assertEquals(addressFieldsExpect, addressBlockFieldNames);

        String[] greetingFieldsExpect = { "Courtesy Title", "Last Name" };

        FieldGreetingLine greetingLineField = (FieldGreetingLine) doc.getRange().getFields().get(1);
        String[] greetingLineFieldNames = greetingLineField.getFieldNames();

        Assert.assertEquals(greetingFieldsExpect, greetingLineFieldNames);
    }

    /// <summary>
    /// Without TestCaseSource/TestCase because of some strange behavior when using long data.
    /// </summary>
    @Test
    public void mustacheTemplateSyntaxTrue() throws Exception
    {
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
    public void mustacheTemplateSyntaxFalse() throws Exception
    {
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
        //ExSummary:Shows how to verify mail merge regions.
        Document doc = new Document(getMyDir() + "Mail merge regions.docx");

        // Returns a full hierarchy of merge regions that contain MERGEFIELDs available in the document.
        MailMergeRegionInfo regionInfo = doc.getMailMerge().getRegionsHierarchy();

        // Get top regions in the document.
        ArrayList<MailMergeRegionInfo> topRegions = regionInfo.getRegions();

        Assert.assertEquals(2, topRegions.size());
        Assert.assertEquals("Region1", topRegions.get(0).getName());
        Assert.assertEquals("Region2", topRegions.get(1).getName());
        Assert.assertEquals(1, topRegions.get(0).getLevel());
        Assert.assertEquals(1, topRegions.get(1).getLevel());

        // Get nested region in first top region.
        ArrayList<MailMergeRegionInfo> nestedRegions = topRegions.get(0).getRegions();

        Assert.assertEquals(2, nestedRegions.size());
        Assert.assertEquals("NestedRegion1", nestedRegions.get(0).getName());
        Assert.assertEquals("NestedRegion2", nestedRegions.get(1).getName());
        Assert.assertEquals(2, nestedRegions.get(0).getLevel());
        Assert.assertEquals(2, nestedRegions.get(1).getLevel());

        // Get list of fields inside the first top region.
        ArrayList<Field> fieldList = topRegions.get(0).getFields();

        Assert.assertEquals(4, fieldList.size());

        FieldMergeField startFieldMergeField = nestedRegions.get(0).getStartField();

        Assert.assertEquals("TableStart:NestedRegion1", startFieldMergeField.getFieldName());

        FieldMergeField endFieldMergeField = nestedRegions.get(0).getEndField();

        Assert.assertEquals("TableEnd:NestedRegion1", endFieldMergeField.getFieldName());
        //ExEnd
    }

    //ExStart
    //ExFor:MailMerge.MailMergeCallback
    //ExFor:IMailMergeCallback
    //ExFor:IMailMergeCallback.TagsReplaced
    //ExSummary:Shows how to define custom logic for handling events during mail merge.
    @Test //ExSkip
    public void callback() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two mail merge tags referencing two columns in a data source.
        builder.write("{{FirstName}}");
        builder.write("{{LastName}}");

        // Create a data source that only contains one of the columns that our merge tags reference.
        DataTable table = new DataTable("Test");
        table.getColumns().add("FirstName");
        table.getRows().add("John");
        table.getRows().add("Jane");

        // Configure our mail merge to use alternative mail merge tags.
        doc.getMailMerge().setUseNonMergeFields(true);

        // Then, ensure that the mail merge will convert tags, such as our "LastName" tag,
        // into MERGEFIELDs in the merge documents.
        doc.getMailMerge().setPreserveUnusedTags(false);

        MailMergeTagReplacementCounter counter = new MailMergeTagReplacementCounter();
        doc.getMailMerge().setMailMergeCallback(counter);
        doc.getMailMerge().execute(table);

        Assert.assertEquals(1, counter.getTagsReplacedCount());
    }

    /// <summary>
    /// Counts the number of times a mail merge replaces mail merge tags that it could not fill with data with MERGEFIELDs.
    /// </summary>
    private static class MailMergeTagReplacementCounter implements IMailMergeCallback
    {
        public void tagsReplaced()
        {
            setTagsReplacedCount(getTagsReplacedCount() + 1)/*Property++*/;
        }

        public int getTagsReplacedCount() { return mTagsReplacedCount; }; private void setTagsReplacedCount(int value) { mTagsReplacedCount = value; };

        private int mTagsReplacedCount;
    }
    //ExEnd

    @Test
    public void getRegionsByName() throws Exception
    {
        Document doc = new Document(getMyDir() + "Mail merge regions.docx");

        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("Region1");
        Assert.assertEquals(1, doc.getMailMerge().getRegionsByName("Region1").size());
        for (MailMergeRegionInfo region : regions) Assert.assertEquals("Region1", region.getName());

        regions = doc.getMailMerge().getRegionsByName("Region2");
        Assert.assertEquals(1, doc.getMailMerge().getRegionsByName("Region2").size());
        for (MailMergeRegionInfo region : regions) Assert.assertEquals("Region2", region.getName());

        regions = doc.getMailMerge().getRegionsByName("NestedRegion1");
        Assert.assertEquals(2, doc.getMailMerge().getRegionsByName("NestedRegion1").size());
        for (MailMergeRegionInfo region : regions) Assert.assertEquals("NestedRegion1", region.getName());
    }

    @Test
    public void cleanupOptions() throws Exception
    {
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

    /// <summary>
    /// Return a data table filled with sample data.
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

    @Test (dataProvider = "unconditionalMergeFieldsAndRegionsDataProvider")
    public void unconditionalMergeFieldsAndRegions(boolean countAllMergeFields) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.UnconditionalMergeFieldsAndRegions
        //ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD nested inside an IF field.
        // Since the IF field statement is false, it will not display the result of the MERGEFIELD.
        // The MERGEFIELD will also not receive any data during a mail merge.
        FieldIf fieldIf = (FieldIf)builder.insertField(" IF 1 = 2 ");
        builder.moveTo(fieldIf.getSeparator());
        builder.insertField(" MERGEFIELD  FullName ");

        // If we set the "UnconditionalMergeFieldsAndRegions" flag to "true",
        // our mail merge will insert data into non-displayed fields such as our MERGEFIELD as well as all others.
        // If we set the "UnconditionalMergeFieldsAndRegions" flag to "false",
        // our mail merge will not insert data into MERGEFIELDs hidden by IF fields with false statements.
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(countAllMergeFields);

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FullName");
        dataTable.getRows().add("James Bond");

        doc.getMailMerge().execute(dataTable);
        
        doc.save(getArtifactsDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");

        Assert.assertEquals(
            countAllMergeFields
                ? "\u0013 IF 1 = 2 \"James Bond\"\u0014\u0015"
                : "\u0013 IF 1 = 2 \u0013 MERGEFIELD  FullName \u0014«FullName»\u0015\u0014\u0015",
            doc.getText().trim());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "unconditionalMergeFieldsAndRegionsDataProvider")
	public static Object[][] unconditionalMergeFieldsAndRegionsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "retainFirstSectionStartDataProvider")
    public void retainFirstSectionStart(boolean isRetainFirstSectionStart, /*SectionStart*/int sectionStart, /*SectionStart*/int expected) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.insertField(" MERGEFIELD  FullName ");

        doc.getFirstSection().getPageSetup().setSectionStart(sectionStart);
        doc.getMailMerge().setRetainFirstSectionStart(isRetainFirstSectionStart);

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FullName");
        dataTable.getRows().add("James Bond");

        doc.getMailMerge().execute(dataTable);

        for (Section section : (Iterable<Section>) doc.getSections())
            Assert.assertEquals(expected, section.getPageSetup().getSectionStart());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "retainFirstSectionStartDataProvider")
	public static Object[][] retainFirstSectionStartDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true,  SectionStart.CONTINUOUS,  SectionStart.CONTINUOUS},
			{true,  SectionStart.NEW_COLUMN,  SectionStart.NEW_COLUMN},
			{true,  SectionStart.NEW_PAGE,  SectionStart.NEW_PAGE},
			{true,  SectionStart.EVEN_PAGE,  SectionStart.EVEN_PAGE},
			{true,  SectionStart.ODD_PAGE,  SectionStart.ODD_PAGE},
			{false,  SectionStart.CONTINUOUS,  SectionStart.NEW_PAGE},
			{false,  SectionStart.NEW_COLUMN,  SectionStart.NEW_PAGE},
			{false,  SectionStart.NEW_PAGE,  SectionStart.NEW_PAGE},
			{false,  SectionStart.EVEN_PAGE,  SectionStart.EVEN_PAGE},
			{false,  SectionStart.ODD_PAGE,  SectionStart.ODD_PAGE},
		};
	}

    @Test
    public void mailMergeSettings() throws Exception
    {
        //ExStart
        //ExFor:Document.MailMergeSettings
        //ExFor:MailMergeCheckErrors
        //ExFor:MailMergeDataType
        //ExFor:MailMergeDestination
        //ExFor:MailMergeMainDocumentType
        //ExFor:MailMergeSettings
        //ExFor:MailMergeSettings.CheckErrors
        //ExFor:MailMergeSettings.Clone
        //ExFor:MailMergeSettings.Destination
        //ExFor:MailMergeSettings.DataType
        //ExFor:MailMergeSettings.DoNotSupressBlankLines
        //ExFor:MailMergeSettings.LinkToQuery
        //ExFor:MailMergeSettings.MainDocumentType
        //ExFor:MailMergeSettings.Odso
        //ExFor:MailMergeSettings.Query
        //ExFor:MailMergeSettings.ViewMergedData
        //ExFor:Odso
        //ExFor:Odso.Clone
        //ExFor:Odso.ColumnDelimiter
        //ExFor:Odso.DataSource
        //ExFor:Odso.DataSourceType
        //ExFor:Odso.FirstRowContainsColumnNames
        //ExFor:OdsoDataSourceType
        //ExSummary:Shows how to execute a mail merge with data from an Office Data Source Object.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField("MERGEFIELD FirstName", "<FirstName>");
        builder.write(" ");
        builder.insertField("MERGEFIELD LastName", "<LastName>");
        builder.writeln(": ");
        builder.insertField("MERGEFIELD Message", "<Message>");

        // Create a data source in the form of an ASCII file, with the "|" character
        // acting as the delimiter that separates columns. The first line contains the three columns' names,
        // and each subsequent line is a row with their respective values.
        String[] lines = { "FirstName|LastName|Message",
            "John|Doe|Hello! This message was created with Aspose Words mail merge." };
        String dataSrcFilename = getArtifactsDir() + "MailMerge.MailMergeSettings.DataSource.txt";

        File.writeAllLines(dataSrcFilename, lines);

        MailMergeSettings settings = doc.getMailMergeSettings();
        settings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        settings.setCheckErrors(MailMergeCheckErrors.SIMULATE);
        settings.setDataType(MailMergeDataType.NATIVE);
        settings.setDataSource(dataSrcFilename);
        settings.setQuery("SELECT * FROM " + doc.getMailMergeSettings().getDataSource());
        settings.setLinkToQuery(true);
        settings.setViewMergedData(true);

        Assert.assertEquals(MailMergeDestination.DEFAULT, settings.getDestination());
        Assert.assertFalse(settings.getDoNotSupressBlankLines());

        Odso odso = settings.getOdso();
        odso.setDataSource(dataSrcFilename);
        odso.setDataSourceType(OdsoDataSourceType.TEXT);
        odso.setColumnDelimiter('|');
        odso.setFirstRowContainsColumnNames(true);

        Assert.assertNotSame(odso, odso.deepClone());
        Assert.assertNotSame(settings, settings.deepClone());

        // Opening this document in Microsoft Word will execute the mail merge before displaying the contents. 
        doc.save(getArtifactsDir() + "MailMerge.MailMergeSettings.docx");
        //ExEnd

        settings = new Document(getArtifactsDir() + "MailMerge.MailMergeSettings.docx").getMailMergeSettings();

        Assert.assertEquals(MailMergeMainDocumentType.MAILING_LABELS, settings.getMainDocumentType());
        Assert.assertEquals(MailMergeCheckErrors.SIMULATE, settings.getCheckErrors());
        Assert.assertEquals(MailMergeDataType.NATIVE, settings.getDataType());
        Assert.assertEquals(getArtifactsDir() + "MailMerge.MailMergeSettings.DataSource.txt", settings.getDataSource());
        Assert.assertEquals("SELECT * FROM " + doc.getMailMergeSettings().getDataSource(), settings.getQuery());
        Assert.assertTrue(settings.getLinkToQuery());
        Assert.assertTrue(settings.getViewMergedData());

        odso = settings.getOdso();
        Assert.assertEquals(getArtifactsDir() + "MailMerge.MailMergeSettings.DataSource.txt", odso.getDataSource());
        Assert.assertEquals(OdsoDataSourceType.TEXT, odso.getDataSourceType());
        Assert.assertEquals('|', odso.getColumnDelimiter());
        Assert.assertTrue(odso.getFirstRowContainsColumnNames());
    }

    @Test
    public void odsoEmail() throws Exception
    {
        //ExStart
        //ExFor:MailMergeSettings.ActiveRecord
        //ExFor:MailMergeSettings.AddressFieldName
        //ExFor:MailMergeSettings.ConnectString
        //ExFor:MailMergeSettings.MailAsAttachment
        //ExFor:MailMergeSettings.MailSubject
        //ExFor:MailMergeSettings.Clear
        //ExFor:Odso.TableName
        //ExFor:Odso.UdlConnectString
        //ExSummary:Shows how to execute a mail merge while connecting to an external data source.
        Document doc = new Document(getMyDir() + "Odso data.docx");
        testOdsoEmail(doc); //ExSkip
        MailMergeSettings settings = doc.getMailMergeSettings();

        System.out.println("Connection string:\n\t{settings.ConnectString}");
        System.out.println("Mail merge docs as attachment:\n\t{settings.MailAsAttachment}");
        System.out.println("Mail merge doc e-mail subject:\n\t{settings.MailSubject}");
        System.out.println("Column that contains e-mail addresses:\n\t{settings.AddressFieldName}");
        System.out.println("Active record:\n\t{settings.ActiveRecord}");

        Odso odso = settings.getOdso();

        System.out.println("File will connect to data source located in:\n\t\"{odso.DataSource}\"");
        System.out.println("Source type:\n\t{odso.DataSourceType}");
        System.out.println("UDL connection string:\n\t{odso.UdlConnectString}");
        System.out.println("Table:\n\t{odso.TableName}");
        System.out.println("Query:\n\t{doc.MailMergeSettings.Query}");

        // We can reset these settings by clearing them. Once we do that and save the document,
        // Microsoft Word will no longer execute a mail merge when we use it to load the document.
        settings.clear();

        doc.save(getArtifactsDir() + "MailMerge.OdsoEmail.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "MailMerge.OdsoEmail.docx");
        Assert.That(doc.getMailMergeSettings().getConnectString(), Is.Empty);
    }

    private void testOdsoEmail(Document doc)
    {
        MailMergeSettings settings = doc.getMailMergeSettings();

        Assert.assertFalse(settings.getMailAsAttachment());
        Assert.assertEquals("test subject", settings.getMailSubject());
        Assert.assertEquals("Email_Address", settings.getAddressFieldName());
        Assert.assertEquals(66, settings.getActiveRecord());
        Assert.assertEquals("SELECT * FROM `Contacts` ", settings.getQuery());

        Odso odso = settings.getOdso();

        Assert.assertEquals(settings.getConnectString(), odso.getUdlConnectString());
        Assert.assertEquals("Personal Folders|", odso.getDataSource());
        Assert.assertEquals(OdsoDataSourceType.EMAIL, odso.getDataSourceType());
        Assert.assertEquals("Contacts", odso.getTableName());
    }

    @Test
    public void mailingLabelMerge() throws Exception
    {
        //ExStart
        //ExFor:MailMergeSettings.DataSource
        //ExFor:MailMergeSettings.HeaderSource
        //ExSummary:Shows how to construct a data source for a mail merge from a header source and a data source.
        // Create a mailing label merge header file, which will consist of a table with one row.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.write("FirstName");
        builder.insertCell();
        builder.write("LastName");
        builder.endTable();

        doc.save(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx");

        // Create a mailing label merge data file consisting of a table with one row
        // and the same number of columns as the header document's table. 
        doc = new Document();
        builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.write("John");
        builder.insertCell();
        builder.write("Doe");
        builder.endTable();

        doc.save(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx");

        // Create a merge destination document with MERGEFIELDS with names that
        // match the column names in the merge header file table.
        doc = new Document();
        builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField("MERGEFIELD FirstName", "<FirstName>");
        builder.write(" ");
        builder.insertField("MERGEFIELD LastName", "<LastName>");

        MailMergeSettings settings = doc.getMailMergeSettings();

        // Construct a data source for our mail merge by specifying two document filenames.
        // The header source will name the columns of the data source table.
        settings.setHeaderSource(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx");

        // The data source will provide rows of data for all the columns in the header document table.
        settings.setDataSource(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx");

        // Configure a mailing label type mail merge, which Microsoft Word will execute
        // as soon as we use it to load the output document.
        settings.setQuery("SELECT * FROM " + settings.getDataSource());
        settings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        settings.setDataType(MailMergeDataType.TEXT_FILE);
        settings.setLinkToQuery(true);
        settings.setViewMergedData(true);

        doc.save(getArtifactsDir() + "MailMerge.MailingLabelMerge.docx");
        //ExEnd

        Assert.assertEquals("FirstName\u0007LastName\u0007\u0007",
            new Document(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx").
                getChild(NodeType.TABLE, 0, true).getText().trim());

        Assert.assertEquals("John\u0007Doe\u0007\u0007",
            new Document(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx").
                getChild(NodeType.TABLE, 0, true).getText().trim());

        doc = new Document(getArtifactsDir() + "MailMerge.MailingLabelMerge.docx");

        Assert.assertEquals(2, doc.getRange().getFields().getCount());

        settings = doc.getMailMergeSettings();

        Assert.assertEquals(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx", settings.getHeaderSource());
        Assert.assertEquals(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx", settings.getDataSource());
        Assert.assertEquals("SELECT * FROM " + settings.getDataSource(), settings.getQuery());
        Assert.assertEquals(MailMergeMainDocumentType.MAILING_LABELS, settings.getMainDocumentType());
        Assert.assertEquals(MailMergeDataType.TEXT_FILE, settings.getDataType());
        Assert.assertTrue(settings.getLinkToQuery());
        Assert.assertTrue(settings.getViewMergedData());
    }

    @Test
    public void odsoFieldMapDataCollection() throws Exception
    {
        //ExStart
        //ExFor:Odso.FieldMapDatas
        //ExFor:OdsoFieldMapData
        //ExFor:OdsoFieldMapData.Clone
        //ExFor:OdsoFieldMapData.Column
        //ExFor:OdsoFieldMapData.MappedName
        //ExFor:OdsoFieldMapData.Name
        //ExFor:OdsoFieldMapData.Type
        //ExFor:OdsoFieldMapDataCollection
        //ExFor:OdsoFieldMapDataCollection.Add(OdsoFieldMapData)
        //ExFor:OdsoFieldMapDataCollection.Clear
        //ExFor:OdsoFieldMapDataCollection.Count
        //ExFor:OdsoFieldMapDataCollection.GetEnumerator
        //ExFor:OdsoFieldMapDataCollection.Item(Int32)
        //ExFor:OdsoFieldMapDataCollection.RemoveAt(Int32)
        //ExFor:OdsoFieldMappingType
        //ExSummary:Shows how to access the collection of data that maps data source columns to merge fields.
        Document doc = new Document(getMyDir() + "Odso data.docx");

        // This collection defines how a mail merge will map columns from a data source
        // to predefined MERGEFIELD, ADDRESSBLOCK and GREETINGLINE fields.
        OdsoFieldMapDataCollection dataCollection = doc.getMailMergeSettings().getOdso().getFieldMapDatas();
        Assert.assertEquals(30, dataCollection.getCount());

        Iterator<OdsoFieldMapData> enumerator = dataCollection.iterator();
        try /*JAVA: was using*/
        {
            int index = 0;
            while (enumerator.hasNext())
            {
                System.out.println("Field map data index {index++}, type \"{enumerator.Current.Type}\":");

                System.out.println(enumerator.next().getType() != OdsoFieldMappingType.NULL
                            ? $"\tColumn \"{enumerator.Current.Name}\", number {enumerator.Current.Column} mapped to merge field \"{enumerator.Current.MappedName}\"."
                            : "\tNo valid column to field mapping data present.");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // Clone the elements in this collection.
        Assert.assertNotEquals(dataCollection.get(0), dataCollection.get(0).deepClone());

        // Use the "RemoveAt" method elements individually by index.
        dataCollection.removeAt(0);

        Assert.assertEquals(29, dataCollection.getCount());

        // Use the "Clear" method to clear the entire collection at once.
        dataCollection.clear();

        Assert.assertEquals(0, dataCollection.getCount());
        //ExEnd
    }

    @Test
    public void odsoRecipientDataCollection() throws Exception
    {
        //ExStart
        //ExFor:Odso.RecipientDatas
        //ExFor:OdsoRecipientData
        //ExFor:OdsoRecipientData.Active
        //ExFor:OdsoRecipientData.Clone
        //ExFor:OdsoRecipientData.Column
        //ExFor:OdsoRecipientData.Hash
        //ExFor:OdsoRecipientData.UniqueTag
        //ExFor:OdsoRecipientDataCollection
        //ExFor:OdsoRecipientDataCollection.Add(OdsoRecipientData)
        //ExFor:OdsoRecipientDataCollection.Clear
        //ExFor:OdsoRecipientDataCollection.Count
        //ExFor:OdsoRecipientDataCollection.GetEnumerator
        //ExFor:OdsoRecipientDataCollection.Item(Int32)
        //ExFor:OdsoRecipientDataCollection.RemoveAt(Int32)
        //ExSummary:Shows how to access the collection of data that designates which merge data source records a mail merge will exclude.
        Document doc = new Document(getMyDir() + "Odso data.docx");

        OdsoRecipientDataCollection dataCollection = doc.getMailMergeSettings().getOdso().getRecipientDatas();

        Assert.assertEquals(70, dataCollection.getCount());

        Iterator<OdsoRecipientData> enumerator = dataCollection.iterator();
        try /*JAVA: was using*/
        {
            int index = 0;
            while (enumerator.hasNext())
            {
                System.out.println("Odso recipient data index {index++} will {(enumerator.Current.Active ? ");
                System.out.println("\tColumn #{enumerator.Current.Column}");
                System.out.println("\tHash code: {enumerator.Current.Hash}");
                System.out.println("\tContents array length: {enumerator.Current.UniqueTag.Length}");
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        // We can clone the elements in this collection.
        Assert.assertNotEquals(dataCollection.get(0), dataCollection.get(0).deepClone());

        // We can also remove elements individually, or clear the entire collection at once.
        dataCollection.removeAt(0);

        Assert.assertEquals(69, dataCollection.getCount());

        dataCollection.clear();

        Assert.assertEquals(0, dataCollection.getCount());
        //ExEnd
    }

    @Test
    public void changeFieldUpdateCultureSource() throws Exception
    {
        //ExStart
        //ExFor:Document.FieldOptions
        //ExFor:FieldOptions
        //ExFor:FieldOptions.FieldUpdateCultureSource
        //ExFor:FieldUpdateCultureSource
        //ExSummary:Shows how to specify the source of the culture used for date formatting during a field update or mail merge.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two merge fields with German locale.
        builder.getFont().setLocaleId(new msCultureInfo("de-DE").getLCID());
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

        // Set the current culture to US English after preserving its original value in a variable.
        msCultureInfo currentCulture = CurrentThread.getCurrentCulture();
        CurrentThread.setCurrentCulture(new msCultureInfo("en-US"));

        // This merge will use the current thread's culture to format the date, US English.
        doc.getMailMerge().execute(new String[] { "Date1" }, new Object[] { new DateTime(2020, 1, 1) });

        // Configure the next merge to source its culture value from the field code. The value of that culture will be German.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { new DateTime(2020, 1, 1) });

        // The first merge result contains a date formatted in English, while the second one is in German.
        Assert.assertEquals("Wednesday, 1 January 2020 - Mittwoch, 1 Januar 2020", doc.getRange().getText().trim());

        // Restore the thread's original culture.
        CurrentThread.setCurrentCulture(currentCulture);
        //ExEnd
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
