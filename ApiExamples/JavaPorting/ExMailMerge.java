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
import com.aspose.ms.System.msString;
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
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.OdsoRecipientDataCollection;
import com.aspose.words.OdsoRecipientData;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.ms.System.Threading.CurrentThread;
import com.aspose.ms.System.DateTime;
import com.aspose.words.FieldUpdateCultureSource;
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
        //ExSummary:Performs a simple insertion of data into merge fields and sends the document to the browser inline.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" MERGEFIELD FullName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Company ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD City ");

        // Fill the fields in the document with user data
        doc.getMailMerge().execute(new String[] { "FullName", "Company", "Address", "City" },
            new Object[] { "James Bond", "MI5 Headquarters", "Milbank", "London" });

        // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser
        Assert.That(() => doc.Save(response, "Artifacts/MailMerge.ExecuteArray.docx", ContentDisposition.INLINE, null),
            Throws.<NullPointerException>TypeOf()); //Thrown because HttpResponse is null in the test.

        // The response will need to be closed manually to make sure that no superfluous content is added to the document after saving
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
        // Create a new document and populate it with merge fields
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

        // Create a connection string which points to the "Northwind" database file in our local file system and open a connection, and set up a query
        String connectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" + getDatabaseDir() + "Northwind.mdb";
        String query = "SELECT Products.ProductName, Suppliers.CompanyName, Products.QuantityPerUnit, {fn ROUND(Products.UnitPrice,2)} as UnitPrice\r\n                                        FROM Products \r\n                                        INNER JOIN Suppliers \r\n                                        ON Products.SupplierID = Suppliers.SupplierID";

        OdbcConnection connection = new OdbcConnection();
        try /*JAVA: was using*/
        {
            connection.ConnectionString = connectionString;
            connection.Open();

            // Create an SQL command that will source data for our mail merge
            // The names of the columns returned by this SELECT statement should correspond to the merge fields we placed above
            OdbcCommand command = connection.CreateCommand();
            command.CommandText = query;

            // This will run the command and store the data in the reader
            OdbcDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

            // Now we can take the data from the reader and use it in the mail merge
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
        // Create a document that will be merged
        Document doc = createSourceDocADOMailMerge();

        // To work with ADO DataSets, we need to add a reference to the Microsoft ActiveX Data Objects library,
        // which is included in the .NET distribution and stored in "adodb.dll", then create a connection
        ADODB.Connection connection = new ADODB.Connection();

        // Create a connection string which points to the "Northwind" database file in our local file system and open a connection
        String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";
        connection.Open(connectionString);

        // Create a record set
        ADODB.Recordset recordset = new ADODB.Recordset();

        // Populate our DataSrt by running an SQL command on the database we are connected to
        // The names of the columns returned here correspond to the values of the MERGEFIELDS that will accommodate our data
        String command = "SELECT ProductName, QuantityPerUnit, UnitPrice FROM Products";
        recordset.Open(command, connection);

        // Execute the mail merge and save the document
        doc.getMailMerge().ExecuteADO(recordset);
        doc.save(getArtifactsDir() + "MailMerge.ExecuteADO.docx");
        TestUtil.mailMergeMatchesQueryResult(getDatabaseDir() + "Northwind.mdb", command, doc, true); //ExSkip
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
    //ExSummary:Shows how to run a mail merge with regions, compiled with data from an ADO dataset.
    @Test (groups = "SkipMono") //ExSkip
    public void executeWithRegionsADO() throws Exception
    {
        // Create a document that will be merged
        Document doc = createSourceDocADOMailMergeWithRegions();

        // To work with ADO DataSets, we need to add a reference to the Microsoft ActiveX Data Objects library,
        // which is included in the .NET distribution and stored in "adodb.dll", then create a connection
        ADODB.Connection connection = new ADODB.Connection();

        // Create a connection string which points to the "Northwind" database file in our local file system and open a connection
        String connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";
        connection.Open(connectionString);

        // Create a record set
        ADODB.Recordset recordset = new ADODB.Recordset();

        // Create an SQL query that fetches data with column names that are suitable for our first mail merge region,
        // then populate our record set with the data
        String command = "SELECT FirstName, LastName, City FROM Employees";
        recordset.Open(command, connection);

        // Run a mail merge on just the first region, filling its MERGEFIELDS with data from the ADO record set
        doc.getMailMerge().ExecuteWithRegionsADO(recordset, "MergeRegion1");

        // Close the record set and reopen it with data from another SQL query
        recordset.Close();
        command = "SELECT * FROM Customers";
        recordset.Open(command, connection);

        // Run a mail merge on the second region and save the document
        doc.getMailMerge().ExecuteWithRegionsADO(recordset, "MergeRegion2");

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegionsADO.docx");
        TestUtil.mailMergeMatchesQueryResultMultiple(getDatabaseDir() + "Northwind.mdb", new String[] { "SELECT FirstName, LastName, City FROM Employees", "SELECT ContactName, Address, City FROM Customers" }, new Document(getArtifactsDir() + "MailMerge.ExecuteWithRegionsADO.docx"), false); //ExSkip
    }

    /// <summary>
    /// Create a blank document and use MERGEFIELDS to create two sequential mail merge regions with TableStart/TableEnd tags
    /// </summary>
    private static Document createSourceDocADOMailMergeWithRegions() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First mail merge region
        builder.writeln("\tEmployees: ");
        builder.insertField(" MERGEFIELD TableStart:MergeRegion1");
        builder.insertField(" MERGEFIELD FirstName");
        builder.write(", ");
        builder.insertField(" MERGEFIELD LastName");
        builder.write(", ");
        builder.insertField(" MERGEFIELD City");
        builder.insertField(" MERGEFIELD TableEnd:MergeRegion1");
        builder.insertParagraph();

        // Second mail merge region
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
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField(" MERGEFIELD CustomerName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");

        // This example creates a table, but you would normally load table from a database
        DataTable table = new DataTable("Test");
        table.getColumns().add("CustomerName");
        table.getColumns().add("Address");
        table.getRows().add(new Object[] { "Thomas Hardy", "120 Hanover Sq., London" });
        table.getRows().add(new Object[] { "Paolo Accorti", "Via Monte Bianco 34, Torino" });

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

        TestUtil.mailMergeMatchesDataTable(table, new Document(getArtifactsDir() + "MailMerge.ExecuteDataTable.docx"), true);

        DataTable rowAsTable = new DataTable();
        rowAsTable.importRow(table.getRows().get(1));

        TestUtil.mailMergeMatchesDataTable(rowAsTable, new Document(getArtifactsDir() + "MailMerge.ExecuteDataTable.OneRow.docx"), true);
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

        TestUtil.mailMergeMatchesDataTable(view.toTable(), new Document(getArtifactsDir() + "MailMerge.ExecuteDataView.docx"), true);
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
        TestUtil.mailMergeMatchesDataSet(customersAndOrders, new Document(getArtifactsDir() + "MailMerge.ExecuteWithRegionsNested.docx"), false); //ExSkip
    }

    /// <summary>
    /// Generates a data set which has two data tables named "Customers" and "Orders", with a one-to-many relationship on the "CustomerID" column.
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

        // Create two data tables that are not linked or related in any way which we still want in the same document
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
        //ExSummary:Shows how to create, list and read mail merge regions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // These tags, which go inside MERGEFIELDs, denote the strings that signify the starts and ends of mail merge regions 
        Assert.assertEquals("TableStart", doc.getMailMerge().getRegionStartTag());
        Assert.assertEquals("TableEnd", doc.getMailMerge().getRegionEndTag());

        // By using these tags, we will start and end a "MailMergeRegion1", which will contain MERGEFIELDs for two columns
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.insertField(" MERGEFIELD Column1");
        builder.write(", ");
        builder.insertField(" MERGEFIELD Column2");
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // We can keep track of merge regions and their columns by looking at these collections
        ArrayList<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");
        Assert.assertEquals(1, regions.size());
        Assert.assertEquals("MailMergeRegion1", regions.get(0).getName());

        String[] mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1");
        Assert.assertEquals("Column1", mergeFieldNames[0]);
        Assert.assertEquals("Column2", mergeFieldNames[1]);

        // Insert a region with the same name as an existing region, which will make it a duplicate
        builder.insertParagraph(); // A single row/paragraph cannot be shared by multiple regions
        builder.insertField(" MERGEFIELD TableStart:MailMergeRegion1");
        builder.insertField(" MERGEFIELD Column3");
        builder.insertField(" MERGEFIELD TableEnd:MailMergeRegion1");

        // Regions that share the same name are still accounted for and can be accessed by index
        regions = doc.getMailMerge().getRegionsByName("MailMergeRegion1");
        Assert.assertEquals(2, regions.size());

        mergeFieldNames = doc.getMailMerge().getFieldNamesForRegion("MailMergeRegion1", 1);
        Assert.assertEquals("Column3", mergeFieldNames[0]);
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
        testMergeDuplicateRegions(dataTable, doc, isMergeDuplicateRegions); //ExSkip
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
    //ExSummary:Shows how to preserve the appearance of alternative mail merge tags that go unused during a mail merge. 
    @Test (dataProvider = "preserveUnusedTagsDataProvider") //ExSkip
    public void preserveUnusedTags(boolean doPreserveUnusedTags) throws Exception
    {
        // Create a document and table that we will merge
        Document doc = createSourceDocWithAlternativeMergeFields();
        DataTable dataTable = createSourceTablePreserveUnusedTags();

        // By default, alternative merge tags that cannot receive data because the data source has no columns with their name
        // are converted to and left on display as MERGEFIELDs after the mail merge
        // We can preserve their original appearance setting this attribute to true
        doc.getMailMerge().setPreserveUnusedTags(doPreserveUnusedTags);
        doc.getMailMerge().execute(dataTable);

        doc.save(getArtifactsDir() + "MailMerge.PreserveUnusedTags.docx");

        Assert.assertEquals(doc.getText().contains("{{ Column2 }}"), doPreserveUnusedTags);
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
    public void mergeWholeDocument(boolean doMergeWholeDocument) throws Exception
    {
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
        Assert.<IllegalStateException>Throws(() => doc.getMailMerge().executeWithRegions(dataTable));

        // If we set this variable to false, paragraphs and mail merge regions are independent so we can safely run our mail merge
        doc.getMailMerge().setUseWholeParagraphAsRegion(false);
        doc.getMailMerge().executeWithRegions(dataTable);

        // Our first region is populated, while our second is safely displayed as unused all across one paragraph
        doc.save(getArtifactsDir() + "MailMerge.UseWholeParagraphAsRegion.docx");
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMerge.UseWholeParagraphAsRegion.docx"), true); //ExSkip
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
    public void trimWhiteSpaces(boolean doTrimWhitespaces) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.TrimWhitespaces
        //ExSummary:Shows how to trimmed whitespaces from mail merge values.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField("MERGEFIELD myMergeField", null);

        doc.getMailMerge().setTrimWhitespaces(doTrimWhitespaces);
        doc.getMailMerge().execute(new String[] { "myMergeField" }, new Object[] { "\t hello world! " });

        if (doTrimWhitespaces)
            Assert.assertEquals("hello world!\f", doc.getText());
        else
            Assert.assertEquals("\t hello world! \f", doc.getText());
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

        // Here is the complete list of cleanable punctuation marks: ! , . : ; ? ¡ ¿
        builder.write(punctuationMark);

        FieldMergeField mergeFieldOption2 = (FieldMergeField) builder.insertField("MERGEFIELD", "Option_2");
        mergeFieldOption2.setFieldName("Option_2");

        doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_PARAGRAPHS);
        // The default value of the option is true which means that the behavior was changed to mimic MS Word
        // We can revert to the old behavior by setting the option to false
        doc.getMailMerge().setCleanupParagraphsWithPunctuationMarks(isCleanupParagraphsWithPunctuationMarks);

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
        // Create a document and table that we will merge
        Document doc = createSourceDocMappedDataFields();
        DataTable dataTable = createSourceTableMappedDataFields();

        // We have a column "Column2" in the data table that does not have a respective MERGEFIELD in the document
        // Also, we have a MERGEFIELD named "Column3" that does not exist as a column in the data source
        // If data from "Column2" is suitable for the "Column3" MERGEFIELD,
        // we can map that column name to the MERGEFIELD in the "MappedDataFields" key/value pair
        MappedDataFieldCollection mappedDataFields = doc.getMailMerge().getMappedDataFields();

        // A data source column name is linked to a MERGEFIELD name by adding an element like this
        mappedDataFields.add("MergeFieldName", "DataSourceColumnName");

        // So, values from "Column2" will now go into MERGEFIELDs named "Column3" as well as "Column2", if there are any
        mappedDataFields.add("Column3", "Column2");

        // The MERGEFIELD name is the "key" to the respective data source column name "value"
        Assert.assertEquals("DataSourceColumnName", mappedDataFields.get("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertTrue(mappedDataFields.containsValue("DataSourceColumnName"));

        // Now if we run this mail merge, the "Column3" MERGEFIELDs will take data from "Column2" of the table
        doc.getMailMerge().execute(dataTable);

        // We can count and iterate over the mapped columns/fields
        Assert.assertEquals(2, mappedDataFields.getCount());

        Iterator<Map.Entry<String, String>> enumerator = mappedDataFields.iterator();
        try /*JAVA: was using*/
    	{
            while (enumerator.hasNext())
                System.out.println("Column named {enumerator.Current.Value} is mapped to MERGEFIELDs named {enumerator.Current.Key}");
    	}
        finally { if (enumerator != null) enumerator.close(); }

        // We can also remove some or all of the elements
        mappedDataFields.remove("MergeFieldName");
        Assert.assertFalse(mappedDataFields.containsKey("MergeFieldName"));
        Assert.assertFalse(mappedDataFields.containsValue("DataSourceColumnName"));

        mappedDataFields.clear();
        Assert.assertEquals(0, mappedDataFields.getCount());

        // Removing the mapped key/value pairs has no effect on the document because the merge was already done with them in place
        doc.save(getArtifactsDir() + "MailMerge.MappedDataFieldCollection.docx");
        TestUtil.mailMergeMatchesDataTable(dataTable, new Document(getArtifactsDir() + "MailMerge.MappedDataFieldCollection.docx"), true); //ExSkip
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
        //ExSummary:Shows how to get MailMergeRegionInfo and work with it.
        Document doc = new Document(getMyDir() + "Mail merge regions.docx");

        // Returns a full hierarchy of regions (with fields) available in the document
        MailMergeRegionInfo regionInfo = doc.getMailMerge().getRegionsHierarchy();

        // Get top regions in the document
        ArrayList<MailMergeRegionInfo> topRegions = regionInfo.getRegions();
        Assert.assertEquals(2, topRegions.size());
        Assert.assertEquals("Region1", topRegions.get(0).getName());
        Assert.assertEquals("Region2", topRegions.get(1).getName());
        Assert.assertEquals(1, topRegions.get(0).getLevel());
        Assert.assertEquals(1, topRegions.get(1).getLevel());

        // Get nested region in first top region
        ArrayList<MailMergeRegionInfo> nestedRegions = topRegions.get(0).getRegions();
        Assert.assertEquals(2, nestedRegions.size());
        Assert.assertEquals("NestedRegion1", nestedRegions.get(0).getName());
        Assert.assertEquals("NestedRegion2", nestedRegions.get(1).getName());
        Assert.assertEquals(2, nestedRegions.get(0).getLevel());
        Assert.assertEquals(2, nestedRegions.get(1).getLevel());

        // Get field list in first top region
        ArrayList<Field> fieldList = topRegions.get(0).getFields();
        Assert.assertEquals(4, fieldList.size());

        FieldMergeField startFieldMergeField = nestedRegions.get(0).getStartField();
        Assert.assertEquals("TableStart:NestedRegion1", startFieldMergeField.getFieldName());

        FieldMergeField endFieldMergeField = nestedRegions.get(0).getEndField();
        Assert.assertEquals("TableEnd:NestedRegion1", endFieldMergeField.getFieldName());
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

        Assert.assertEquals(1, mailMergeCallbackStub.getTagsReplacedCounter());
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

    @Test (dataProvider = "unconditionalMergeFieldsAndRegionsDataProvider")
    public void unconditionalMergeFieldsAndRegions(boolean doCountAllMergeFields) throws Exception
    {
        //ExStart
        //ExFor:MailMerge.UnconditionalMergeFieldsAndRegions
        //ExSummary:Shows how to merge fields or regions regardless of the parent IF field's condition.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a MERGEFIELD nested inside an IF field
        // Since the statement of the IF field is false, the result of the inner MERGEFIELD will not be displayed
        // and the MERGEFIELD will not receive any data during a mail merge
        FieldIf fieldIf = (FieldIf)builder.insertField(" IF 1 = 2 ");
        builder.moveTo(fieldIf.getSeparator());
        builder.insertField(" MERGEFIELD  FullName ");

        // We can still count MERGEFIELDs inside IF fields with false statements if we set this flag to true
        doc.getMailMerge().setUnconditionalMergeFieldsAndRegions(doCountAllMergeFields);

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FullName");
        dataTable.getRows().add("James Bond");

        // Execute the mail merge
        doc.getMailMerge().execute(dataTable);

        // The result will not be visible in the document because the IF field is false, but the inner MERGEFIELD did indeed receive data
        doc.save(getArtifactsDir() + "MailMerge.UnconditionalMergeFieldsAndRegions.docx");

        if (doCountAllMergeFields)
            Assert.assertEquals("\u0013 IF 1 = 2 \"James Bond\"\u0014\u0015", msString.trim(doc.getText()));
        else
            Assert.assertEquals("\u0013 IF 1 = 2 \u0013 MERGEFIELD  FullName \u0014«FullName»\u0015\u0014\u0015", msString.trim(doc.getText()));
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
        //ExSummary:Shows how to execute an Office Data Source Object mail merge with MailMergeSettings.
        // We'll create a simple document that will act as a destination for mail merge data
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField("MERGEFIELD FirstName", "<FirstName>");
        builder.write(" ");
        builder.insertField("MERGEFIELD LastName", "<LastName>");
        builder.writeln(": ");
        builder.insertField("MERGEFIELD Message", "<Message>");

        // We will use an ASCII file as a data source
        // We can use any character we want as a delimiter, in this case we'll choose '|'
        // The delimiter character is selected in the ODSO settings of mail merge settings
        String[] lines = { "FirstName|LastName|Message",
            "John|Doe|Hello! This message was created with Aspose Words mail merge." };
        String dataSrcFilename = getArtifactsDir() + "MailMerge.MailMergeSettings.DataSource.txt";

        File.writeAllLines(dataSrcFilename, lines);

        // Set the data source, query and other things
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

        // Office Data Source Object settings
        Odso odso = settings.getOdso();
        odso.setDataSource(dataSrcFilename);
        odso.setDataSourceType(OdsoDataSourceType.TEXT);
        odso.setColumnDelimiter('|');
        odso.setFirstRowContainsColumnNames(true);

        // ODSO/MailMergeSettings objects can also be cloned
        Assert.assertNotSame(odso, odso.deepClone());
        Assert.assertNotSame(settings, settings.deepClone());

        // The mail merge will be performed when this document is opened 
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

        // We can clear the settings, which will take place during saving
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
        //ExSummary:Shows how to execute a mail merge while drawing data from a header and a data file.
        // Create a mailing label merge header file, which will consist of a table with one row 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.write("FirstName");
        builder.insertCell();
        builder.write("LastName");
        builder.endTable();

        doc.save(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx");

        // Create a mailing label merge date file, which will consist of a table with one row and the same amount of columns as 
        // the header table, which will determine the names for these columns
        doc = new Document();
        builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();
        builder.write("John");
        builder.insertCell();
        builder.write("Doe");
        builder.endTable();

        doc.save(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx");

        // Create a merge destination document with MERGEFIELDS that will accept data
        doc = new Document();
        builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField("MERGEFIELD FirstName", "<FirstName>");
        builder.write(" ");
        builder.insertField("MERGEFIELD LastName", "<LastName>");

        // Configure settings to draw data and headers from other documents
        MailMergeSettings settings = doc.getMailMergeSettings();

        // The "header" document contains column names for the data in the "data" document,
        // which will correspond to the names of our MERGEFIELDs
        settings.setHeaderSource(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx");
        settings.setDataSource(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx");

        // Configure the rest of the MailMergeSettings object
        settings.setQuery("SELECT * FROM " + settings.getDataSource());
        settings.setMainDocumentType(MailMergeMainDocumentType.MAILING_LABELS);
        settings.setDataType(MailMergeDataType.TEXT_FILE);
        settings.setLinkToQuery(true);
        settings.setViewMergedData(true);

        // The mail merge will be performed when this document is opened 
        doc.save(getArtifactsDir() + "MailMerge.MailingLabelMerge.docx");
        //ExEnd

        Assert.assertEquals("FirstName\u0007LastName\u0007\u0007",
            msString.trim(new Document(getArtifactsDir() + "MailMerge.MailingLabelMerge.Header.docx").
                getChild(NodeType.TABLE, 0, true).getText()));

        Assert.assertEquals("John\u0007Doe\u0007\u0007",
            msString.trim(new Document(getArtifactsDir() + "MailMerge.MailingLabelMerge.Data.docx").
                getChild(NodeType.TABLE, 0, true).getText()));

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

        // This collection defines how columns from an external data source will be mapped to predefined MERGEFIELD,
        // ADDRESSBLOCK and GREETINGLINE fields during a mail merge
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

        // Elements of the collection can be cloned
        msAssert.areNotEqual(dataCollection.get(0), dataCollection.get(0).deepClone());

        // The collection can have individual entries removed or be cleared like this
        dataCollection.removeAt(0);
        Assert.assertEquals(29, dataCollection.getCount()); //ExSkip
        dataCollection.clear();
        Assert.assertEquals(0, dataCollection.getCount()); //ExSkip
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
        //ExSummary:Shows how to access the collection of data that designates merge data source records to be excluded from a merge.
        Document doc = new Document(getMyDir() + "Odso data.docx");

        // Records in this collection that do not have the "Active" flag set to true will be excluded from the mail merge
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

        // Elements of the collection can be cloned
        msAssert.areNotEqual(dataCollection.get(0), dataCollection.get(0).deepClone());

        // The collection can have individual entries removed or be cleared like this
        dataCollection.removeAt(0);
        Assert.assertEquals(69, dataCollection.getCount()); //ExSkip
        dataCollection.clear();
        Assert.assertEquals(0, dataCollection.getCount()); //ExSkip
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
        //ExSummary:Shows how to specify where the culture used for date formatting during a field update or mail merge is sourced from.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two merge fields with German locale.
        builder.getFont().setLocaleId(1031);
        builder.insertField("MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder.write(" - ");
        builder.insertField("MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");

        // Set the current culture to US English after preserving its original value in a variable.
        msCultureInfo currentCulture = CurrentThread.getCurrentCulture();
        CurrentThread.setCurrentCulture(new msCultureInfo("en-US"));

        // This merge will use the current thread's culture to format the date, which will be US English.
        doc.getMailMerge().execute(new String[] { "Date1" }, new Object[] { new DateTime(2020, 1, 1) });

        // Configure the next merge to source its culture value from the field code. The value of that culture will be German.
        doc.getFieldOptions().setFieldUpdateCultureSource(FieldUpdateCultureSource.FIELD_CODE);
        doc.getMailMerge().execute(new String[] { "Date2" }, new Object[] { new DateTime(2020, 1, 1) });

        // The first merge result contains a date formatted in English, while the second one is in German.
        Assert.assertEquals("Wednesday, 1 January 2020 - Mittwoch, 1 Januar 2020", msString.trim(doc.getRange().getText()));

        // Restore the original culture.
        CurrentThread.setCurrentCulture(currentCulture);
        //ExEnd
    }
}
