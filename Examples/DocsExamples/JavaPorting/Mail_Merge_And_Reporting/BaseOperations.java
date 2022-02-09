package DocsExamples.Mail_Merge_and_Reporting;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataView;
import java.text.MessageFormat;
import com.aspose.words.net.System.Data.DataSet;
import java.util.ArrayList;
import com.aspose.words.MailMergeRegionInfo;
import org.testng.Assert;


class BaseOperations extends DocsExamplesBase
{
    @Test
    public void simpleMailMerge() throws Exception
    {
        //ExStart:SimpleMailMerge
        // Include the code for our template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create Merge Fields.
        builder.insertField(" MERGEFIELD CustomerName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Item ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Quantity ");

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(new String[] { "CustomerName", "Item", "Quantity" },
            new Object[] { "John Doe", "Hawaiian", "2" });

        doc.save(getArtifactsDir() + "BaseOperations.SimpleMailMerge.docx");
        //ExEnd:SimpleMailMerge
    }

    @Test
    public void useIfElseMustache() throws Exception
    {
        //ExStart:UseOfifelseMustacheSyntax
        Document doc = new Document(getMyDir() + "Mail merge destinations - Mustache syntax.docx");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().execute(new String[] { "GENDER" }, new Object[] { "MALE" });

        doc.save(getArtifactsDir() + "BaseOperations.IfElseMustache.docx");
        //ExEnd:UseOfifelseMustacheSyntax
    }

    @Test
    public void mustacheSyntaxUsingDataTable() throws Exception
    {
        //ExStart:MustacheSyntaxUsingDataTable
        Document doc = new Document(getMyDir() + "Mail merge destinations - Vendor.docx");

        // Loop through each row and fill it with data.
        DataTable dataTable = new DataTable("list");
        dataTable.getColumns().add("Number");
        for (int i = 0; i < 10; i++)
        {
            DataRow dataRow = dataTable.newRow();
            dataTable.getRows().add(dataRow);
            dataRow.set(0, "Number " + i);
        }

        // Activate performing a mail merge operation into additional field types.
        doc.getMailMerge().setUseNonMergeFields(true);

        doc.getMailMerge().executeWithRegions(dataTable);

        doc.save(getArtifactsDir() + "WorkingWithXmlData.MustacheSyntaxUsingDataTable.docx");
        //ExEnd:MustacheSyntaxUsingDataTable
    }

    @Test
    public void executeWithRegionsDataTable() throws Exception
    {
        //ExStart:ExecuteWithRegionsDataTable
        Document doc = new Document(getMyDir() + "Mail merge destinations - Orders.docx");

        // Use DataTable as a data source.
        int orderId = 10444;
        DataTable orderTable = getTestOrder(orderId);
        doc.getMailMerge().executeWithRegions(orderTable);

        // Instead of using DataTable, you can create a DataView for custom sort or filter and then mail merge.
        DataView orderDetailsView = new DataView(getTestOrderDetails(orderId));
        orderDetailsView.setSort("ExtendedPrice DESC");

        // Execute the mail merge operation.
        doc.getMailMerge().executeWithRegions(orderDetailsView);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegions.docx");
        //ExEnd:ExecuteWithRegionsDataTable
    }

    //ExStart:ExecuteWithRegionsDataTableMethods
    private DataTable getTestOrder(int orderId)
    {
        DataTable table = executeDataTable($"SELECT * FROM AsposeWordOrders WHERE OrderId = {orderId}");
        table.setTableName("Orders");
        
        return table;
    }

    private DataTable getTestOrderDetails(int orderId)
    {
        DataTable table = executeDataTable(
            $"SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {orderId} ORDER BY ProductID");
        table.setTableName("OrderDetails");
        
        return table;
    }

    /// <summary>
    /// Utility function that creates a connection, command, executes the command and returns the result in a DataTable.
    /// </summary>
    private DataTable executeDataTable(String commandText)
    {
        String connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";

        OleDbConnection conn = new OleDbConnection(connString);
        conn.Open();

        OleDbCommand cmd = new OleDbCommand(commandText, conn);
        OleDbDataAdapter da = new OleDbDataAdapter(cmd);

        DataTable table = new DataTable();
        da.Fill(table);

        conn.Close();

        return table;
    }
    //ExEnd:ExecuteWithRegionsDataTableMethods

    @Test
    public void produceMultipleDocuments() throws Exception
    {
        //ExStart:ProduceMultipleDocuments
        String connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + getDatabaseDir() + "Northwind.mdb";

        Document doc = new Document(getMyDir() + "Mail merge destination - Northwind suppliers.docx");

        OleDbConnection conn = new OleDbConnection(connString);
        conn.Open();
        
        OleDbCommand cmd = new OleDbCommand("SELECT * FROM Customers", conn);
        OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        
        DataTable data = new DataTable();
        da.Fill(data);

        // Perform a loop through each DataRow to iterate through the DataTable. Clone the template document
        // instead of loading it from disk for better speed performance before the mail merge operation.
        // You can load the template document from a file or stream but it is faster to load the document
        // only once and then clone it in memory before each mail merge operation.
        
        int counter = 1;
        for (DataRow row : (Iterable<DataRow>) data.getRows())
        {
            Document dstDoc = (Document) doc.deepClone(true);

            dstDoc.getMailMerge().execute(row);

            dstDoc.save(MessageFormat.format(getArtifactsDir() + "BaseOperations.ProduceMultipleDocuments_{0}.docx", counter++));
        }
        //ExEnd:ProduceMultipleDocuments
    }

    //ExStart:MailMergeWithRegions
    @Test
    public void mailMergeWithRegions() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The start point of mail merge with regions the dataset.
        builder.insertField(" MERGEFIELD TableStart:Customers");
        
        // Data from rows of the "CustomerName" column of the "Customers" table will go in this MERGEFIELD.
        builder.write("Orders for ");
        builder.insertField(" MERGEFIELD CustomerName");
        builder.write(":");

        // Create column headers.
        builder.startTable();
        builder.insertCell();
        builder.write("Item");
        builder.insertCell();
        builder.write("Quantity");
        builder.endRow();

        // We have a second data table called "Orders", which has a many-to-one relationship with "Customers"
        // picking up rows with the same CustomerID value.
        builder.insertCell();
        builder.insertField(" MERGEFIELD TableStart:Orders");
        builder.insertField(" MERGEFIELD ItemName");
        builder.insertCell();
        builder.insertField(" MERGEFIELD Quantity");
        builder.insertField(" MERGEFIELD TableEnd:Orders");
        builder.endTable();

        // The end point of mail merge with regions.
        builder.insertField(" MERGEFIELD TableEnd:Customers");

        // Pass our dataset to perform mail merge with regions.          
        DataSet customersAndOrders = createDataSet();
        doc.getMailMerge().executeWithRegions(customersAndOrders);

        doc.save(getArtifactsDir() + "BaseOperations.MailMergeWithRegions.docx");
    }
    //ExEnd:MailMergeWithRegions

    //ExStart:CreateDataSet
    private DataSet createDataSet()
    {
        // Create the customers table.
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
        tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

        // Create the orders table.
        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
        tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
        tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

        // Add both tables to a data set.
        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);

        // The "CustomerID" column, also the primary key of the customers table is the foreign key for the Orders table.
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        return dataSet;
    }
    //ExEnd:CreateDataSet

    @Test
    public void getRegionsByName() throws Exception
    {
        //ExStart:GetRegionsByName
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
        //ExEnd:GetRegionsByName
    }
}
