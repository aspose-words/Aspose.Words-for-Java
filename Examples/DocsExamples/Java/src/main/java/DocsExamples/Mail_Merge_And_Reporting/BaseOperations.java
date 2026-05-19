package DocsExamples.Mail_Merge_And_Reporting;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.MailMergeRegionInfo;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.ref.Ref;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.sql.*;
import java.text.MessageFormat;
import java.util.ArrayList;

@Test
public class BaseOperations extends DocsExamplesBase {
    @Test
    public void simpleMailMerge() throws Exception {
        //ExStart:ExecuteSimpleMailMerge
        //GistId:c1173266e34062134a6f4178cd531c15
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
        doc.getMailMerge().execute(new String[]{"CustomerName", "Item", "Quantity"},
                new Object[]{"John Doe", "Hawaiian", "2"});

        doc.save(getArtifactsDir() + "BaseOperations.SimpleMailMerge.docx");
        //ExEnd:ExecuteSimpleMailMerge
    }

    @Test
    public void useIfElseMustache() throws Exception {
        //ExStart:UseIfElseMustache
        //GistId:c0c6109b3d172b64370a9684b6330fe4
        Document doc = new Document(getMyDir() + "Mail merge destinations - Mustache syntax.docx");

        doc.getMailMerge().setUseNonMergeFields(true);
        doc.getMailMerge().execute(new String[]{"GENDER"}, new Object[]{"MALE"});

        doc.save(getArtifactsDir() + "BaseOperations.IfElseMustache.docx");
        //ExEnd:UseIfElseMustache
    }

    @Test
    public void mustacheSyntaxUsingDataTable() throws Exception {
        //ExStart:MustacheSyntaxUsingDataTable
        //GistId:c0c6109b3d172b64370a9684b6330fe4
        Document doc = new Document(getMyDir() + "Mail merge destinations - Vendor.docx");

        // Loop through each row and fill it with data.
        DataTable dataTable = new DataTable("list");
        dataTable.getColumns().add("Number");
        for (int i = 0; i < 10; i++) {
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
    public void executeWithRegionsDataTable() throws Exception {
        //ExStart:ExecuteWithRegionsDataTable
        //GistId:cb1cd45863a6c54154294f1778b4f991
        Document doc = new Document(getMyDir() + "Mail merge destinations - Orders.docx");

        // Use custom data source implementation
        int orderId = 10444;

        // Execute mail merge with Orders data
        OrderDataSource orderData = getTestOrder(orderId);
        doc.getMailMerge().executeWithRegions(orderData);

        // Execute mail merge with OrderDetails data (sorted by ExtendedPrice DESC)
        OrderDetailsDataSource orderDetailsData = getTestOrderDetails(orderId, "ExtendedPrice DESC");
        doc.getMailMerge().executeWithRegions(orderDetailsData);

        doc.save(getArtifactsDir() + "MailMerge.ExecuteWithRegions.docx");
        //ExEnd:ExecuteWithRegionsDataTable
    }

    //ExStart:ExecuteWithRegionsDataTableMethods
    private OrderDataSource getTestOrder(int orderId) throws SQLException {
        String sql = "SELECT * FROM AsposeWordOrders WHERE OrderId = " + orderId;
        ResultSet rs = executeQuery(sql);
        return new OrderDataSource(rs, "Orders");
    }

    private OrderDetailsDataSource getTestOrderDetails(int orderId, String sortOrder) throws SQLException {
        String sql = "SELECT * FROM AsposeWordOrderDetails WHERE OrderId = " + orderId +
                " ORDER BY " + (sortOrder != null ? sortOrder.replace(" DESC", " DESC") : "ProductID");
        ResultSet rs = executeQuery(sql);
        return new OrderDetailsDataSource(rs, "OrderDetails");
    }

    /// <summary>
    /// Utility function that creates a connection, command, executes the command and returns the result in a DataTable.
    /// </summary>
    private ResultSet executeQuery(String commandText) throws SQLException {
        String connString = "jdbc:ucanaccess://" + getDatabaseDir() + "Northwind.accdb";
        Connection conn = DriverManager.getConnection(connString, "Admin", "");
        Statement stmt = conn.createStatement();
        return stmt.executeQuery(commandText);
    }
    //ExEnd:ExecuteWithRegionsDataTableMethods

    @Test
    public void produceMultipleDocuments() throws Exception {
        //ExStart:ProduceMultipleDocuments
        //GistId:c1173266e34062134a6f4178cd531c15
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String connString = "jdbc:ucanaccess://" + getDatabaseDir() + "Northwind.accdb";

        Document doc = new Document(getMyDir() + "Mail merge destination - Suppliers.docx");

        Connection connection = DriverManager.getConnection(connString, "Admin", "");

        Statement statement = connection.createStatement();
        ResultSet resultSet = statement.executeQuery("SELECT * FROM Customers");

        DataTable dataTable = new DataTable(resultSet, "Customers");

        // Perform a loop through each DataRow to iterate through the DataTable. Clone the template document
        // instead of loading it from disk for better speed performance before the mail merge operation.
        // You can load the template document from a file or stream but it is faster to load the document
        // only once and then clone it in memory before each mail merge operation.
        int counter = 1;
        for (DataRow row : dataTable.getRows()) {
            Document dstDoc = (Document) doc.deepClone(true);

            dstDoc.getMailMerge().execute(row);

            dstDoc.save(MessageFormat.format(getArtifactsDir() + "BaseOperations.ProduceMultipleDocuments_{0}.docx", counter++));
        }

        connection.close();
        //ExEnd:ProduceMultipleDocuments
    }

    @Test
    public void mailMergeWithRegions() throws Exception {
        //ExStart:MailMergeWithRegions
        //GistId:c1173266e34062134a6f4178cd531c15
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
        //ExEnd:MailMergeWithRegions
    }

    //ExStart:CreateDataSet
    //GistId:c1173266e34062134a6f4178cd531c15
    private DataSet createDataSet() {
        // Create the customers table.
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        // Create the orders table.
        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

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
    public void getRegionsByName() throws Exception {
        //ExStart:GetRegionsByName
        //GistId:7d267f5905ac7f3d5144e8d4126a0d6d
        Document doc = new Document(getMyDir() + "Mail merge regions.docx");

        //ExStart:GetRegionsHierarchy
        //GistId:7d267f5905ac7f3d5144e8d4126a0d6d
        MailMergeRegionInfo regionInfo = doc.getMailMerge().getRegionsHierarchy();
        //ExEnd:GetRegionsHierarchy

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

// Custom data source for Orders.
class OrderDataSource implements com.aspose.words.IMailMergeDataSource {
    private ResultSet resultSet;
    private String tableName;
    private boolean isFirst = true;

    public OrderDataSource(ResultSet rs, String tableName) {
        this.resultSet = rs;
        this.tableName = tableName;
    }

    @Override
    public String getTableName() {
        return tableName;
    }

    @Override
    public boolean moveNext() throws Exception {
        if (isFirst) {
            isFirst = false;
            return resultSet.next();
        }
        return resultSet.next();
    }

    @Override
    public boolean getValue(String fieldName, Ref<Object> fieldValue) {
        try {
            Object value = resultSet.getObject(fieldName);
            fieldValue.set(value);
            return true;
        } catch (SQLException e) {
            return false;
        }
    }

    @Override
    public com.aspose.words.IMailMergeDataSource getChildDataSource(String tableName) {
        return null;
    }
}

// Custom data source for OrderDetails.
class OrderDetailsDataSource implements IMailMergeDataSource {
    private ResultSet resultSet;
    private String tableName;
    private boolean isFirst = true;

    public OrderDetailsDataSource(ResultSet rs, String tableName) {
        this.resultSet = rs;
        this.tableName = tableName;
    }

    @Override
    public String getTableName() {
        return tableName;
    }

    @Override
    public boolean moveNext() throws Exception {
        if (isFirst) {
            isFirst = false;
            return resultSet.next();
        }
        return resultSet.next();
    }

    @Override
    public boolean getValue(String fieldName, Ref<Object> fieldValue) {
        try {
            Object value = resultSet.getObject(fieldName);
            fieldValue.set(value);
            return true;
        } catch (SQLException e) {
            return false;
        }
    }

    @Override
    public com.aspose.words.IMailMergeDataSource getChildDataSource(String tableName) {
        return null;
    }
}
