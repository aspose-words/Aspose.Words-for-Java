//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;


public class ExMailMerge extends ExBase
{
    @Test
    public void executeArray() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.Execute(String[],Object[])
        //ExFor:ContentDisposition
        //ExId:MailMergeArray
        //ExSummary:Performs a simple insertion of data into merge fields.
        // Open an existing document.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteArray.doc");

        // Fill the fields in the document with user data.
        doc.getMailMerge().execute(
            new String[] {"FullName", "Company", "Address", "Address2", "City"},
            new Object[] {"James Bond", "MI5 Headquarters", "Milbank", "", "London"});

        doc.save(getMyDir() + "MailMerge.ExecuteArray Out.doc");
        //ExEnd
    }

    @Test
    public void executeDataTable() throws Exception
    {
        //ExStart
        //ExFor:Document
        //ExFor:MailMerge
        //ExFor:MailMerge.Execute(DataTable)
        //ExFor:Document.MailMerge
        //ExSummary:Executes mail merge from data stored in a ResultSet.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteDataTable.doc");

        // This example creates a table, but you would normally load table from a database.
        java.sql.ResultSet resultSet = createCachedRowSet(new String[] {"CustomerName", "Address"});
        addRow(resultSet, new String[] {"Thomas Hardy", "120 Hanover Sq., London"});
        addRow(resultSet, new String[] {"Paolo Accorti", "Via Monte Bianco 34, Torino"});
        com.aspose.words.DataTable table = new com.aspose.words.DataTable(resultSet, "Test");

        // Field values from the table are inserted into the mail merge fields found in the document.
        doc.getMailMerge().execute(table);

        doc.save(getMyDir() + "MailMerge.ExecuteDataTable Out.doc");
        //ExEnd
    }

    @Test
    public void executeWithRegionsDataSet() throws Exception
    {
        //ExStart
        //ExFor:MailMerge.ExecuteWithRegions(DataSet)
        //ExSummary:Executes a mail merge with repeatable regions from an ADO.NET DataSet.
        // Open the document.
        // For a mail merge with repeatable regions, the document should have mail merge regions
        // in the document designated with MERGEFIELD TableStart:MyTableName and TableEnd:MyTableName.
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteWithRegions.doc");

        int orderId = 10444;

        // Populate tables and add them to the dataset.
        // For a mail merge with repeatable regions, DataTable.TableName should be
        // set to match the name of the region defined in the document.
        com.aspose.words.DataSet dataSet = new com.aspose.words.DataSet();

        com.aspose.words.DataTable orderTable = getTestOrder(orderId);
        dataSet.getTables().add(orderTable);

        com.aspose.words.DataTable orderDetailsTable = getTestOrderDetails(orderId, "ProductID");
        dataSet.getTables().add(orderDetailsTable);

        // This looks through all mail merge regions inside the document and for each
        // region tries to find a DataTable with a matching name inside the DataSet.
        // If a table is found, its content is merged into the mail merge region in the document.
        doc.getMailMerge().executeWithRegions(dataSet);

        doc.save(getMyDir() + "MailMerge.ExecuteWithRegionsDataSet Out.doc");
        //ExEnd
    }

    @Test
    public void executeWithRegionsDataTableCaller() throws Exception
    {
        executeWithRegionsDataTable();
    }

    //ExStart
    //ExFor:Document.MailMerge
    //ExFor:MailMerge.ExecuteWithRegions(DataTable)
    //ExId:MailMergeRegions
    //ExSummary:Executes a mail merge with repeatable regions.
    public void executeWithRegionsDataTable() throws Exception
    {
        Document doc = new Document(getMyDir() + "MailMerge.ExecuteWithRegions.doc");

        int orderId = 10444;

        // Perform several mail merge operations populating only part of the document each time.

        // Use DataTable as a data source.
        // The table name property should be set to match the name of the region defined in the document.
        com.aspose.words.DataTable orderTable = getTestOrder(orderId);
        doc.getMailMerge().executeWithRegions(orderTable);

        com.aspose.words.DataTable orderDetailsTable = getTestOrderDetails(orderId, "ExtendedPrice DESC");
        doc.getMailMerge().executeWithRegions(orderDetailsTable);

        doc.save(getMyDir() + "MailMerge.ExecuteWithRegionsDataTable Out.doc");
    }

    private static com.aspose.words.DataTable getTestOrder(int orderId) throws Exception
    {
        java.sql.ResultSet resultSet = executeDataTable(java.text.MessageFormat.format(
            "SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", Integer.toString(orderId)));

        return new com.aspose.words.DataTable(resultSet, "Orders");
    }

    private static com.aspose.words.DataTable getTestOrderDetails(int orderId, String orderBy) throws Exception
    {
        StringBuilder builder = new StringBuilder();

        builder.append(java.text.MessageFormat.format(
            "SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0}", Integer.toString(orderId)));

        if ((orderBy != null) && (orderBy.length() > 0))
        {
            builder.append(" ORDER BY ");
            builder.append(orderBy);
        }

        java.sql.ResultSet resultSet = executeDataTable(builder.toString());
        return new com.aspose.words.DataTable(resultSet, "OrderDetails");
    }

    /**
     * Utility function that creates a connection, command,
     * executes the command and return the result in a DataTable.
     */
    private static java.sql.ResultSet executeDataTable(String commandText) throws Exception
    {
        Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");// Loads the driver

        // Open the database connection.
        String connString = "jdbc:odbc:DRIVER={Microsoft Access Driver (*.mdb)};" +
            "DBQ=" + getDatabaseDir() + "Northwind.mdb" + ";UID=Admin";

        // From Wikipedia: The Sun driver has a known issue with character encoding and Microsoft Access databases.
        // Microsoft Access may use an encoding that is not correctly translated by the driver, leading to the replacement
        // in strings of, for example, accented characters by question marks.
        //
        // In this case I have to set CP1252 for the european characters to come through in the data values.
        java.util.Properties props = new java.util.Properties();
        props.put("charSet", "Cp1252");

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString, props);

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        return statement.executeQuery(commandText);
    }
    //ExEnd

    @Test
    public void mappedDataFields() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.MappedDataFields
        //ExFor:MappedDataFieldCollection
        //ExFor:MappedDataFieldCollection.Add
        //ExId:MailMergeMappedDataFields
        //ExSummary:Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
        doc.getMailMerge().getMappedDataFields().add("MyFieldName_InDocument", "MyFieldName_InDataSource");
        //ExEnd
    }

    @Test
    public void getFieldNames() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.GetFieldNames
        //ExId:MailMergeGetFieldNames
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
        //ExId:MailMergeDeleteFields
        //ExSummary:Shows how to delete all merge fields from a document.
        doc.getMailMerge().deleteFields();
        //ExEnd
    }

    @Test
    public void removeEmptyParagraphs() throws Exception
    {
        Document doc = new Document();
        //ExStart
        //ExFor:MailMerge.RemoveEmptyParagraphs
        //ExId:MailMergeRemoveEmptyParagraphs
        //ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
        doc.getMailMerge().setRemoveEmptyParagraphs(true);
        //ExEnd
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
}

