/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package mailmergeandreporting.nestedmailmerge.java;

import java.io.File;
import java.net.URI;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import com.aspose.words.*;
import com.sun.rowset.CachedRowSetImpl;


public class NestedMailMerge
{
    //ExStart
    //ExId:NestedMailMerge
    //ExSummary:Shows how to generate an invoice using nested mail merge regions.
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/mailmergeandreporting/nestedmailmerge/data/";

        // Create the dataset which will hold each DataTable used for mail merge.
        DataSet pizzaDs = new DataSet();

        // Create a connection to the database
        createConnection(dataDir);

        // Populate each DataTable from the database. Each query which return a ResultSet object containing the data from the table.
        // This ResultSet is wrapped into an Aspose.Words implementation of the DataTable class and added to a DataSet.
        DataTable orders = new DataTable(executeQuery("SELECT * from Orders"), "Orders");
        pizzaDs.getTables().add(orders);

        DataTable itemDetails = new DataTable(executeQuery("SELECT * from Items"), "Items");
        pizzaDs.getTables().add(itemDetails);

        // In order for nested mail merge to work, the mail merge engine must know the relation between parent and child tables.
        // Add a DataRelation to specify relations between these tables.
        pizzaDs.getRelations().add(new DataRelation(
                "OrderToItemDetails",
                orders,
                itemDetails,
                new String[]{"OrderID"},
                new String[]{"OrderID"}));

        // Open the template document.
        Document doc = new Document(dataDir + "Invoice Template.doc");

        // Execute nested mail merge with regions
        doc.getMailMerge().executeWithRegions(pizzaDs);

        // Save the output to disk
        doc.save(dataDir + "Invoice Out.doc");

        assert doc.getMailMerge().getFieldNames().length == 0 : "There was a problem with mail merge"; //ExSkip
    }

    /**
     * Executes a query to the demo database using a new statement and returns the result in a ResultSet.
     */
    protected static ResultSet executeQuery(String query) throws Exception
    {
        return createStatement().executeQuery(query);
    }

    /**
     * Utility function that creates a connection to the Database.
     */
    public static void createConnection(String dataDir) throws Exception
    {
        //  Load a DB driver that is used by the demos
        Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");

        // The path to the database on the disk.
        File dataBase = new File(dataDir, "InvoiceDB.mdb");

        // Compose connection string.
        String connectionString = "jdbc:odbc:DRIVER={Microsoft Access Driver (*.mdb)};" +
                "DBQ=" + dataBase + ";UID=Admin";
        // Create a connection to the database.
        mConnection = DriverManager.getConnection(connectionString);
    }

    private static Connection mConnection;
    //ExEnd

    //ExStart
    //ExId:CreateDatabaseStatement
    //ExSummary:Shows how to create a connection to a database which is scrollable.
    /**
     * Utility function that creates a statement to the database.
     */
    public static Statement createStatement() throws Exception
    {
        return mConnection.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
    }
    //ExEnd

    protected static void cachedRowSetExample() throws Exception
    {
        //ExStart
        //ExId:CreateCachedRowSet
        //ExSummary:Shows how to store the results from a database query in a CachedRowSet so the database connection can be closed.
        ResultSet resultSet = createStatement().executeQuery("SELECT * FROM Orders");
        CachedRowSetImpl cached = new CachedRowSetImpl();
        // This loads the data into a CachedResultSet. The connection can be closed after this line.
        cached.populate(resultSet);

        // Load the cached data into a new DataTable.
        DataTable orders = new DataTable(cached, "Orders");
        //ExEnd
    }

    public static void createRelationship() throws Exception
	{
	     DataSet dataSet = new DataSet();
	     DataTable orderTable = new DataTable(null, "Orders");
	     DataTable itemTable = new DataTable(null, "Items");
	     //ExStart
	     //ExId:NestedMailMergeCreateRelationship
	     //ExSummary:Shows how to create a simple DataRelation for use in nested mail merge.
	     dataSet.getRelations().add(new DataRelation("OrderToItem", orderTable, itemTable, new String[] {"Order_Id"}, new String[] {"Order_Id"}));
	     //ExEnd
    }
}