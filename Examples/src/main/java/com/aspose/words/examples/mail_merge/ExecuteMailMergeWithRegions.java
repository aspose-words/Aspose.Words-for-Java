package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataTable;
import com.aspose.words.MailMergeRegionInfo;
import java.util.*;

public class ExecuteMailMergeWithRegions {
	//ExStart: ExecuteMailMergeWithRegions
    private static final String dataDir = Utils.getSharedDataDir(ExecuteMailMergeWithRegions.class) + "MailMerge/";

    public static void main(String[] args) throws Exception {
    	
        Document doc = new Document(dataDir + "MailMerge.ExecuteWithRegions.doc");

        int orderId = 10444;

        // Perform several mail merge operations populating only part of the document each time.
        // Use DataTable as a data source.
        // The table name property should be set to match the name of the region defined in the document.
        DataTable orderTable = getTestOrder(orderId);
        doc.getMailMerge().executeWithRegions(orderTable);

        DataTable orderDetailsTable = getTestOrderDetails(orderId, "ExtendedPrice DESC");
        doc.getMailMerge().executeWithRegions(orderDetailsTable);

        doc.save(dataDir + "MailMerge.ExecuteWithRegionsDataTable Out.doc");
    }

    private static DataTable getTestOrder(int orderId) throws Exception {
        java.sql.ResultSet resultSet = executeDataTable(java.text.MessageFormat.format("SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", Integer.toString(orderId)));

        return new DataTable(resultSet, "Orders");
    }

    private static DataTable getTestOrderDetails(int orderId, String orderBy) throws Exception {
        StringBuilder builder = new StringBuilder();

        builder.append(java.text.MessageFormat.format("SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0}", Integer.toString(orderId)));

        if ((orderBy != null) && (orderBy.length() > 0)) {
            builder.append(" ORDER BY ");
            builder.append(orderBy);
        }

        java.sql.ResultSet resultSet = executeDataTable(builder.toString());
        return new DataTable(resultSet, "OrderDetails");
    }

    /**
     * Utility function that creates a connection, command, executes the command
     * and return the result in a DataTable.
     */
    private static java.sql.ResultSet executeDataTable(String commandText) throws Exception {
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String connString = "jdbc:ucanaccess://" + dataDir + "Northwind.mdb";

        // From Wikipedia: The Sun driver has a known issue with character encoding and Microsoft Access databases.
        // Microsoft Access may use an encoding that is not correctly translated by the driver, leading to the replacement
        // in strings of, for example, accented characters by question marks.
        //
        // In this case I have to set CP1252 for the European characters to come through in the data values.
        java.util.Properties props = new java.util.Properties();
        props.put("charSet", "Cp1252");

        // DSN-less DB connection.
        java.sql.Connection conn = java.sql.DriverManager.getConnection(connString, props);

        // Create and execute a command.
        java.sql.Statement statement = conn.createStatement();
        return statement.executeQuery(commandText);
    }
    //ExEnd: ExecuteMailMergeWithRegions

    private static void GetRegionsByName() throws Exception {
    	//ExStart: GetRegionsByName
    	Document doc = new Document(dataDir + "Mail merge regions.docx");

    	List<MailMergeRegionInfo> regions = doc.getMailMerge().getRegionsByName("Region1");
    	for (MailMergeRegionInfo region : regions)
    		System.out.println(region.getName());

    	regions = doc.getMailMerge().getRegionsByName("Region2");
    	for (MailMergeRegionInfo region : regions)
    		System.out.println(region.getName());

    	regions = doc.getMailMerge().getRegionsByName("NestedRegion1");
    	for (MailMergeRegionInfo region : regions)
    		System.out.println(region.getName());
    	//ExEnd: GetRegionsByName
    }
}
