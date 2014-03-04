/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package mailmergeandreporting.multipledocsinmailmerge.java;

import java.io.File;
import java.net.URI;
import java.sql.*;
import java.text.MessageFormat;
import java.util.Hashtable;

import com.aspose.words.*;


public class MultipleDocsInMailMerge
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/mailmergeandreporting/multipledocsinmailmerge/data/";

        produceMultipleDocuments(dataDir, "TestFile.doc");
    }

    public static void produceMultipleDocuments(String dataDir, String srcDoc) throws Exception
    {
        // Open the database connection.
        ResultSet rs = getData(dataDir, "SELECT * FROM Customers");

        // Open the template document.
        Document doc = new Document(dataDir + srcDoc);

        // A record of how many documents that have been generated so far.
        int counter = 1;

        // Loop though all records in the data source.
        while(rs.next())
        {
            // Clone the template instead of loading it from disk (for speed).
            Document dstDoc = (Document)doc.deepClone(true);

            // Extract the data from the current row of the ResultSet into a Hashtable.
            Hashtable dataMap = getRowData(rs);

            // Execute mail merge.
            dstDoc.getMailMerge().execute(keySetToArray(dataMap), dataMap.values().toArray());

            // Save the document.
            dstDoc.save(MessageFormat.format(dataDir + "TestFile Out {0}.doc", counter++));
        }
    }

    /**
	 * Creates a Hashtable from the name and value of each column in the current row of the ResultSet.
	 */
    public static Hashtable getRowData(ResultSet rs) throws Exception
    {
        ResultSetMetaData metaData = rs.getMetaData();
        Hashtable values = new Hashtable();

        for(int i = 1; i <= metaData.getColumnCount(); i++)
        {
            values.put(metaData.getColumnName(i), rs.getObject(i));
        }

        return values;
    }

    /**
	 * Utility function that returns the keys of a Hashtable as an array of Strings.
	 */
    public static String[] keySetToArray(Hashtable table)
    {
        return (String[])table.keySet().toArray(new String[table.size()]);
    }

    /**
	 * Utility function that creates a connection to the Database.
	 */
    public static ResultSet getData(String dataDir, String query) throws Exception
	{
		// Load a DB driver that is used by the demos
		Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		// Compose connection string.
		String connectionString = "jdbc:odbc:DRIVER={Microsoft Access Driver (*.mdb)};" +
								  "DBQ=" + new File(dataDir, "Customers.mdb") + ";UID=Admin";
		// DSN-less DB connection.
		Connection connection = DriverManager.getConnection(connectionString);

        Statement statement = connection.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);

        return statement.executeQuery(query);
	}
}
//ExEnd