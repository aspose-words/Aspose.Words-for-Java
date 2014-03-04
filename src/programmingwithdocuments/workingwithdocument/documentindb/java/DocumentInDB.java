/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithdocument.documentindb.java;

import com.aspose.words.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.net.URI;
import java.sql.*;
import java.text.MessageFormat;


public class DocumentInDB
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithdocument/documentindb/data/";
        String dbName = dataDir + "DocDB.mdb";
        final String FILE_NAME = "TestFile.doc";

        // Create a connection to the database.
        createConnection(dbName);

        // Open the document.
        Document doc = new Document(dataDir + FILE_NAME);

        //ExStart
        //ExId:DocumentInDB_Main
        //ExSummary:Stores the document to a database, then reads the same document back again, and finally deletes the record containing the document from the database.
        // Store the document to the database.
        storeToDatabase(doc);
        // Read the document from the database and store the file to disk.
        Document dbDoc = readFromDatabase(FILE_NAME);

        // Save the retrieved document to disk.
        String newFileName = new File(FILE_NAME).getName() + " from DB" + FILE_NAME.substring(FILE_NAME.lastIndexOf("."));
        dbDoc.save(dataDir + newFileName);

        // Delete the document from the database.
        deleteFromDatabase(FILE_NAME);
        //ExEnd
    }

    /**
     * Stores a document object to the specified database.
     *
     * @param doc The source document.
     */
    //ExStart
    //ExId:DocumentInDB_StoreToDB
    //ExSummary:Stores the document to the specified database.
    public static void storeToDatabase(Document doc) throws Exception
    {
        // Save the document to a OutputStream object.
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        doc.save(outputStream, SaveFormat.DOC);

        // Get the filename from the document.
        String fileName = new File(doc.getOriginalFileName()).getName();

        // Create the SQL command.
        String commandString = "INSERT INTO Documents (FileName, FileContent) VALUES(?, ?)";

        // Prepare the statement to store the data into the database.
        PreparedStatement statement = mConnection.prepareStatement(commandString);

        // Add the parameter value for FileName.
        statement.setString(1, fileName);

        // Add the parameter value for FileContent.
        statement.setBinaryStream(2, new ByteArrayInputStream(outputStream.toByteArray()), outputStream.size());

        // Execute and commit the changes.
        statement.execute();
        mConnection.commit();
    }
    //ExEnd

    /**
     * Retreives a document from the specified database and saves it to disk.
     *
     * @param fileName The name of the document file.
     */
    //ExStart
    //ExId:DocumentInDB_ReadFromDB
    //ExSummary:Retrieves and returns the document from the specified database using the filename as a key to fetch the document.
    public static Document readFromDatabase(String fileName) throws Exception
    {
        // Create the SQL command.
        String commandString = "SELECT * FROM Documents WHERE FileName='" + fileName + "'";

        // Retrieve the results from the database.
        ResultSet resultSet = executeQuery(commandString);

        // Check there was a matching record found from the database and throw an exception if no record was found.
        if(!resultSet.isBeforeFirst())
            throw new IllegalArgumentException(MessageFormat.format("Could not find any record matching the document \"{0}\" in the database.", fileName));

        // Move to the first record.
        resultSet.next();

        // The document is stored in byte form in the FileContent column.
        // Retrieve these bytes of the first matching record to a new buffer.
        byte[] buffer = resultSet.getBytes("FileContent");

        // Wrap the bytes from the buffer into a new ByteArrayInputStream object.
        ByteArrayInputStream newStream = new ByteArrayInputStream(buffer);

        // Read the document from the input stream.
        Document doc = new Document(newStream);

        // Return the retrieved document.
        return doc;

    }
    //ExEnd

    /**
     * Deletes the records containing the specified document name from the database.
     *
     * @param fileName The name of the document file.
     */
    //ExStart
    //ExId:DocumentInDB_DeleteFromDB
    //ExSummary:Delete the document from the database, using filename to fetch the record.
    public static void deleteFromDatabase(String fileName) throws Exception
    {
        // Create the SQL command.
        String commandString = "DELETE * FROM Documents WHERE FileName='" + fileName + "'";

        // Execute the command.
        createStatement().executeUpdate(commandString);
    }
    //ExEnd

    //ExStart
    //ExId:DocumentInDB_DatabaseHelpers
    //ExSummary:Helper methods used to connect to and execute queries on a database.
    /**
     * Utility function that creates a connection to the Database.
     */
    public static void createConnection(String dataBasePath) throws Exception
    {
        //  Load a DB driver that is used by the demos
        Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");

        // The path to the database on the disk.
        File dataBase = new File(dataBasePath);

        // Compose connection string.
        String connectionString = "jdbc:odbc:DRIVER={Microsoft Access Driver (*.mdb)};" +
                "DBQ=" + dataBase + ";UID=Admin";
        // Create a connection to the database.
        mConnection = DriverManager.getConnection(connectionString);
    }

    /**
     * Executes a query on the database.
     */
    protected static ResultSet executeQuery(String query) throws Exception
    {
        return createStatement().executeQuery(query);
    }

    /**
     * Creates a new database statement.
     */
    public static Statement createStatement() throws Exception
    {
        return mConnection.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
    }
    //ExEnd

    /*
     * A connection to the database.
     */
    public static Connection mConnection;
}