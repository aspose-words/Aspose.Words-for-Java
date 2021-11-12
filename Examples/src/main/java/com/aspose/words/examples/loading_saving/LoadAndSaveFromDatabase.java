package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Blob;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;

public class LoadAndSaveFromDatabase {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(LoadAndSaveFromDatabase.class);
		String fileName = "Test File (doc).doc";
		// Open the document.
		Document doc = new Document(dataDir + fileName);

		// ExStart: OpenDatabaseConnection
		String url1 = "jdbc:mysql://localhost:3306/test";
		String user = "root";
		String password = "123";

		// Open a database connection.
		Connection mConnection = DriverManager.getConnection(url1, user, password);
		// ExEnd: OpenDatabaseConnection
		System.out.println("Database Connection Successfull.");

		// ExStart:OpenRetrieveAndDelete
		// Store the document to the database.
		StoreToDatabase(doc, mConnection);

		// Read the document from the database and store the file to disk.
		Document dbDoc = ReadFromDatabase(dataDir + fileName, mConnection);

		// Save the retrieved document to disk.
		dbDoc.save(dataDir + fileName);

		// Delete the document from the database.
		DeleteFromDatabase(dataDir + fileName, mConnection);

		// Close the connection to the database.
		mConnection.close();
		// ExEnd:OpenRetrieveAndDelete
	}

	// ExStart: DeleteFromDatabase
	private static void DeleteFromDatabase(String fileName, Connection mConnection) throws Exception {
		// Create the SQL command.
		String commandString = "DELETE FROM Documents WHERE FileName='" + fileName + "'";
		Statement statement = mConnection.createStatement();
		// Delete the record.
		statement.execute(commandString);
	}
	// ExEnd: DeleteFromDatabase

	// ExStart: ReadFromDatabase
	private static Document ReadFromDatabase(String fileName, Connection mConnection) throws Exception {
		// Create the SQL command.
		String commandString = "SELECT * FROM Documents WHERE FileName=?";
		PreparedStatement statement = mConnection.prepareStatement(commandString);
		statement.setString(1, fileName);

		Document doc = null;
		ResultSet result = statement.executeQuery();
		if (result.next()) {
			Blob blob = result.getBlob("FileContent");
			InputStream inputStream = blob.getBinaryStream();
			doc = new Document(inputStream);
			inputStream.close();
			System.out.println("File saved");
		}
		result.close();
		return doc;
	}
	// ExEnd: ReadFromDatabase

	// ExStart: StoreToDatabase
	public static void StoreToDatabase(Document doc, Connection mConnection) throws Exception {
		// Create an output stream which uses byte array to save data
		ByteArrayOutputStream aout = new ByteArrayOutputStream();
		// Save the document to byte array
		doc.save(aout, SaveFormat.DOCX);
		// Get the byte array from output steam
		// the byte array now contains the document
		byte[] buffer = aout.toByteArray();

		// Get the filename from the document.
		String fileName = doc.getOriginalFileName();
		String filePath = fileName.replace("\\", "\\\\");

		// Create the SQL command.
		String commandString = "INSERT INTO Documents (FileName, FileContent) VALUES('" + filePath + "', '" + buffer
				+ "')";
		Statement statement = mConnection.createStatement();
		statement.executeUpdate(commandString);
	}
	// ExEnd: StoreToDatabase

}
