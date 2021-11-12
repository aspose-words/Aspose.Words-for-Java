package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;

public class LoadAndSave {
	private static Connection mConnection;

	public static void main(String[] args) throws Exception {

		// ExStart:LoadAndSave
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(LoadAndSave.class);

		Document doc = new Document(dataDir+ "Test File (doc).doc");

		// Save the finished document to disk.
		doc.save(dataDir + "Test File (doc)_out.doc", SaveFormat.PNG);
		// ExEnd:LoadAndSave
		System.out.println("Document loaded and saved successfully.");
	}
}
