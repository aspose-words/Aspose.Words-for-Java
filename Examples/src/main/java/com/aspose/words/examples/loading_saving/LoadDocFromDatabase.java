package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;

import java.io.ByteArrayInputStream;

public class LoadDocFromDatabase {
	public static void main(String[] args) throws Exception {

		// ExStart:LoadDocFromDatabase
		// Retrieve the blob from database
		byte[] buffer = new byte[100];
		// Now we have the document in a byte array buffer

		// Create an input steam which uses byte array to read data
		ByteArrayInputStream bin = new ByteArrayInputStream(buffer);
		// Open the doucment from input stream
		Document doc = new Document(bin);
		// ExEnd:LoadDocFromDatabase
	}
}
