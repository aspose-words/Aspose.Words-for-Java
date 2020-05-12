package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class CreateDocument {
	public static void main(String[] args) throws Exception {
		// ExStart:CreateDocument
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(CreateDocument.class);

		// Load the document.
		Document doc = new Document();

		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.write("hello world");

		doc.save(dataDir + "output.docx");
		// ExEnd:CreateDocument
		System.out.println("Document created successfully.");
	}
}
