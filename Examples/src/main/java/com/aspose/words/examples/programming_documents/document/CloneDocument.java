package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class CloneDocument {

	public static final String dataDir = Utils.getSharedDataDir(CloneDocument.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// ExStart:CloneDocument
		// Load the document from disk.
		Document doc = new Document(dataDir + "Document.doc");

		Document clone = doc.deepClone();

		// Save the document to disk.
		clone.save(dataDir + "TestFile_clone_out.doc");
		// ExEnd:CloneDocument
	}

}
