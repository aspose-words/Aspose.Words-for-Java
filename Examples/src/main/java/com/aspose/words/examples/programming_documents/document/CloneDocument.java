package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.document.properties.AccessingDocumentProperties;

public class CloneDocument {
	
	public static final String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Document.doc");
		Document clone = doc.deepClone();
	}

}
