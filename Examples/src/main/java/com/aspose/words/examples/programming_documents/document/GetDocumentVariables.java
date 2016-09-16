package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.document.properties.AccessingDocumentProperties;

public class GetDocumentVariables {
	
	public static final String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document(dataDir + "Document.doc");

		for (java.util.Map.Entry entry : doc.getVariables()) {
		    String name = entry.getKey().toString();
		    String value = entry.getValue().toString();

		    // Do something useful.
		    System.out.println("Name: " + name + ", Value: " + value);
		}
	}

}
