package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class GetDocumentVariables {

	public static final String dataDir = Utils.getSharedDataDir(GetDocumentVariables.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// ExStart:GetDocumentVariables
		Document doc = new Document(dataDir + "Document.doc");

		for (java.util.Map.Entry entry : doc.getVariables()) {
			String name = entry.getKey().toString();
			String value = entry.getValue().toString();

			// Do something useful.
			System.out.println("Name: " + name + ", Value: " + value);
		}
		// ExEnd:GetDocumentVariables
	}

}
