package com.aspose.words.examples.programming_documents.document.properties;

import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.examples.Utils;

public class AccessingDocumentProperties {
	
	public static final String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";
	
	public static void main(String[] args) throws Exception {
		String fileName = dataDir + "Properties.doc";
		Document doc = new Document(fileName);

		System.out.println("1. Document name: " + fileName);

		System.out.println("2. Built-in Properties");
		for (DocumentProperty prop : doc.getBuiltInDocumentProperties())
		    System.out.println(prop.getName() + " : " + prop.getValue());

		System.out.println("3. Custom Properties");
		for (DocumentProperty prop : doc.getCustomDocumentProperties())
		    System.out.println(prop.getName() + " : " + prop.getValue());
	}

}
