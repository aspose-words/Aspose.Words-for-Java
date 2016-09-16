package com.aspose.words.examples.programming_documents.document.properties;

import java.util.Date;

import com.aspose.words.CustomDocumentProperties;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class AddOrRemoveDocumentProperties {

	public static final String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// Checks if a custom property with a given name exists in a document and adds few more custom document properties
		addDocumentProperties();
		
		// Remove a custom document property
		removeDocumentProperty();
	}

	public static void addDocumentProperties() throws Exception {
		Document doc = new Document(dataDir + "Properties.doc");

		CustomDocumentProperties props = doc.getCustomDocumentProperties();

		if (props.get("Authorized") == null) {
			props.add("Authorized", true);
			props.add("Authorized By", "John Smith");
			props.add("Authorized Date", new Date());
			props.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
			props.add("Authorized Amount", 123.45);
		}
	}
	
	public static void removeDocumentProperty() throws Exception {
		Document doc = new Document(dataDir + "Properties.doc");
		doc.getCustomDocumentProperties().remove("Authorized Date");
	}

}
