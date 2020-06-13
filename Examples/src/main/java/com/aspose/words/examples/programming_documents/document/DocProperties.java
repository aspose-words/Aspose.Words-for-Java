package com.aspose.words.examples.programming_documents.document;

import java.util.Date;

import com.aspose.words.CustomDocumentProperties;
import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.examples.Utils;

public class DocProperties {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getDataDir(DocProperties.class);

		EnumerateProperties(dataDir);
		CustomAdd(dataDir);
		CustomRemove(dataDir);
		ConfiguringLinkToContent(dataDir);
	}

	private static void EnumerateProperties(String dataDir) throws Exception {
		// ExStart:EnumerateProperties
		String fileName = dataDir + "Properties.doc";
		Document doc = new Document(fileName);

		System.out.println("1. Document name: " + fileName);

		System.out.println("2. Built-in Properties");
		for (DocumentProperty prop : doc.getBuiltInDocumentProperties())
			System.out.println(prop.getName() + " : " + prop.getValue());

		System.out.println("3. Custom Properties");
		for (DocumentProperty prop : doc.getCustomDocumentProperties())
			System.out.println(prop.getName() + " : " + prop.getValue());
		// ExEnd:EnumerateProperties
	}

	private static void CustomAdd(String dataDir) throws Exception {
		// ExStart:CustomAdd
		Document doc = new Document(dataDir + "Properties.doc");

		CustomDocumentProperties props = doc.getCustomDocumentProperties();

		if (props.get("Authorized") == null) {
			props.add("Authorized", true);
			props.add("Authorized By", "John Smith");
			props.add("Authorized Date", new Date());
			props.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
			props.add("Authorized Amount", 123.45);
		}
		// ExEnd:CustomAdd
	}

	private static void CustomRemove(String dataDir) throws Exception {
		// ExStart:CustomRemove
		Document doc = new Document(dataDir + "Properties.doc");
		doc.getCustomDocumentProperties().remove("Authorized Date");
		// ExEnd:CustomRemove
	}

	private static void ConfiguringLinkToContent(String dataDir) throws Exception {
		// ExStart:ConfiguringLinkToContent
		Document doc = new Document(dataDir + "test.docx");

		// Retrieve a list of all custom document properties from the file.
		CustomDocumentProperties customProperties = doc.getCustomDocumentProperties();

		// Add linked to content property.
		DocumentProperty customProperty = customProperties.addLinkToContent("PropertyName", "BookmarkName");

		// Also, accessing the custom document property can be performed by using the
		// property name.
		customProperty = customProperties.get("PropertyName");

		// Check whether the property is linked to content.
		boolean isLinkedToContent = customProperty.isLinkToContent();

		// Get the source of the property.
		String source = customProperty.getLinkSource();

		// Get the value of the property.
		String value = customProperty.getValue().toString();
		// ExEnd:ConfiguringLinkToContent

		System.out.println("\nIs Linked To Content: " + isLinkedToContent);
		System.out.println("\nLink Source: " + source);
		System.out.println("\nProperty Value: " + value);
	}
}
