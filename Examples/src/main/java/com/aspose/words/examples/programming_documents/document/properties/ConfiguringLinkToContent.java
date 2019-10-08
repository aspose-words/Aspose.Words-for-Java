package com.aspose.words.examples.programming_documents.document.properties;

import com.aspose.words.CustomDocumentProperties;
import com.aspose.words.Document;
import com.aspose.words.DocumentProperty;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.loading_saving.WorkingWithVbaMacros;

public class ConfiguringLinkToContent {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(DocProperties.class);
		
		// ExStart:ConfiguringLinkToContent        
        Document doc = new Document(dataDir + "test.docx");

        // Retrieve a list of all custom document properties from the file.
        CustomDocumentProperties customProperties = doc.getCustomDocumentProperties();

        // Add linked to content property.
        DocumentProperty customProperty = customProperties.addLinkToContent("PropertyName", "BookmarkName");

        // Also, accessing the custom document property can be performed by using the property name.
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
