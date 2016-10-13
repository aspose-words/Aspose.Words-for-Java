package com.aspose.words.examples.programming_documents.tableofcontents;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;

public class InsertATableOfContentsField {
	
	private static final String dataDir = Utils.getSharedDataDir(InsertATableOfContentsField.class) + "TableOfContents/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Insert a table of contents at the beginning of the document.
		builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

		// The newly inserted table of contents will be initially empty.
		// It needs to be populated by updating the fields in the document.

        doc.updateFields();
		
		doc.save(dataDir + "InsertATableOfContentsField_out.docx");
	}
}
