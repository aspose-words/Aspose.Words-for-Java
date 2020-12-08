package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class CloneDocument {

	public static final String dataDir = Utils.getSharedDataDir(CloneDocument.class) + "Document/";

	public static void main(String[] args) throws Exception {
		CloneDocument();
		CloneADocument();
	}
	
	public static void CloneDocument() throws Exception {
		// ExStart:CloneDocument
		// Load the document from disk.
		Document doc = new Document(dataDir + "Document.doc");

		Document clone = doc.deepClone();

		// Save the document to disk.
		clone.save(dataDir + "TestFile_clone_out.doc");
		// ExEnd:CloneDocument
	}
	
	public static void CloneADocument() throws Exception {
		// ExStart:CloneADocument
		// Create a document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.writeln("This is the original document before applying the clone method"); 

		// Clone the document.
		Document clone = doc.deepClone();

		// Edit the cloned document.
		builder = new DocumentBuilder(clone);
		builder.write("Section 1");
		builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
		builder.write("Section 2");

		// This shows what is in the document originally. The document has two sections.
		System.out.println(clone.getText().trim());

		// Duplicate the last section and append the copy to the end of the document.
		int lastSectionIdx = clone.getSections().getCount() - 1;
		Section newSection = clone.getSections().get(lastSectionIdx).deepClone();
		clone.getSections().add(newSection);

		// Check what the document contains after we changed it.
		System.out.println(clone.getText().trim());
		//ExEnd:CloneADocument
	}

}
