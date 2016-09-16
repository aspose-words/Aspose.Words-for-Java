package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.document.properties.AccessingDocumentProperties;

public class TrackChanges {

	public static final String dataDir = Utils.getSharedDataDir(AccessingDocumentProperties.class) + "Document/";

	public static void main(String[] args) throws Exception {
		
		Document doc = new Document(dataDir + "Document.doc");

		// Start tracking and make some revisions.
		doc.startTrackRevisions("Author");
		doc.getFirstSection().getBody().appendParagraph("Hello world!");

		// Revisions will now show up as normal text in the output document.
		doc.acceptAllRevisions();
		
		doc.save(dataDir + "Document.AcceptedRevisions_out_.doc");
	}

}
