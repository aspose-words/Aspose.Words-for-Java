package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class ShowGrammaticalAndSpellingErrors {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ShowGrammaticalAndSpellingErrors.class);

		// ExStart: ShowGrammaticalAndSpellingErrors
		Document doc = new Document(dataDir + "Document.doc");

		doc.setShowGrammaticalErrors(true);
		doc.setShowSpellingErrors(true);

		doc.save(dataDir + "Document.ShowErrorsInDocument_out.docx");
		// ExEnd: ShowGrammaticalAndSpellingErrors
		System.out.println("\nDocument saved successfully.\nFile saved at " + dataDir);
	}

}
