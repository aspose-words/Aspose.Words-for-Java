package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.examples.Utils;

import java.util.Date;

/**
 * Created by Home on 5/29/2017.
 */
public class CompareTwoWordDocumentswithCompareOptions {

	public static void main(String[] args) throws Exception {

		// ExStart:CompareTwoWordDocumentswithCompareOptions
		String dataDir = Utils.getDataDir(CompareTwoWordDocumentswithCompareOptions.class);

		com.aspose.words.Document docA = new com.aspose.words.Document(dataDir + "TestFile.docx");
		com.aspose.words.Document docB = new com.aspose.words.Document(dataDir + "TestFile2.docx");

		com.aspose.words.CompareOptions options = new com.aspose.words.CompareOptions();
		options.setIgnoreFormatting(true);
		options.setIgnoreHeadersAndFooters(true);
		options.setIgnoreCaseChanges(true);
		options.setIgnoreTables(true);
		options.setIgnoreFields(true);
		options.setIgnoreComments(true);
		options.setIgnoreTextboxes(true);
		options.setIgnoreFootnotes(true);

		
		// DocA now contains changes as revisions.
		docA.compare(docB, "user", new Date(), options);
		if (docA.getRevisions().getCount() == 0)
			System.out.println("Documents are equal");
		else
			System.out.println("Documents are not equal");
		// ExEnd:CompareTwoWordDocumentswithCompareOptions
	}
}
