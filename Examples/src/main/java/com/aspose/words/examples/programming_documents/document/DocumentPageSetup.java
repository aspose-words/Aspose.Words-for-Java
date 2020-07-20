package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.SectionLayoutMode;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 8/10/2017.
 */
public class DocumentPageSetup {

	public static final String dataDir = Utils.getSharedDataDir(DocumentPageSetup.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// ExStart:DocumentPageSetup
		// The path to the documents directory.

		Document doc = new Document(dataDir + "Document.doc");
		// Set the layout mode for a section allowing to define the document grid
		// behavior
		// Note that the Document Grid tab becomes visible in the Page Setup dialog of
		// MS Word if any Asian language is defined as editing language.
		doc.getFirstSection().getPageSetup().setLayoutMode(SectionLayoutMode.GRID);
		// Set the number of characters per line in the document grid.
		doc.getFirstSection().getPageSetup().setCharactersPerLine(30);
		// Set the number of lines per page in the document grid.
		doc.getFirstSection().getPageSetup().setLinesPerPage(10);
		// Save the document
		doc.save(dataDir + "Document.PageSetup_out.doc");
		// ExEnd:DocumentPageSetup
		System.out.println("PageSetup properties are set successfully.");
	}
}
