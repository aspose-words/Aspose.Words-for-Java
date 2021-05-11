package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class WorkingWithStyle {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(WorkingWithStyle.class);

		cleansUnusedStylesandLists(dataDir);
		copyStyles(dataDir);
		CleanupDuplicateStyle(dataDir);
	}

	public static void cleansUnusedStylesandLists(String dataDir) throws Exception {
		// ExStart:CleansUnusedStylesandLists
		Document doc = new Document(dataDir + "TestFile.doc");

		// Count of styles before Cleanup.
		System.out.println(doc.getStyles().getCount());
		// Count of lists before Cleanup.
		System.out.println(doc.getLists().getCount());

		CleanupOptions cleanupoptions = new CleanupOptions();
		cleanupoptions.setUnusedLists(false);
		cleanupoptions.setUnusedStyles(true);

		// Cleans unused styles and lists from the document depending on given
		// CleanupOptions.
		doc.cleanup(cleanupoptions);

		// Count of styles after Cleanup was decreased.
		System.out.println(doc.getStyles().getCount());
		// Count of lists after Cleanup is the same.
		System.out.println(doc.getLists().getCount());

		doc.save(dataDir + "Document.Cleanup_out.docx");
		// ExEnd:CleansUnusedStylesandLists

		System.out.println("Document unused Styles cleaned successfully.");
	}

	public static void copyStyles(String dataDir) throws Exception {
		// ExStart:CopyStylesFromDocument
		Document doc = new Document(dataDir + "template.docx");
		Document target = new Document(dataDir + "TestFile.doc");

		target.copyStylesFromTemplate(doc);

		dataDir = dataDir + "CopyStyles_out.docx";
		doc.save(dataDir);
		// ExEnd:CopyStylesFromDocument
		System.out.println("\\nStyles are copied from document successfully.\\nFile saved at " + dataDir);
	}

	private static void CleanupDuplicateStyle(String dataDir) throws Exception {
		// ExStart:CleanupDuplicateStyle
		Document doc = new Document(dataDir + "Document.doc");

		// Count of styles before Cleanup.
		System.out.println(doc.getStyles().getCount());

		CleanupOptions options = new CleanupOptions();
		options.setDuplicateStyle(true);

		// Cleans duplicate styles from the document.
		doc.cleanup(options);

		// Count of styles after Cleanup was decreased.
		System.out.println(doc.getStyles().getCount());

		doc.save(dataDir + "Document.CleanupDuplicateStyle_out.docx");
		// ExEnd:CleanupDuplicateStyle
		System.out.println("\nAll revisions accepted.\nFile saved at " + dataDir);
	}
}