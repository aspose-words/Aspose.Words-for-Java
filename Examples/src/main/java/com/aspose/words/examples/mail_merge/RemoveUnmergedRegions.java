package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;

public class RemoveUnmergedRegions {

	private static final String dataDir = Utils.getSharedDataDir(RemoveUnmergedRegions.class) + "MailMerge/";

	public static void main(String[] args) throws Exception {

		// Open the document.
		Document doc = new Document(dataDir + "TestFile.doc");

		// Create a dummy data source containing no data.
		DataSet data = new DataSet();

		// Set the appropriate mail merge clean up options to remove any unused regions from the document.
		doc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_UNUSED_REGIONS);

		// Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
		// automatically as they are unused.
		doc.getMailMerge().executeWithRegions(data);

		// Save the output document to disk.
		doc.save(dataDir + "TestFile.RemoveEmptyRegions Out.doc");

	}
}