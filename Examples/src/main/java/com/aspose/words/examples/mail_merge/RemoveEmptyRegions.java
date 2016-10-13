package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataSet;


public class RemoveEmptyRegions
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemoveEmptyRegions.class);

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
        doc.save(dataDir + "Output.doc");

        assert doc.getMailMerge().getFieldNames().length == 0: "Error: There are still unused regions remaining in the document";

        System.out.println("Non empty regions removed during mail merge successfully.");
    }
}