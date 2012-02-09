//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package RemoveEmptyRegions;

import com.aspose.words.MailMergeCleanupOptions;
import com.sun.rowset.CachedRowSetImpl;
import java.io.File;
import java.net.URI;

import com.aspose.words.DataSet;
import com.aspose.words.DataTable;
import com.aspose.words.Document;

class Program
{
    public static void main(String[] args) throws Exception
    {
        // Sample infrastructure.
        URI exeDir = Program.class.getResource("").toURI();
        String dataDir = new File(exeDir.resolve("../../Data")) + File.separator;

        //ExStart
        //ExFor:MailMerge.RemoveEmptyRegions
        //ExId:RemoveEmptyRegions
        //ExSummary:Shows how to remove unmerged mail merge regions from the document.
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
        //ExEnd

        assert doc.getMailMerge().getFieldNames().length == 0: "Error: There are still unused regions remaining in the document";
    }
}
