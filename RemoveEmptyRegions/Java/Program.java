//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package RemoveEmptyRegions;

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

        // Create a dummy data source containing two empty DataTables which corresponds to the regions in the document.
        DataSet data = new DataSet();
        DataTable suppliers = new DataTable(new CachedRowSetImpl(), "Suppliers");
        DataTable storeDetails = new DataTable(new CachedRowSetImpl(), "StoreDetails");
        data.getTables().add(suppliers);
        data.getTables().add(storeDetails);

        // Set the RemoveEmptyRegions to true in order to remove unmerged mail merge regions from the document.
        doc.getMailMerge().setRemoveEmptyRegions(true);

        // Execute mail merge. It will have no effect as there is no data.
        doc.getMailMerge().executeWithRegions(data);

        // Save the output document to disk.
        doc.save(dataDir + "TestFile.RemoveEmptyRegions Out.doc");
        //ExEnd
    }
}
