/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.splitintohtmlpages.java;

import java.io.File;
import java.net.URI;


public class SplitIntoHtmlPages
{
    public static void main(String[] args) throws Exception
    {
        // You need to have a valid license for Aspose.Words.
        // The best way is to embed the license as a resource into the project
        // and specify only file name without path in the following call.
        // Aspose.Words.License license = new Aspose.Words.License();
        // license.SetLicense(@"Aspose.Words.lic");


        String dataDir = "src/loadingandsaving/loadingandsavinghtml/splitintohtmlpages/data/";

        String srcFileName = dataDir + "SOI 2007-2012-DeeM with footnote added.doc";
        String tocTemplate = dataDir + "TocTemplate.doc";

        File outDir = new File(dataDir, "Out");
        outDir.mkdirs();

        // This class does the job.
        Worker w = new Worker();
        w.execute(srcFileName, tocTemplate, outDir.getPath());

        System.out.println("Success.");
    }
}