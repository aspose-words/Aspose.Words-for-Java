package com.aspose.words.examples.linq;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.examples.Utils;

public class CommonMasterDetail {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {
        //ExStart:CommonMasterDetail
        // The path to the documents directory.

        String dataDir = Utils.getDataDir(CommonMasterDetail.class);

        String fileName = "CommonMasterDetail.doc";
        // Load the template document.
        Document doc = new Document(dataDir + fileName);

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Execute the build report.
        // engine.getKnownTypes().add(DateUtil.class);
        engine.buildReport(doc, Common.GetContracts(), "ds");

        dataDir = dataDir + Utils.GetOutputFilePath(fileName);

        // Save the finished document to disk.
        doc.save(dataDir);
        //ExEnd:CommonMasterDetail

        System.out.println("\nCommon master detail template document is populated with the data about managers and it's contracts.\nFile saved at " + dataDir);

    }
}
