package com.aspose.words.examples.linq;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.examples.Utils;

public class SingleRow {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {
        //ExStart:SingleRow
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SingleRow.class);

        String fileName = "SingleRow.doc";
        // Load the template document.
        Document doc = new Document(dataDir + fileName);

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Execute the build report.
        engine.buildReport(doc, Common.GetManager(), "manager");

        dataDir = dataDir + Utils.GetOutputFilePath(fileName);

        // Save the finished document to disk.
        doc.save(dataDir);
        //ExEnd:SingleRow

        System.out.println("\nSingle row template document is populated with the data about manager.\nFile saved at " + dataDir);

    }
}
