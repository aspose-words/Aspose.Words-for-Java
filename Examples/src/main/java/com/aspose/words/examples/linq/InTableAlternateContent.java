package com.aspose.words.examples.linq;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.examples.Utils;

public class InTableAlternateContent {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.

        String dataDir = Utils.getDataDir(InTableAlternateContent.class);

        String fileName = "InTableAlternateContent.doc";
        // Load the template document.
        Document doc = new Document(dataDir + fileName);

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Execute the build report.
        engine.buildReport(doc, Common.GetContracts(), "ds");

        dataDir = dataDir + Utils.GetOutputFilePath(fileName);

        // Save the finished document to disk.
        doc.save(dataDir);
//   System.out.println("\nIn-Table list template document is populated with the data about managers.\nFile saved at " + dataDir);

    }
}
