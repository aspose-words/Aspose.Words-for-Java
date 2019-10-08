package com.aspose.words.examples.linq;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.examples.Utils;

public class SettingBackgroundColor {
    /**
     * The main entry point for the application.
     */
    public static void main(String[] args) throws Exception {
        //ExStart:SettingBackgroundColor
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SettingBackgroundColor.class);

        String fileName = "SettingBackgroundColor.docx";
        // Load the template document.
        Document doc = new Document(dataDir + fileName);

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        // Execute the build report.
        engine.buildReport(doc, new Object());

        dataDir = dataDir + "SettingBackgroundColor_out.docx";
        // Save the finished document to disk.
        doc.save(dataDir);
        //ExEnd:SettingBackgroundColor

        System.out.println("\nSet the background color of text and shape successfully.\nFile saved at " + dataDir);
    }
}
