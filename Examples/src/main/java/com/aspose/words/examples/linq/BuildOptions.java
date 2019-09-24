package com.aspose.words.examples.linq;

import com.aspose.words.Document;
import com.aspose.words.ReportingEngine;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class BuildOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(BuildOptions.class);

        removeEmptyParagraphs(dataDir);
    }

    public static void removeEmptyParagraphs(String dataDir) throws Exception {
        //ExStart:RemoveEmptyParagraphs
        // Load the template document.
        Document doc = new Document(dataDir + "template_cleanup.docx");

        // Create a Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        //engine.setOptions(ReportBuildOptions.REMOVE_EMPTY_PARAGRAPHS);
        engine.buildReport(doc, Common.GetClients());

        dataDir = dataDir + "output.docx";
        doc.save(dataDir, SaveFormat.DOCX);
        //ExEnd:RemoveEmptyParagraphs
        System.out.println("\nEmpty paragraphs are removed from the document successfully.\nFile saved at " + dataDir);
    }
}
