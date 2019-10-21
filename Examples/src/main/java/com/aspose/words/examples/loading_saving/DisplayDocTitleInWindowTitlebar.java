package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 7/12/2017.
 */
public class DisplayDocTitleInWindowTitlebar {
    public static void main(String[] args) throws Exception {
        //ExStart:DisplayDocTitleInWindowTitlebar
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DisplayDocTitleInWindowTitlebar.class);

        // Load the document.
        Document doc = new Document(dataDir + "Test File (doc).doc");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.setDisplayDocTitle(true);

        // Save the document in PDF format.
        doc.save(dataDir + "Test File.Pdf", saveOptions);
        //ExEnd:DisplayDocTitleInWindowTitlebar

        System.out.println("Document converted to PDF successfully.");
    }
}
