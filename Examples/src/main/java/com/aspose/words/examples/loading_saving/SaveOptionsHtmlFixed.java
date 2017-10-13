package com.aspose.words.examples.loading_saving;
import com.aspose.words.Document;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.examples.Utils;
/**
 * Created by Home on 10/13/2017.
 */
public class SaveOptionsHtmlFixed {
    public static void main(String[] args) throws Exception {
        UseFontFromTargetMachine();
    }

    static void UseFontFromTargetMachine() throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(SaveOptionsHtmlFixed.class) + "LoadingSavingAndConverting/";
        // The path to the document which is to be processed.
        String filePath = dataDir + "Test File (doc).doc";
        // ExStart:UseFontFromTargetMachine
        // Load the document from disk.
        Document doc = new Document(filePath);

        HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
        options.setUseTargetMachineFonts(true);

        dataDir = dataDir + "Test File_out.html";

        // Save the document to disk.
        doc.save(dataDir, options);
        // ExEnd:UseFontFromTargetMachine
        System.out.println("\nHTML file created successfully.\nFile saved at " + dataDir);
    }
}
