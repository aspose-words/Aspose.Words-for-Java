package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.words.examples.Utils;

public class SaveOptionsHtmlFixed {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(SaveOptionsHtmlFixed.class) + "LoadingSavingAndConverting/";

        UseFontFromTargetMachine(dataDir);
        writeAllCSSrulesinSingleFile(dataDir);
    }

    static void UseFontFromTargetMachine(String dataDir) throws Exception {
        // ExStart:UseFontFromTargetMachine
        // Load the document from disk.
        Document doc = new Document(dataDir + "Test File (doc).doc");

        HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
        options.setUseTargetMachineFonts(true);

        dataDir = dataDir + "Test File_out.html";

        // Save the document to disk.
        doc.save(dataDir, options);
        // ExEnd:UseFontFromTargetMachine
        System.out.println("\nHTML file created successfully.\nFile saved at " + dataDir);
    }

    static void writeAllCSSrulesinSingleFile(String dataDir) throws Exception {
        // ExStart:WriteAllCSSrulesinSingleFile
        // Load the document from disk.
        Document doc = new Document(dataDir + "Test File (doc).doc");

        HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
        //Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
        //Default value is false.
        //All CSS rules are written into single file "styles.css
        options.setSaveFontFaceCssSeparately(false);

        dataDir = dataDir + "WriteAllCSSrulesinSingleFile_out.html";
        // Save the document to disk.
        doc.save(dataDir, options);
        // ExEnd:WriteAllCSSrulesinSingleFile
        System.out.println("\nWrite all CSS rules in single file successfully.\nFile saved at " + dataDir);
    }
}