package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.TxtSaveOptions;
import com.aspose.words.examples.Utils;

public class WorkingWithTxt {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(WorkingWithTxt.class);

        saveAsTxt(dataDir);
        addBidiMarks(dataDir);
    }

    public static void saveAsTxt(String dataDir) throws Exception {
        // ExStart:SaveAsTxt
        Document doc = new Document(dataDir + "Document.doc");
        dataDir = dataDir + "Document.ConvertToTxt_out.txt";
        doc.save(dataDir);
        // ExEnd:SaveAsTxt
        System.out.println("\nDocument saved as TXT.\nFile saved at " + dataDir);
    }

    public static void addBidiMarks(String dataDir) throws Exception {
        // ExStart:AddBidiMarks
        Document doc = new Document(dataDir + "Input.docx");
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.setAddBidiMarks(false);

        dataDir = dataDir + "Document.AddBidiMarks_out.txt";
        doc.save(dataDir, saveOptions);
        // ExEnd:AddBidiMarks
        System.out.println("\nAdd bi-directional marks set successfully.\nFile saved at " + dataDir);
    }
}
