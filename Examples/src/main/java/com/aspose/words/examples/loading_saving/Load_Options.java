package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class Load_Options {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Load_Options.class);

        loadOptionsUpdateDirtyFields(dataDir);
        loadAndSaveEncryptedODT(dataDir);
        verifyODTdocument(dataDir);
    }

    public static void loadOptionsUpdateDirtyFields(String dataDir) throws Exception
    {
        // ExStart:LoadOptionsUpdateDirtyFields
        LoadOptions lo = new LoadOptions();
        //Update the fields with the dirty attribute
        lo.setUpdateDirtyFields(true);
        //Load the Word document
        Document doc = new Document(dataDir + "input.docx", lo);
        //Save the document into DOCX
        dataDir = dataDir + "output.docx";
        doc.save(dataDir, SaveFormat.DOCX);
        // ExEnd:LoadOptionsUpdateDirtyFields
        System.out.println("\nUpdate the fields with the dirty attribute successfully.\nFile saved at " + dataDir);
    }

    public static void loadAndSaveEncryptedODT(String dataDir) throws Exception
    {
        // ExStart:LoadAndSaveEncryptedODT
        Document doc = new Document(dataDir + "encrypted.odt", new com.aspose.words.LoadOptions("password"));
        doc.save(dataDir + "out.odt", new OdtSaveOptions("newpassword"));
        // ExEnd:LoadAndSaveEncryptedODT
        System.out.println("\nLoad and save encrypted document successfully.\nFile saved at " + dataDir);
    }

    public static void verifyODTdocument(String dataDir) throws Exception
    {
        // ExStart:VerifyODTdocument
        FileFormatInfo info = FileFormatUtil.detectFileFormat(dataDir + "encrypted.odt");
        System.out.println(info.isEncrypted());
        // ExEnd:VerifyODTdocument
    }
}
