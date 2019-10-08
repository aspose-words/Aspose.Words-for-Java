package com.aspose.words.examples.loading_saving;

import com.aspose.words.DocSaveOptions;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.SaveOptions;
import com.aspose.words.examples.Utils;

public class WorkingWithDocSaveOptions {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(WorkingWithDocSaveOptions.class);

        EncryptDocumentWithPassword(dataDir);
        AlwaysCompressMetafiles(dataDir);
        SavePictureBullet(dataDir);
	}

    public static void EncryptDocumentWithPassword(String dataDir) throws Exception {
        //ExStart: EncryptDocumentWithPassword
        Document doc = new Document(dataDir + "Document.doc");
        DocSaveOptions docSaveOptions = new DocSaveOptions();
        docSaveOptions.setPassword("password");
        dataDir = dataDir + "Document.Password_out.doc";
        doc.save(dataDir, docSaveOptions);
        //ExEnd: EncryptDocumentWithPassword
        System.out.println("\nThe password of document is set using RC4 encryption method. \nFile saved at " + dataDir);
    }

    public static void AlwaysCompressMetafiles(String dataDir) throws Exception {
        //ExStart: AlwaysCompressMetafiles
        Document doc = new Document(dataDir + "Document.doc");
        DocSaveOptions saveOptions = new DocSaveOptions();

        saveOptions.setAlwaysCompressMetafiles(false);
        doc.save("SmallMetafilesUncompressed.doc", saveOptions);
        //ExEnd: AlwaysCompressMetafiles
        System.out.println("\nThe document is saved with AlwaysCompressMetafiles setting to false. \nFile saved at " + dataDir);
    }
    
    public static void SavePictureBullet(String dataDir) throws Exception
    {
        //ExStart:SavePictureBullet
        Document doc = new Document(dataDir + "in.doc");
        DocSaveOptions saveOptions = (DocSaveOptions)SaveOptions.createSaveOptions(SaveFormat.DOC);
        saveOptions.setSavePictureBullet(false);
        doc.save(dataDir + "out.doc", saveOptions);
        //ExEnd:SavePictureBullet
        System.out.println("\nThe document is saved with SavePictureBullet setting to false. \nFile saved at " + dataDir);
    }
}
