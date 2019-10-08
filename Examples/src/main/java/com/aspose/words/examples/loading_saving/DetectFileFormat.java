package com.aspose.words.examples.loading_saving;

import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.examples.Utils;


public class DetectFileFormat {
    public static void main(String[] args) throws Exception {
        //ExStart:DetectFileFormat
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DetectDocumentSignatures.class);

        // The path to the document which is to be processed.
        String filePath = dataDir + "Document.Signed.docx";

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        System.out.println("The document format is: " + FileFormatUtil.loadFormatToExtension(info.getLoadFormat()));
        System.out.println("Document is encrypted: " + info.isEncrypted());
        System.out.println("Document has a digital signature: " + info.hasDigitalSignature());
        //ExEnd:DetectFileFormat
    }
}
