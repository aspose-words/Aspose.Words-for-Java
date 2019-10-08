package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.examples.Utils;

import java.io.File;


public class DigitallySignedPdf {
    public static void main(String[] args) throws Exception {
        //ExStart:DigitallySignedPdf
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DigitallySignedPdf.class);

        // The path to the document which is to be processed.
        String filePath = dataDir + "Document.Signed.docx";
        Document doc = new Document();

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature()) {
            System.out.println(java.text.MessageFormat.format(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    new File(filePath).getName()));
        }
        //ExEnd:DigitallySignedPdf
    }
}
