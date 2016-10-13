package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.image.BufferedImage;
import java.io.File;


public class DetectDocumentSignatures
{
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DetectDocumentSignatures.class);

        // The path to the document which is to be processed.
        String filePath = dataDir + "Document.Signed.docx";

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature())
        {
            System.out.println(java.text.MessageFormat.format(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    new File(filePath).getName()));
        }
    }
}
