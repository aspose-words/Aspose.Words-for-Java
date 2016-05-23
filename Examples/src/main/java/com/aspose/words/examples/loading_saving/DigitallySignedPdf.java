/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.examples.Utils;

import java.io.File;


public class DigitallySignedPdf
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:DetectDocumentSignatures
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DigitallySignedPdf.class);

        // The path to the document which is to be processed.
        String filePath = dataDir + "Document.Signed.docx";
        Document doc = new Document();

        FileFormatInfo info = FileFormatUtil.detectFileFormat(filePath);
        if (info.hasDigitalSignature())
        {
            System.out.println(java.text.MessageFormat.format(
                    "Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.",
                    new File(filePath).getName()));
        }
        // ExEnd:DetectDocumentSignatures
    }


}
//ExEnd