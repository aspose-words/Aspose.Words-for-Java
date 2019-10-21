/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.DigitalSignature;
import com.aspose.words.Document;
import com.aspose.words.examples.Utils;
//FIXME: no input file

public class AccessAndVerifySignature {
    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(AccessAndVerifySignature.class) + "LoadingSavingAndConverting/";
        // The path to the document which is to be processed.
        String filePath = dataDir + "Document.Signed.docx";
        Document doc = new Document(filePath);

        for (DigitalSignature signature : doc.getDigitalSignatures()) {
            System.out.println("*** Signature Found ***");
            System.out.println("Is valid: " + signature.isValid());
            System.out.println("Reason for signing: " + signature.getComments()); // This property is available in MS Word documents only.
            System.out.println("Time of signing: " + signature.getSignTime());
        }
        //ExEnd:1
    }
}
