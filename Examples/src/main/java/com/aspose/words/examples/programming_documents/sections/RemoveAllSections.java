/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class RemoveAllSections
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:RemoveAllSections
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document doc = new Document(dataDir + "Document.doc");
        doc.getSections().clear();
        // ExEnd:RemoveAllSections
        System.out.println("All sections removed successfully form the document.");
    }
}