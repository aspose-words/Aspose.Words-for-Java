/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.sections;

import com.aspose.words.Document;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class CopySection
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:CopySection
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SectionsAccessByIndex.class);
        Document srcDoc = new Document(dataDir + "Document.doc");
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section)dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);
        dataDir = dataDir + "Document.Copy_out_.doc";
        dstDoc.save(dataDir);
        // ExEnd:CopySection
        System.out.println("\nSection copied successfully.\nFile saved at " + dataDir);
    }
}