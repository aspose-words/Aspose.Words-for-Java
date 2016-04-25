/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;
import com.aspose.words.*;
public class ConvertDocumentToHtmlWithRoundtrip
{
    public static void main(String[] args) throws Exception
    {
        // ExStart:ConvertDocumentToHtmlWithRoundtrip
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToByte.class);

        // Load the document.
        Document doc = new Document(dataDir + "Test File (doc).doc");

        HtmlSaveOptions options = new HtmlSaveOptions();

        //HtmlSaveOptions.ExportRoundtripInformation property specifies
        //whether to write the roundtrip information when saving to HTML, MHTML or EPUB.
        //Default value is true for HTML and false for MHTML and EPUB.
        options.setExportRoundtripInformation(true);
        doc.save(dataDir + "ExportRoundtripInformation_out_.html", options);

        doc = new Document(dataDir + "ExportRoundtripInformation_out_.html");

        //Save the document Docx file format
        doc.save(dataDir + "TestFile_out_.docx", SaveFormat.DOCX);
        // ExEnd:ConvertDocumentToHtmlWithRoundtrip
        System.out.println("Document converted to html with roundtrip informations successfully.");
    }
}
