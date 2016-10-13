package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;
import com.aspose.words.*;
public class ConvertDocumentToHtmlWithRoundtrip
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToHtmlWithRoundtrip.class);

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
        System.out.println("Document converted to html with roundtrip informations successfully.");
    }
}
