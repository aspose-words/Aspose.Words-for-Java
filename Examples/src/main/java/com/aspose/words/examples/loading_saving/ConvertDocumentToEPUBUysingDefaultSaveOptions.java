package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

import java.nio.charset.Charset;

public class ConvertDocumentToEPUBUysingDefaultSaveOptions {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ConvertDocumentToEPUBUysingDefaultSaveOptions.class);
        // Open an existing document from disk.
        Document doc = new Document(dataDir + "Document.EpubConversion.doc");

        // Create a new instance of HtmlSaveOptions. This object allows us to set options that control
        // how the output document is saved.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Specify the desired encoding.
        saveOptions.setEncoding(Charset.forName("UTF-8"));

        // Specify at what elements to split the internal HTML at. This creates a new HTML within the EPUB
        // which allows you to limit the size of each HTML part. This is useful for readers which cannot read
        // HTML files greater than a certain size e.g 300kb.
        saveOptions.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);

        // Specify that we want to export document properties.
        saveOptions.setExportDocumentProperties(true);

        // Specify that we want to save in EPUB format.
        saveOptions.setSaveFormat(SaveFormat.EPUB);

        // Export the document as an EPUB file.
        doc.save(dataDir + "Document.EpubConversion_out_.epub", saveOptions);
        System.out.println("Document using save options converted to EPUB successfully.");
    }
}
