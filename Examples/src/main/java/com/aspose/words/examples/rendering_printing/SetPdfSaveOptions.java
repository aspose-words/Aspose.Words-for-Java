package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class SetPdfSaveOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SetPdfSaveOptions.class);

        escapeUriInPdf(dataDir);
    }

    public static void escapeUriInPdf(String dataDir) throws Exception
    {
        // ExStart:EscapeUriInPdf
        // The path to the documents directory.
        Document doc = new Document(dataDir + "EscapeUri.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setEscapeUri(false);

        dataDir = dataDir + "EscapeUri_out.pdf";
        doc.save(dataDir, options);
        // ExEnd:EscapeUriInPdf
        System.out.println("\nFile saved at " + dataDir);
    }
}
