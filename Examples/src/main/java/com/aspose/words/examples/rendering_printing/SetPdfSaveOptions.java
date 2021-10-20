package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class SetPdfSaveOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SetPdfSaveOptions.class);

        escapeUriInPdf(dataDir);
        exportHeaderFooterBookmarks(dataDir);
        scaleWmfFontsToMetafileSize(dataDir);
    }

    public static void escapeUriInPdf(String dataDir) throws Exception {
        // ExStart:EscapeUriInPdf
        // The path to the documents directory.
        Document doc = new Document(dataDir + "EscapeUri.docx");

        doc.save(dataDir + "EscapeUri_out.pdf");
        // ExEnd:EscapeUriInPdf
        System.out.println("\nFile saved at " + dataDir);
    }

    public static void exportHeaderFooterBookmarks(String dataDir) throws Exception {
        // ExStart:ExportHeaderFooterBookmarks
        // The path to the documents directory.
        Document doc = new Document(dataDir + "TestFile.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.getOutlineOptions().setDefaultBookmarksOutlineLevel(1);
        options.setHeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode.FIRST);

        dataDir = dataDir + "ExportHeaderFooterBookmarks_out.pdf";
        doc.save(dataDir, options);
        // ExEnd:ExportHeaderFooterBookmarks
        System.out.println("\nFile saved at " + dataDir);
    }

    public static void scaleWmfFontsToMetafileSize(String dataDir) throws Exception {
        // ExStart:ScaleWmfFontsToMetafileSize
        // The path to the documents directory.
        Document doc = new Document(dataDir + "MetafileRendering.docx");

        MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();
        metafileRenderingOptions.setScaleWmfFontsToMetafileSize(false);

        // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap.
        PdfSaveOptions options = new PdfSaveOptions();
        options.setMetafileRenderingOptions(metafileRenderingOptions);

        dataDir = dataDir + "ScaleWmfFontsToMetafileSize_out.pdf";
        doc.save(dataDir, options);
        // ExEnd:ScaleWmfFontsToMetafileSize
        System.out.println("\nFonts as metafile are rendered to its default size in PDF. File saved at " + dataDir);
    }
}
