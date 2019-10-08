package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooterBookmarksExportMode;
import com.aspose.words.MetafileRenderingOptions;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.examples.Utils;

public class WorkingWithPdfSaveOptions {

    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub

        String dataDir = Utils.getDataDir(WorkingWithPdfSaveOptions.class);

        EscapeUriInPdf(dataDir);
        ExportHeaderFooterBookmarks(dataDir);
        ScaleWmfFontsToMetafileSize(dataDir);
        AdditionalTextPositioning(dataDir);
    }

    public static void EscapeUriInPdf(String dataDir) throws Exception {
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

    public static void ExportHeaderFooterBookmarks(String dataDir) throws Exception {
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

    public static void ScaleWmfFontsToMetafileSize(String dataDir) throws Exception {
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

    public static void AdditionalTextPositioning(String dataDir) throws Exception {
        // ExStart:AdditionalTextPositioning
        // The path to the documents directory.
        Document doc = new Document(dataDir + "TestFile.docx");

        PdfSaveOptions options = new PdfSaveOptions();
        options.setAdditionalTextPositioning(true);

        dataDir = dataDir + "AdditionalTextPositioning_out.pdf";
        doc.save(dataDir, options);
        // ExEnd:AdditionalTextPositioning
        System.out.println("\nFile saved at " + dataDir);
    }

}
