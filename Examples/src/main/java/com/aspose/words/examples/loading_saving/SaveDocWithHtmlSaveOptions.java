package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HtmlMetafileFormat;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.examples.Utils;

public class SaveDocWithHtmlSaveOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SaveDocWithHtmlSaveOptions.class);

        saveHtmlWithMetafileFormat(dataDir);
        importExportSVGinHTML(dataDir);
    }

    public static void saveHtmlWithMetafileFormat(String dataDir) throws Exception
    {
        // ExStart:SaveHtmlWithMetafileFormat
        Document doc = new Document(dataDir + "Document.docx");
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF);

        dataDir = dataDir + "SaveHtmlWithMetafileFormat_out.html";
        doc.save(dataDir, options);
        // ExEnd:SaveHtmlWithMetafileFormat
        System.out.println("\nDocument saved with Metafile format.\nFile saved at " + dataDir);
    }

    public static void importExportSVGinHTML(String dataDir) throws Exception
    {
        // ExStart:ImportExportSVGinHTML
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Here is an SVG image: ");
        builder.insertHtml("<svg height='210' width='500'> <polygon points='100,10 40,198 190,78 10,78 160,198' style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' /></svg> ");
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setMetafileFormat(HtmlMetafileFormat.SVG);

        dataDir = dataDir + "ExportSVGinHTML_out.html";
        doc.save(dataDir, options);
        // ExEnd:ImportExportSVGinHTML
        System.out.println("\nDocument saved with SVG Metafile format.\nFile saved at " + dataDir);
    }
}
