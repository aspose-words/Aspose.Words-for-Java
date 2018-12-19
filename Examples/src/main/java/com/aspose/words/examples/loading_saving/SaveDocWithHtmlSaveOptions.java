package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class SaveDocWithHtmlSaveOptions {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SaveDocWithHtmlSaveOptions.class);

        saveHtmlWithMetafileFormat(dataDir);
        importExportSVGinHTML(dataDir);
        setCssClassNamePrefix(dataDir);
        setExportCidUrlsForMhtmlResources(dataDir);
        setResolveFontNames(dataDir);
    }

    public static void saveHtmlWithMetafileFormat(String dataDir) throws Exception {
        // ExStart:SaveHtmlWithMetafileFormat
        Document doc = new Document(dataDir + "Document.docx");
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF);

        dataDir = dataDir + "SaveHtmlWithMetafileFormat_out.html";
        doc.save(dataDir, options);
        // ExEnd:SaveHtmlWithMetafileFormat
        System.out.println("\nDocument saved with Metafile format.\nFile saved at " + dataDir);
    }

    public static void importExportSVGinHTML(String dataDir) throws Exception {
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

    public static void setCssClassNamePrefix(String dataDir) throws Exception {
        // ExStart:SetCssClassNamePrefix
        Document doc = new Document(dataDir + "Document.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setCssClassNamePrefix("pfx_");

        dataDir = dataDir + "CssClassNamePrefix_out.html";
        doc.save(dataDir, saveOptions);
        // ExEnd:SetCssClassNamePrefix
        System.out.println("\nDocument saved with CSS prefix pfx_.\nFile saved at " + dataDir);
    }

    public static void setExportCidUrlsForMhtmlResources(String dataDir) throws Exception {
        // ExStart:SetExportCidUrlsForMhtmlResources
        Document doc = new Document(dataDir + "CidUrls.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        saveOptions.setPrettyFormat(true);
        saveOptions.setExportCidUrlsForMhtmlResources(true);

        dataDir = dataDir + "SetExportCidUrlsForMhtmlResources_out.mhtml";
        doc.save(dataDir, saveOptions);
        // ExEnd:SetExportCidUrlsForMhtmlResources
        System.out.println("\nDocument has saved with Content - Id URL scheme.\nFile saved at " + dataDir);
    }

    public static void setResolveFontNames(String dataDir) throws Exception {
        // ExStart:SetResolveFontNames
        Document doc = new Document(dataDir + "Test File (docx).docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        saveOptions.setPrettyFormat(true);
        saveOptions.setResolveFontNames(true);

        dataDir = dataDir + "ResolveFontNames_out.html";
        doc.save(dataDir, saveOptions);
        // ExEnd:SetResolveFontNames
        System.out.println("\nFontSettings is used to resolve font family name.\nFile saved at " + dataDir);
    }
}
