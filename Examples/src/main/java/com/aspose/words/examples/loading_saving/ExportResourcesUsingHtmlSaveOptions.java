/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.loading_saving;

import com.aspose.words.CssStyleSheetType;
import com.aspose.words.Document;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.examples.Utils;
//FIXME: no input file

public class ExportResourcesUsingHtmlSaveOptions {

    // The path to the documents directory.
    private static final String dataDir = Utils.getDataDir(ExportResourcesUsingHtmlSaveOptions.class);

    public static void main(String[] args) throws Exception {
        //ExStart:ExportResourcesUsingHtmlSaveOptions
        // The path to the document which is to be processed.
        String filePath = dataDir + "Document.doc";
        Document doc = new Document(filePath);
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder(dataDir + "\\Resources");
        doc.save(dataDir + "ExportResourcesUsingHtmlSaveOptions_out.html", saveOptions);
        //ExEnd:ExportResourcesUsingHtmlSaveOptions
    }
}