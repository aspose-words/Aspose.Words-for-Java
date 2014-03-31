/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package loadingandsaving.pagesplitter.java;

import java.io.File;
import java.text.MessageFormat;

import com.aspose.words.*;

public class PageSplitter
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/loadingandsaving/pagesplitter/data/";

        SplitAllDocumentsToPages(dataDir);
    }

    public static void SplitDocumentToPages(File docName) throws Exception
    {
        String folderName = docName.getParent();
        String fileName = docName.getName();
        String extensionName = fileName.substring(fileName.lastIndexOf("."));
        String outFolder = new File(folderName, "Out").getAbsolutePath();

        System.out.println("Processing document: " + fileName + extensionName);

        Document doc = new Document(docName.getAbsolutePath());

        // Create and attach collector to the document before page layout is built.
        LayoutCollector layoutCollector = new LayoutCollector(doc);

        // This will build layout model and collect necessary information.
        doc.updatePageLayout();

        // Split nodes in the document into separate pages.
        DocumentPageSplitter splitter = new DocumentPageSplitter(layoutCollector);

        // Save each page to the disk as a separate document.
        for (int page = 1; page <= doc.getPageCount(); page++)
        {
            Document pageDoc = splitter.GetDocumentOfPage(page);
            pageDoc.save(new File(outFolder, MessageFormat.format("{0} - page{1} Out{2}", fileName, page, extensionName)).getAbsolutePath());
        }

        // Detach the collector from the document.
        layoutCollector.setDocument(null);
    }

    public static void SplitAllDocumentsToPages(String folderName) throws Exception
    {
        File[] files = new File(folderName).listFiles();

        for (File file : files) {
            if (file.isFile()) {
                SplitDocumentToPages(file);
            }
        }
    }
}