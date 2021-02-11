package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.examples.Utils;

public class DocumentBuilderHeadersAndFooters {
    public static void main(String[] args) throws Exception {
        //ExStart:DocumentBuilderHeaderAndFooters
        String dataDir = Utils.getDataDir(DocumentBuilderHeadersAndFooters.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want headers and footers different for first, even and odd pages.
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        // Create the headers.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header for even pages");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for all other pages");

        // Create two pages in the document.
        builder.moveToSection(0);
        builder.writeln("Page1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page2");

        doc.save(dataDir + "DocumentBuilder.HeadersAndFooters.docx");
        //ExEnd:DocumentBuilderHeaderAndFooters
    }
}