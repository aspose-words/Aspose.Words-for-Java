
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;


public class DocumentBuilderHeadersAndFooters {
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DocumentBuilderHeadersAndFooters.class);

        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header First");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header Even");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header Odd");

        builder.moveToSection(0);
        builder.writeln("Page1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page3");

        doc.save(dataDir + "output.doc");

    }
}