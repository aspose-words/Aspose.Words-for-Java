package com.aspose.words.examples.programming_documents.HeadersAndFooters;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.creation.BuildTableFromDataTable;

public class RemoveFootersButLeaveHeadersIntact {

    private static final String dataDir = Utils.getSharedDataDir(BuildTableFromDataTable.class) + "HeadersAndFooters/";

    public static void main(String[] args) throws Exception {
        //ExStart:RemoveFootersButLeaveHeadersIntact
        Document doc = new Document(dataDir + "HeaderFooter.RemoveFooters.doc");

        for (Section section : doc.getSections()) {
            // Up to three different footers are possible in a section (for first, even and odd pages).
            // We check and delete all of them.
            HeaderFooter footer;

            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
            if (footer != null)
                footer.remove();

            // Primary footer is the footer used for odd pages.
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            if (footer != null)
                footer.remove();

            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
            if (footer != null)
                footer.remove();
        }

        doc.save(dataDir + "HeaderFooter.RemoveFooters Out.doc");
        //ExEnd:RemoveFootersButLeaveHeadersIntact

    }

}
