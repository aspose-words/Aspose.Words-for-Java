package com.aspose.words.examples.featurescomparison.headerfooter.removeheaderfooter;

import com.aspose.words.Document;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Section;
import com.aspose.words.examples.Utils;

public class AsposeHeaderFooterRemove 
{
    public static void main(String[] args) throws Exception 
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeHeaderFooterRemove.class);

        Document doc = new Document(dataDir + "AsposeHeaderFooter.docx");

        for (Section section : doc.getSections())
        {
            // Up to three different header footers are possible in a section (for first, even and odd pages).
            // We check and delete all of them.
            HeaderFooter header;
            HeaderFooter footer;

            header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST);
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
            if (header != null)
                header.remove();
            if (footer != null)
                footer.remove();

            // Primary header and footer is used for odd pages.
            header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            if (header != null)
                header.remove();
            if (footer != null)
                footer.remove();

            header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN);
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
            if (header != null)
                header.remove();
            if (footer != null)
                footer.remove();
        }

        doc.save(dataDir + "AsposeHeaderFooterRemoved.docx");
        System.out.println("Done.");
    }
}
