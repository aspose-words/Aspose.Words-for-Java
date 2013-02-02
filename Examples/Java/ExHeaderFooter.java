//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;


public class ExHeaderFooter extends ExBase
{
    @Test
    public void createFooter() throws Exception
    {
        //ExStart
        //ExFor:HeaderFooter
        //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
        //ExFor:HeaderFooterCollection
        //ExFor:Story.AppendParagraph
        //ExSummary:Creates a footer using the document object model and inserts it into a section.
        Document doc = new Document();

        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FOOTER_PRIMARY);
        doc.getFirstSection().getHeadersFooters().add(footer);

        // Add a paragraph with text to the footer.
        footer.appendParagraph("TEST FOOTER");

        doc.save(getMyDir() + "HeaderFooter.CreateFooter Out.doc");
        //ExEnd

        doc = new Document(getMyDir() + "HeaderFooter.CreateFooter Out.doc");
        Assert.assertTrue(doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getRange().getText().contains("TEST FOOTER"));
    }

    @Test
    public void removeFooters() throws Exception
    {
        //ExStart
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExFor:HeaderFooterType
        //ExId:RemoveFooters
        //ExSummary:Deletes all footers from all sections, but leaves headers intact.
        Document doc = new Document(getMyDir() + "HeaderFooter.RemoveFooters.doc");

        for (Section section : doc.getSections())
        {
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

        doc.save(getMyDir() + "HeaderFooter.RemoveFooters Out.doc");
        //ExEnd
    }

    @Test
    public void SetExportHeadersFootersMode() throws Exception
    {
        //ExStart
        //ExFor:ExportHeadersFootersMode
        //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
        //ExSummary:Demonstrates how to disable the export of headers and footers when saving to HTML based formats.
        Document doc = new Document(getMyDir() + "HeaderFooter.RemoveFooters.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE); // Disables exporting headers and footers.

        doc.save(getMyDir() + "HeaderFooter.DisableHeadersFooters Out.html", saveOptions);
        //ExEnd

        // Verify that the output document is correct.
        doc = new Document(getMyDir() + "HeaderFooter.DisableHeadersFooters Out.html");
        Assert.assertFalse(doc.getRange().getText().contains("DYNAMIC TEMPLATE"));
    }

    @Test
    public void replaceText() throws Exception
    {
        //ExStart
        //ExFor:Document.FirstSection
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExFor:Range.Replace(String, String, Boolean, Boolean)
        //ExSummary:Shows how to replace text in the document footer.
        // Open the template document, containing obsolete copyright information in the footer.
        Document doc = new Document(getMyDir() + "HeaderFooter.ReplaceText.doc");

        HeaderFooterCollection headersFooters = doc.getFirstSection().getHeadersFooters();
        HeaderFooter footer = headersFooters.getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
        footer.getRange().replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2011 by Aspose Pty Ltd.", false, false);

        doc.save(getMyDir() + "HeaderFooter.ReplaceText Out.doc");
        //ExEnd

        // Verify that the appropriate changes were made to the output document.
        doc = new Document(getMyDir() + "HeaderFooter.ReplaceText Out.doc");
        Assert.assertTrue(doc.getRange().getText().contains("Copyright (C) 2011 by Aspose Pty Ltd."));
    }

    @Test
    public void headerFooterPrimerCaller() throws Exception
    {
        primer();
    }

    //ExStart
    //ExId:HeaderFooterPrimer
    //ExSummary:Maybe a bit complicated example, but demonstrates many things that can be done with headers/footers.
    public void primer() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Section currentSection = builder.getCurrentSection();
        PageSetup pageSetup = currentSection.getPageSetup();

        // Specify if we want headers/footers of the first page to be different from other pages.
        // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
        // different headers/footers for odd and even pages.
        pageSetup.setDifferentFirstPageHeaderFooter(true);

        // --- Create header for the first page. ---
        pageSetup.setHeaderDistance(20);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Set font properties for header text.
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14);
        // Specify header title for the first page.
        builder.write("Aspose.Words Header/Footer Creation Primer - Title Page.");

        // --- Create header for pages other than first. ---
        pageSetup.setHeaderDistance(20);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        // Insert absolutely positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        String imageFileName = getMyDir() + "Aspose.Words.gif";
        builder.insertImage(imageFileName, RelativeHorizontalPosition.PAGE, 10, RelativeVerticalPosition.PAGE, 10, 50, 50, WrapType.THROUGH);

        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        // Specify another header title for other pages.
        builder.write("Aspose.Words Header/Footer Creation Primer.");

        // --- Create footer for pages other than first. ---
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);

        // We use table with two cells to make one part of the text on the line (with page numbering)
        // to be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder.startTable();

        // Clear table borders
        builder.getCellFormat().clearFormatting();

        builder.insertCell();
        // Set first cell to 1/3 of the page width.
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 /3));

        // Insert page numbering text here.
        // It uses PAGE and NUMPAGES fields to auto calculate current page number and total number of pages.
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");

        // Align this text to the left.
        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);

        builder.insertCell();
        // Set the second cell to 2/3 of the page width.
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 * 2 / 3));

        builder.write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

        // Align this text to the right.
        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        builder.endRow();
        builder.endTable();

        builder.moveToDocumentEnd();
        // Make page break to create a second page on which the primary headers/footers will be seen.
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Make section break to create a third page with different page orientation.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        // Get the new section and its page setup.
        currentSection = builder.getCurrentSection();
        pageSetup = currentSection.getPageSetup();

        // Set page orientation of the new section to landscape.
        pageSetup.setOrientation(Orientation.LANDSCAPE);

        // This section does not need different first page header/footer.
        // We need only one title page in the document and the header/footer for this page
        // has already been defined in the previous section
        pageSetup.setDifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section by default.
        // Call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this.
        // Page width is different for the new section and therefore we need to set
        // a different cell widths for a footer table.
        currentSection.getHeadersFooters().linkToPrevious(false);

        // If we want to use the already existing header/footer set for this section
        // but with some minor modifications then it may be expedient to copy headers/footers
        // from the previous section and apply the necessary modifications where we want them.
        copyHeadersFootersFromPreviousSection(currentSection);

        // Find the footer that we want to change.
        HeaderFooter primaryFooter = currentSection.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);

        Row row = primaryFooter.getTables().get(0).getFirstRow();
        row.getFirstCell().getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 / 3));
        row.getLastCell().getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 * 2 / 3));

        // Save the resulting document.
        doc.save(getMyDir() + "HeaderFooter.Primer Out.doc");
    }

    /**
     * Clones and copies headers/footers form the previous section to the specified section.
     */
    private static void copyHeadersFootersFromPreviousSection(Section section) throws Exception
    {
        Section previousSection = (Section)section.getPreviousSibling();

        if (previousSection == null)
            return;

        section.getHeadersFooters().clear();

        for (HeaderFooter headerFooter : previousSection.getHeadersFooters())
            section.getHeadersFooters().add(headerFooter.deepClone(true));
    }
    //ExEnd
}

