package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Section;
import com.aspose.words.PageSetup;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.PreferredWidth;
import com.aspose.words.BreakType;
import com.aspose.words.Orientation;
import com.aspose.words.HeaderFooter;
import com.aspose.words.Row;

@Test
public class WorkingWithHeadersAndFooters extends DocsExamplesBase
{
    @Test
    public void createHeaderFooter() throws Exception
    {
        //ExStart:CreateHeaderFooterUsingDocBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Section currentSection = builder.getCurrentSection();
        PageSetup pageSetup = currentSection.getPageSetup();
        // Specify if we want headers/footers of the first page to be different from other pages.
        // You can also use PageSetup.OddAndEvenPagesHeaderFooter property to specify
        // different headers/footers for odd and even pages.
        pageSetup.setDifferentFirstPageHeaderFooter(true);
        pageSetup.setHeaderDistance(20.0);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14.0);

        builder.write("Aspose.Words Header/Footer Creation Primer - Title Page.");

        pageSetup.setHeaderDistance(20.0);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        // Insert a positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        builder.insertImage(getImagesDir() + "Graphics Interchange Format.gif", RelativeHorizontalPosition.PAGE, 10.0,
            RelativeVerticalPosition.PAGE, 10.0, 50.0, 50.0, WrapType.THROUGH);

        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        builder.write("Aspose.Words Header/Footer Creation Primer.");

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);

        // We use a table with two cells to make one part of the text on the line (with page numbering).
        // To be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder.startTable();

        builder.getCellFormat().clearFormatting();

        builder.insertCell();

        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 / 3));

        // It uses PAGE and NUMPAGES fields to auto calculate the current page number and many pages.
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");

        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);

        builder.insertCell();

        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 * 2 / 3));

        builder.write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        builder.endRow();
        builder.endTable();

        builder.moveToDocumentEnd();

        // Make a page break to create a second page on which the primary headers/footers will be seen.
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        currentSection = builder.getCurrentSection();
        pageSetup = currentSection.getPageSetup();
        pageSetup.setOrientation(Orientation.LANDSCAPE);
        // This section does not need a different first-page header/footer we need only one title page in the document,
        // and the header/footer for this page has already been defined in the previous section.
        pageSetup.setDifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section
        // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
        // is different for the new section, and therefore we need to set different cell widths for a footer table.
        currentSection.getHeadersFooters().linkToPrevious(false);

        // If we want to use the already existing header/footer set for this section.
        // But with some minor modifications, then it may be expedient to copy headers/footers
        // from the previous section and apply the necessary modifications where we want them.
        copyHeadersFootersFromPreviousSection(currentSection);

        HeaderFooter primaryFooter = currentSection.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);

        Row row = primaryFooter.getTables().get(0).getFirstRow();
        row.getFirstCell().getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 / 3));
        row.getLastCell().getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100 * 2 / 3));

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
        //ExEnd:CreateHeaderFooterUsingDocBuilder
    }

    //ExStart:CopyHeadersFootersFromPreviousSection
    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    private void copyHeadersFootersFromPreviousSection(Section section)
    {
        Section previousSection = (Section)section.getPreviousSibling();

        if (previousSection == null)
            return;

        section.getHeadersFooters().clear();

        for (HeaderFooter headerFooter : (Iterable<HeaderFooter>) previousSection.getHeadersFooters())
            section.getHeadersFooters().add(headerFooter.deepClone(true));
    }
    //ExEnd:CopyHeadersFootersFromPreviousSection        
}
