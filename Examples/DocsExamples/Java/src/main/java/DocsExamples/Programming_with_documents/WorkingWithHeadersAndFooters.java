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
    public void CreateHeaderFooter() throws Exception {
        //ExStart:CreateHeaderFooter
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //ExStart:HeaderFooterType
        //GistId:58431f54e34e5597f8cbaf97481d5321
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page.");
        //ExEnd:HeaderFooterType

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Header for odd page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
        //ExEnd:CreateHeaderFooter
    }

    @Test
    public void DifferentFirstPage() throws Exception {
        //ExStart:DifferentFirstPage
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want different headers and footers for first page.
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page.");

        builder.moveToSection(0);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.DifferentFirstPage.docx");
        //ExEnd:DifferentFirstPage
    }

    @Test
    public void OddEvenPages() throws Exception {
        //ExStart:OddEvenPages
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want different headers and footers for even and odd pages.
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header for even pages.");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for odd pages.");

        builder.moveToSection(0);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.OddEvenPages.docx");
        //ExEnd:OddEvenPages
    }

    @Test
    public void InsertImage() throws Exception {
        //ExStart:InsertImage
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.insertImage(getImagesDir() + "Logo.jpg", RelativeHorizontalPosition.RIGHT_MARGIN, 10,
                RelativeVerticalPosition.PAGE, 10, 50, 50, WrapType.THROUGH);

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.InsertImage.docx");
        //ExEnd:InsertImage
    }

    @Test
    public void FontProps() throws Exception {
        //ExStart:FontProps
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14);
        builder.write("Header for odd page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.HeaderFooterFontProps.docx");
        //ExEnd:FontProps
    }

    @Test
    public void PageNumbers() throws Exception {
        //ExStart:PageNumbers
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.PageNumbers.docx");
        //ExEnd:PageNumbers
    }

    @Test
    public void LinkToPreviousHeaderFooter() throws Exception {
        //ExStart:LinkToPreviousHeaderFooter
        //GistId:58431f54e34e5597f8cbaf97481d5321
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14);
        builder.write("Header for the first page.");

        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        Section currentSection = builder.getCurrentSection();
        PageSetup pageSetup = currentSection.getPageSetup();
        pageSetup.setOrientation(Orientation.LANDSCAPE);
        // This section does not need a different first-page header/footer we need only one title page in the document,
        // and the header/footer for this page has already been defined in the previous section.
        pageSetup.setDifferentFirstPageHeaderFooter(false);

        // This section displays headers/footers from the previous section
        // by default call currentSection.HeadersFooters.LinkToPrevious(false) to cancel this page width
        // is different for the new section.
        currentSection.getHeadersFooters().linkToPrevious(false);
        currentSection.getHeadersFooters().clear();

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        builder.getFont().setName("Arial");
        builder.getFont().setSize(12);
        builder.write("New Header for the first page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.LinkToPreviousHeaderFooter.docx");
        //ExEnd:LinkToPreviousHeaderFooter
    }

    @Test
    public void SectionsWithDifferentHeaders() throws Exception {
        //ExStart:SectionsWithDifferentHeaders
        //GistId:7c0668453e53ed7a57d3ea3a05520f21
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getCurrentSection().getPageSetup();
        pageSetup.setDifferentFirstPageHeaderFooter(true);
        pageSetup.setHeaderDistance(20);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14);
        builder.write("Header for the first page.");

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        // Insert a positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        builder.insertImage(getImagesDir() + "Logo.jpg", RelativeHorizontalPosition.PAGE, 10,
                RelativeVerticalPosition.PAGE, 10, 50, 50, WrapType.THROUGH);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Header for odd page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.SectionsWithDifferentHeaders.docx");
        //ExEnd:SectionsWithDifferentHeaders
    }

    //ExStart:CopyHeadersFootersFromPreviousSection
    //GistId:58431f54e34e5597f8cbaf97481d5321
    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    private void copyHeadersFootersFromPreviousSection(Section section)
    {
        Section previousSection = (Section)section.getPreviousSibling();

        if (previousSection == null)
            return;

        section.getHeadersFooters().clear();

        for (HeaderFooter headerFooter : previousSection.getHeadersFooters())
            section.getHeadersFooters().add(headerFooter.deepClone(true));
    }
    //ExEnd:CopyHeadersFootersFromPreviousSection
}
