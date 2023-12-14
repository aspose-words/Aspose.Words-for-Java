package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.BreakType;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.Section;
import com.aspose.words.PageSetup;
import com.aspose.words.Orientation;
import com.aspose.words.HeaderFooter;


class WorkingWithHeadersAndFooters extends DocsExamplesBase
{
    @Test
    public void createHeaderFooter() throws Exception
    {
        //ExStart:CreateHeaderFooter
        //GistId:84cab3a22008f041ee6c1e959da09949
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use HeaderPrimary and FooterPrimary
        // if you want to set header/footer for all document.
        // This header/footer type also responsible for odd pages.
        //ExStart:HeaderFooterType
        //GistId:84cab3a22008f041ee6c1e959da09949
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for page.");
        //ExEnd:HeaderFooterType

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Footer for page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.CreateHeaderFooter.docx");
        //ExEnd:CreateHeaderFooter
    }

    @Test
    public void differentFirstPage() throws Exception
    {
        //ExStart:DifferentFirstPage
        //GistId:84cab3a22008f041ee6c1e959da09949
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify that we want different headers and footers for first page.
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.write("Header for the first page.");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_FIRST);
        builder.write("Footer for the first page.");

        builder.moveToSection(0);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.DifferentFirstPage.docx");
        //ExEnd:DifferentFirstPage
    }

    @Test
    public void oddEvenPages() throws Exception
    {
        //ExStart:OddEvenPages
        //GistId:84cab3a22008f041ee6c1e959da09949
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // Specify that we want different headers and footers for even and odd pages.            
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.write("Header for even pages.");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header for odd pages.");            
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_EVEN);
        builder.write("Footer for even pages.");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Footer for odd pages.");

        builder.moveToSection(0);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.OddEvenPages.docx");
        //ExEnd:OddEvenPages
    }

    @Test
    public void insertImage() throws Exception
    {
        //ExStart:InsertImage
        //GistId:84cab3a22008f041ee6c1e959da09949
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);            
        builder.insertImage(getImagesDir() + "Logo.jpg", RelativeHorizontalPosition.RIGHT_MARGIN, 10.0,
            RelativeVerticalPosition.PAGE, 10.0, 50.0, 50.0, WrapType.THROUGH);            

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.InsertImage.docx");
        //ExEnd:InsertImage
    }

    @Test
    public void fontProps() throws Exception
    {
        //ExStart:FontProps
        //GistId:84cab3a22008f041ee6c1e959da09949
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14.0);
        builder.write("Header for page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.FontProps.docx");
        //ExEnd:FontProps
    }

    @Test
    public void pageNumbers() throws Exception
    {
        //ExStart:PageNumbers
        //GistId:84cab3a22008f041ee6c1e959da09949
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
    public void linkToPreviousHeaderFooter() throws Exception
    {
        //ExStart:LinkToPreviousHeaderFooter
        //GistId:84cab3a22008f041ee6c1e959da09949
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getPageSetup().setDifferentFirstPageHeaderFooter(true);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14.0);
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
        builder.getFont().setSize(12.0);
        builder.write("New Header for the first page.");

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.LinkToPreviousHeaderFooter.docx");
        //ExEnd:LinkToPreviousHeaderFooter
    }

    @Test
    public void sectionsWithDifferentHeaders() throws Exception
    {
        //ExStart:SectionsWithDifferentHeaders            
        //GistId:1afca4d3da7cb4240fb91c3d93d8c30d            
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getCurrentSection().getPageSetup();
        pageSetup.setDifferentFirstPageHeaderFooter(true);
        pageSetup.setHeaderDistance(20.0);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14.0);
        builder.write("Header for the first page.");

        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        // Insert a positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        builder.insertImage(getImagesDir() + "Logo.jpg", RelativeHorizontalPosition.PAGE, 10.0,
            RelativeVerticalPosition.PAGE, 10.0, 50.0, 50.0, WrapType.THROUGH);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Header for odd page.");            

        doc.save(getArtifactsDir() + "WorkingWithHeadersAndFooters.SectionsWithDifferentHeaders.docx");
        //ExEnd:SectionsWithDifferentHeaders
    }

    //ExStart:CopyHeadersFootersFromPreviousSection
    //GistId:84cab3a22008f041ee6c1e959da09949
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
