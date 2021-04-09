// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.Paragraph;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.Section;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.ExportHeadersFootersMode;
import com.aspose.words.HeaderFooterCollection;
import com.aspose.words.FindReplaceOptions;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.PageSetup;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.PreferredWidth;
import com.aspose.words.Orientation;
import com.aspose.words.Row;
import org.testng.annotations.DataProvider;


@Test
public class ExHeaderFooter extends ApiExampleBase
{
    @Test
    public void create() throws Exception
    {
        //ExStart
        //ExFor:HeaderFooter
        //ExFor:HeaderFooter.#ctor(DocumentBase, HeaderFooterType)
        //ExFor:HeaderFooter.HeaderFooterType
        //ExFor:HeaderFooter.IsHeader
        //ExFor:HeaderFooterCollection
        //ExFor:Paragraph.IsEndOfHeaderFooter
        //ExFor:Paragraph.ParentSection
        //ExFor:Paragraph.ParentStory
        //ExFor:Story.AppendParagraph
        //ExSummary:Shows how to create a header and a footer.
        Document doc = new Document();
        
        // Create a header and append a paragraph to it. The text in that paragraph
        // will appear at the top of every page of this section, above the main body text.
        HeaderFooter header = new HeaderFooter(doc, HeaderFooterType.HEADER_PRIMARY);
        doc.getFirstSection().getHeadersFooters().add(header);

        Paragraph para = header.appendParagraph("My header.");

        Assert.assertTrue(header.isHeader());
        Assert.assertTrue(para.isEndOfHeaderFooter());

        // Create a footer and append a paragraph to it. The text in that paragraph
        // will appear at the bottom of every page of this section, below the main body text.
        HeaderFooter footer = new HeaderFooter(doc, HeaderFooterType.FOOTER_PRIMARY);
        doc.getFirstSection().getHeadersFooters().add(footer);

        para = footer.appendParagraph("My footer.");

        Assert.assertFalse(footer.isHeader());
        Assert.assertTrue(para.isEndOfHeaderFooter());

        Assert.assertEquals(footer, para.getParentStory());
        Assert.assertEquals(footer.getParentSection(), para.getParentSection());
        Assert.assertEquals(footer.getParentSection(), header.getParentSection());

        doc.save(getArtifactsDir() + "HeaderFooter.Create.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "HeaderFooter.Create.docx");

        Assert.assertTrue(doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getRange().getText()
            .contains("My header."));
        Assert.assertTrue(doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getRange().getText()
            .contains("My footer."));
    }

    @Test
    public void link() throws Exception
    {
        //ExStart
        //ExFor:HeaderFooter.IsLinkedToPrevious
        //ExFor:HeaderFooterCollection.Item(System.Int32)
        //ExFor:HeaderFooterCollection.LinkToPrevious(Aspose.Words.HeaderFooterType,System.Boolean)
        //ExFor:HeaderFooterCollection.LinkToPrevious(System.Boolean)
        //ExFor:HeaderFooter.ParentSection
        //ExSummary:Shows how to link headers and footers between sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 3");

        // Move to the first section and create a header and a footer. By default,
        // the header and the footer will only appear on pages in the section that contains them.
        builder.moveToSection(0);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("This is the header, which will be displayed in sections 1 and 2.");

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("This is the footer, which will be displayed in sections 1, 2 and 3.");

        // We can link a section's headers/footers to the previous section's headers/footers
        // to allow the linking section to display the linked section's headers/footers.
        doc.getSections().get(1).getHeadersFooters().linkToPrevious(true);

        // Each section will still have its own header/footer objects. When we link sections,
        // the linking section will display the linked section's header/footers while keeping its own.
        Assert.assertNotEquals(doc.getSections().get(0).getHeadersFooters().get(0), doc.getSections().get(1).getHeadersFooters().get(0));
        Assert.assertNotEquals(doc.getSections().get(0).getHeadersFooters().get(0).getParentSection(), doc.getSections().get(1).getHeadersFooters().get(0).getParentSection());

        // Link the headers/footers of the third section to the headers/footers of the second section.
        // The second section already links to the first section's header/footers,
        // so linking to the second section will create a link chain.
        // The first, second, and now the third sections will all display the first section's headers.
        doc.getSections().get(2).getHeadersFooters().linkToPrevious(true);

        // We can un-link a previous section's header/footers by passing "false" when calling the LinkToPrevious method.
        doc.getSections().get(2).getHeadersFooters().linkToPrevious(false);

        // We can also select only a specific type of header/footer to link using this method.
        // The third section now will have the same footer as the second and first sections, but not the header.
        doc.getSections().get(2).getHeadersFooters().linkToPrevious(HeaderFooterType.FOOTER_PRIMARY, true);

        // The first section's header/footers cannot link themselves to anything because there is no previous section.
        Assert.assertEquals(2, doc.getSections().get(0).getHeadersFooters().getCount());
        Assert.AreEqual(2, doc.getSections().get(0).getHeadersFooters().Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious));
        
        // All the second section's header/footers are linked to the first section's headers/footers.
        Assert.assertEquals(6, doc.getSections().get(1).getHeadersFooters().getCount());
        Assert.AreEqual(6, doc.getSections().get(1).getHeadersFooters().Count(hf => ((HeaderFooter)hf).IsLinkedToPrevious));

        // In the third section, only the footer is linked to the first section's footer via the second section.
        Assert.assertEquals(6, doc.getSections().get(2).getHeadersFooters().getCount());
        Assert.AreEqual(5, doc.getSections().get(2).getHeadersFooters().Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious));
        Assert.assertTrue(doc.getSections().get(2).getHeadersFooters().get(3).isLinkedToPrevious());

        doc.save(getArtifactsDir() + "HeaderFooter.Link.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "HeaderFooter.Link.docx");

        Assert.assertEquals(2, doc.getSections().get(0).getHeadersFooters().getCount());
        Assert.AreEqual(2, doc.getSections().get(0).getHeadersFooters().Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious));

        Assert.assertEquals(0, doc.getSections().get(1).getHeadersFooters().getCount());
        Assert.AreEqual(0, doc.getSections().get(1).getHeadersFooters().Count(hf => ((HeaderFooter)hf).IsLinkedToPrevious));

        Assert.assertEquals(5, doc.getSections().get(2).getHeadersFooters().getCount());
        Assert.AreEqual(5, doc.getSections().get(2).getHeadersFooters().Count(hf => !((HeaderFooter)hf).IsLinkedToPrevious));
    }

    @Test
    public void removeFooters() throws Exception
    {
        //ExStart
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExSummary:Shows how to delete all footers from a document.
        Document doc = new Document(getMyDir() + "Header and footer types.docx");

        // Iterate through each section and remove footers of every kind.
        for (Section section : doc.<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            // There are three kinds of footer and header types.
            // 1 -  The "First" header/footer, which only appears on the first page of a section.
            HeaderFooter footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
            footer?.Remove();

            // 2 -  The "Primary" header/footer, which appears on odd pages.
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            footer?.Remove();

            // 3 -  The "Even" header/footer, which appears on odd even pages. 
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
            footer?.Remove();

            Assert.AreEqual(0, section.getHeadersFooters().Count(hf => !((HeaderFooter)hf).IsHeader));
        }

        doc.save(getArtifactsDir() + "HeaderFooter.RemoveFooters.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "HeaderFooter.RemoveFooters.docx");

        Assert.assertEquals(1, doc.getSections().getCount());
        Assert.AreEqual(0, doc.getFirstSection().getHeadersFooters().Count(hf => !((HeaderFooter)hf).IsHeader));
        Assert.AreEqual(3, doc.getFirstSection().getHeadersFooters().Count(hf => ((HeaderFooter)hf).IsHeader));
    }

    @Test
    public void exportMode() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportHeadersFootersMode
        //ExFor:ExportHeadersFootersMode
        //ExSummary:Shows how to omit headers/footers when saving a document to HTML.
        Document doc = new Document(getMyDir() + "Header and footer types.docx");

        // This document contains headers and footers. We can access them via the "HeadersFooters" collection.
        Assert.assertEquals("First header", doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST).getText().trim());

        // Formats such as .html do not split the document into pages, so headers/footers will not function the same way
        // they would when we open the document as a .docx using Microsoft Word.
        // If we convert a document with headers/footers to html, the conversion will assimilate the headers/footers into body text.
        // We can use a SaveOptions object to omit headers/footers while converting to html.
        HtmlSaveOptions saveOptions =
            new HtmlSaveOptions(SaveFormat.HTML); { saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE); }

        doc.save(getArtifactsDir() + "HeaderFooter.ExportMode.html", saveOptions);

        // Open our saved document and verify that it does not contain the header's text
        doc = new Document(getArtifactsDir() + "HeaderFooter.ExportMode.html");

        Assert.assertFalse(doc.getRange().getText().contains("First header"));
        //ExEnd
    }

    @Test
    public void replaceText() throws Exception
    {
        //ExStart
        //ExFor:Document.FirstSection
        //ExFor:Section.HeadersFooters
        //ExFor:HeaderFooterCollection.Item(HeaderFooterType)
        //ExFor:HeaderFooter
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace text in a document's footer.
        Document doc = new Document(getMyDir() + "Footer.docx");

        HeaderFooterCollection headersFooters = doc.getFirstSection().getHeadersFooters();
        HeaderFooter footer = headersFooters.getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);

        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setMatchCase(false);
            options.setFindWholeWordsOnly(false);
        }

        int currentYear = new Date().getYear();
        footer.getRange().replace("(C) 2006 Aspose Pty Ltd.", $"Copyright (C) {currentYear} by Aspose Pty Ltd.", options);

        doc.save(getArtifactsDir() + "HeaderFooter.ReplaceText.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "HeaderFooter.ReplaceText.docx");

        Assert.assertTrue(doc.getRange().getText().contains($"Copyright (C) {currentYear} by Aspose Pty Ltd."));
    }

    //ExStart
    //ExFor:IReplacingCallback
    //ExFor:PageSetup.DifferentFirstPageHeaderFooter
    //ExSummary:Shows how to track the order in which a text replacement operation traverses nodes.
    @Test (dataProvider = "orderDataProvider") //ExSkip
    public void order(boolean differentFirstPageHeaderFooter) throws Exception
    {
        Document doc = new Document(getMyDir() + "Header and footer types.docx");

        Section firstPageSection = doc.getFirstSection();

        ReplaceLog logger = new ReplaceLog();
        FindReplaceOptions options = new FindReplaceOptions(); { options.setReplacingCallback(logger); }
        
        // Using a different header/footer for the first page will affect the search order.
        firstPageSection.getPageSetup().setDifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter);
        doc.getRange().replaceInternal(new Regex("(header|footer)"), "", options);

        if (differentFirstPageHeaderFooter)
            Assert.AreEqual("First header\nFirst footer\nSecond header\nSecond footer\nThird header\nThird footer\n", 
                logger.Text.Replace("\r", ""));
        else
            Assert.AreEqual("Third header\nFirst header\nThird footer\nFirst footer\nSecond header\nSecond footer\n", 
                logger.Text.Replace("\r", ""));
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "orderDataProvider")
	public static Object[][] orderDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    /// <summary>
    /// During a find-and-replace operation, records the contents of every node that has text that the operation 'finds',
    /// in the state it is in before the replacement takes place.
    /// This will display the order in which the text replacement operation traverses nodes.
    /// </summary>
    private static class ReplaceLog implements IReplacingCallback
    {
        public /*ReplaceAction*/int replacing(ReplacingArgs args)
        {
            mTextBuilder.AppendLine(args.getMatchNode().getText());
            return ReplaceAction.SKIP;
        }

         String Text => private mTextBuilder.ToStringmTextBuilder();

        private /*final*/ StringBuilder mTextBuilder = new StringBuilder();
    }
    //ExEnd

    @Test
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

        // Create header for the first page.
        pageSetup.setHeaderDistance(20.0);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        builder.getFont().setName("Arial");
        builder.getFont().setBold(true);
        builder.getFont().setSize(14.0);
        builder.write("Aspose.Words Header/Footer Creation Primer - Title Page.");

        // Create header for pages other than first.
        pageSetup.setHeaderDistance(20.0);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        // Insert an absolutely positioned image into the top/left corner of the header.
        // Distance from the top/left edges of the page is set to 10 points.
        String imageFileName = getImageDir() + "Logo.jpg";
        builder.insertImage(imageFileName, RelativeHorizontalPosition.PAGE, 10.0, RelativeVerticalPosition.PAGE, 10.0,
            50.0, 50.0, WrapType.THROUGH);

        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Aspose.Words Header/Footer Creation Primer.");

        // Create footer for pages other than first.
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);

        // We use a table with two cells to make one part of the text on the line (with page numbering)
        // to be aligned left, and the other part of the text (with copyright) to be aligned right.
        builder.startTable();

        builder.getCellFormat().clearFormatting();

        builder.insertCell();

        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100.0F / 3f));

        // Insert page numbering text here.
        // It uses PAGE and NUMPAGES fields to auto calculate the current page number and a total number of pages.
        builder.write("Page ");
        builder.insertField("PAGE", "");
        builder.write(" of ");
        builder.insertField("NUMPAGES", "");

        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.LEFT);

        builder.insertCell();
        builder.getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100.0F * 2f / 3f));

        builder.write("(C) 2001 Aspose Pty Ltd. All rights reserved.");

        builder.getCurrentParagraph().getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        builder.endRow();
        builder.endTable();

        builder.moveToDocumentEnd();
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Make section break to create a third page with a different page orientation.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        currentSection = builder.getCurrentSection();
        pageSetup = currentSection.getPageSetup();

        pageSetup.setOrientation(Orientation.LANDSCAPE);

        // This section does not need different first page header/footer.
        // We need only one title page in the document and the header/footer for this page
        // has already been defined in the previous section.
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

        HeaderFooter primaryFooter = currentSection.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);

        Row row = primaryFooter.getTables().get(0).getFirstRow();
        row.getFirstCell().getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100.0F / 3f));
        row.getLastCell().getCellFormat().setPreferredWidth(PreferredWidth.fromPercent(100.0F * 2f / 3f));

        doc.save(getArtifactsDir() + "HeaderFooter.Primer.docx");
    }

    /// <summary>
    /// Clones and copies headers/footers form the previous section to the specified section.
    /// </summary>
    private static void copyHeadersFootersFromPreviousSection(Section section)
    {
        Section previousSection = (Section)section.getPreviousSibling();

        if (previousSection == null)
            return;

        section.getHeadersFooters().clear();

        for (HeaderFooter headerFooter : previousSection.getHeadersFooters().<HeaderFooter>OfType() !!Autoporter error: Undefined expression type )
        {
            section.getHeadersFooters().add(headerFooter.deepClone(true));
        }
    }
}
