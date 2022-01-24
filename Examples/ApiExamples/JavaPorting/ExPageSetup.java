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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Orientation;
import com.aspose.words.PageVerticalAlignment;
import com.aspose.words.BreakType;
import org.testng.Assert;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.SectionLayoutMode;
import com.aspose.words.Paragraph;
import com.aspose.words.SectionStart;
import com.aspose.words.TextColumnCollection;
import com.aspose.ms.System.Drawing.Printing.PrinterSettings;
import com.aspose.words.Section;
import com.aspose.words.PaperSize;
import com.aspose.words.ConvertUtil;
import com.aspose.words.PageSetup;
import com.aspose.words.TextColumn;
import com.aspose.words.LineNumberRestartMode;
import com.aspose.words.PageBorderDistanceFrom;
import com.aspose.words.PageBorderAppliesTo;
import com.aspose.words.Border;
import com.aspose.words.BorderType;
import com.aspose.words.LineStyle;
import java.awt.Color;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.NumberStyle;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.FootnoteType;
import com.aspose.words.FootnoteOptions;
import com.aspose.words.FootnotePosition;
import com.aspose.words.FootnoteNumberingRule;
import com.aspose.words.EndnoteOptions;
import com.aspose.words.EndnotePosition;
import com.aspose.words.MultiplePagesType;
import com.aspose.words.TextOrientation;
import com.aspose.words.Body;
import org.testng.annotations.DataProvider;


@Test
public class ExPageSetup extends ApiExampleBase
{
    @Test
    public void clearFormatting() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.PageSetup
        //ExFor:DocumentBuilder.InsertBreak
        //ExFor:DocumentBuilder.Document
        //ExFor:PageSetup
        //ExFor:PageSetup.Orientation
        //ExFor:PageSetup.VerticalAlignment
        //ExFor:PageSetup.ClearFormatting
        //ExFor:Orientation
        //ExFor:PageVerticalAlignment
        //ExFor:BreakType
        //ExSummary:Shows how to apply and revert page setup settings to sections in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Modify the page setup properties for the builder's current section and add text.
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setVerticalAlignment(PageVerticalAlignment.CENTER);
        builder.writeln("This is the first section, which landscape oriented with vertically centered text.");

        // If we start a new section using a document builder,
        // it will inherit the builder's current page setup properties.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        Assert.assertEquals(Orientation.LANDSCAPE, doc.getSections().get(1).getPageSetup().getOrientation());
        Assert.assertEquals(PageVerticalAlignment.CENTER, doc.getSections().get(1).getPageSetup().getVerticalAlignment());

        // We can revert its page setup properties to their default values using the "ClearFormatting" method.
        builder.getPageSetup().clearFormatting();

        Assert.assertEquals(Orientation.PORTRAIT, doc.getSections().get(1).getPageSetup().getOrientation());
        Assert.assertEquals(PageVerticalAlignment.TOP, doc.getSections().get(1).getPageSetup().getVerticalAlignment());

        builder.writeln("This is the second section, which is in default Letter paper size, portrait orientation and top alignment.");

        doc.save(getArtifactsDir() + "PageSetup.ClearFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.ClearFormatting.docx");

        Assert.assertEquals(Orientation.LANDSCAPE, doc.getSections().get(0).getPageSetup().getOrientation());
        Assert.assertEquals(PageVerticalAlignment.CENTER, doc.getSections().get(0).getPageSetup().getVerticalAlignment());

        Assert.assertEquals(Orientation.PORTRAIT, doc.getSections().get(1).getPageSetup().getOrientation());
        Assert.assertEquals(PageVerticalAlignment.TOP, doc.getSections().get(1).getPageSetup().getVerticalAlignment());
    }

    @Test (dataProvider = "differentFirstPageHeaderFooterDataProvider")
    public void differentFirstPageHeaderFooter(boolean differentFirstPageHeaderFooter) throws Exception
    {
        //ExStart
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExSummary:Shows how to enable or disable primary headers/footers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two types of header/footers.
        // 1 -  The "First" header/footer, which appears on the first page of the section.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.writeln("First page header.");

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_FIRST);
        builder.writeln("First page footer.");

        // 2 -  The "Primary" header/footer, which appears on every page in the section.
        // We can override the primary header/footer by a first and an even page header/footer. 
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Primary header.");

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.writeln("Primary footer.");

        builder.moveToSection(0);
        builder.writeln("Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3.");

        // Each section has a "PageSetup" object that specifies page appearance-related properties
        // such as orientation, size, and borders.
        // Set the "DifferentFirstPageHeaderFooter" property to "true" to apply the first header/footer to the first page.
        // Set the "DifferentFirstPageHeaderFooter" property to "false"
        // to make the first page display the primary header/footer.
        builder.getPageSetup().setDifferentFirstPageHeaderFooter(differentFirstPageHeaderFooter);

        doc.save(getArtifactsDir() + "PageSetup.DifferentFirstPageHeaderFooter.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.DifferentFirstPageHeaderFooter.docx");

        Assert.assertEquals(differentFirstPageHeaderFooter, doc.getFirstSection().getPageSetup().getDifferentFirstPageHeaderFooter());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "differentFirstPageHeaderFooterDataProvider")
	public static Object[][] differentFirstPageHeaderFooterDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "oddAndEvenPagesHeaderFooterDataProvider")
    public void oddAndEvenPagesHeaderFooter(boolean oddAndEvenPagesHeaderFooter) throws Exception
    {
        //ExStart
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExSummary:Shows how to enable or disable even page headers/footers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Below are two types of header/footers.
        // 1 -  The "Primary" header/footer, which appears on every page in the section.
        // We can override the primary header/footer by a first and an even page header/footer. 
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Primary header.");

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.writeln("Primary footer.");

        // 2 -  The "Even" header/footer, which appears on every even page of this section.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.writeln("Even page header.");

        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_EVEN);
        builder.writeln("Even page footer.");

        builder.moveToSection(0);
        builder.writeln("Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3.");

        // Each section has a "PageSetup" object that specifies page appearance-related properties
        // such as orientation, size, and borders.
        // Set the "OddAndEvenPagesHeaderFooter" property to "true"
        // to display the even page header/footer on even pages.
        // Set the "OddAndEvenPagesHeaderFooter" property to "false"
        // to display the primary header/footer on even pages.
        builder.getPageSetup().setOddAndEvenPagesHeaderFooter(oddAndEvenPagesHeaderFooter);

        doc.save(getArtifactsDir() + "PageSetup.OddAndEvenPagesHeaderFooter.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.OddAndEvenPagesHeaderFooter.docx");

        Assert.assertEquals(oddAndEvenPagesHeaderFooter, doc.getFirstSection().getPageSetup().getOddAndEvenPagesHeaderFooter());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "oddAndEvenPagesHeaderFooterDataProvider")
	public static Object[][] oddAndEvenPagesHeaderFooterDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void charactersPerLine() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.CharactersPerLine
        //ExFor:PageSetup.LayoutMode
        //ExFor:SectionLayoutMode
        //ExSummary:Shows how to specify a for the number of characters that each line may have.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // Enable pitching, and then use it to set the number of characters per line in this section.
        builder.getPageSetup().setLayoutMode(SectionLayoutMode.GRID);
        builder.getPageSetup().setCharactersPerLine(10);

        // The number of characters also depends on the size of the font.
        doc.getStyles().get("Normal").getFont().setSize(20.0);

        Assert.assertEquals(8, doc.getFirstSection().getPageSetup().getCharactersPerLine());

        builder.writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        doc.save(getArtifactsDir() + "PageSetup.CharactersPerLine.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.CharactersPerLine.docx");

        Assert.assertEquals(SectionLayoutMode.GRID, doc.getFirstSection().getPageSetup().getLayoutMode());
        Assert.assertEquals(8, doc.getFirstSection().getPageSetup().getCharactersPerLine());
    }

    @Test
    public void linesPerPage() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.LinesPerPage
        //ExFor:PageSetup.LayoutMode
        //ExFor:ParagraphFormat.SnapToGrid
        //ExFor:SectionLayoutMode
        //ExSummary:Shows how to specify a limit for the number of lines that each page may have.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        // Enable pitching, and then use it to set the number of lines per page in this section.
        // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
        builder.getPageSetup().setLayoutMode(SectionLayoutMode.LINE_GRID);
        builder.getPageSetup().setLinesPerPage(15);

        builder.getParagraphFormat().setSnapToGrid(true);

        for (int i = 0; i < 30; i++)
            builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

        doc.save(getArtifactsDir() + "PageSetup.LinesPerPage.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.LinesPerPage.docx");

        Assert.assertEquals(SectionLayoutMode.LINE_GRID, doc.getFirstSection().getPageSetup().getLayoutMode());
        Assert.assertEquals(15, doc.getFirstSection().getPageSetup().getLinesPerPage());

        for (Paragraph paragraph : (Iterable<Paragraph>) doc.getFirstSection().getBody().getParagraphs())
            Assert.assertTrue(paragraph.getParagraphFormat().getSnapToGrid());
    }

    @Test
    public void setSectionStart() throws Exception
    {
        //ExStart
        //ExFor:SectionStart
        //ExFor:PageSetup.SectionStart
        //ExFor:Document.Sections
        //ExSummary:Shows how to specify how a new section separates itself from the previous.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("This text is in section 1.");

        // Section break types determine how a new section separates itself from the previous section.
        // Below are five types of section breaks.
        // 1 -  Starts the next section on a new page:
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.writeln("This text is in section 2.");

        Assert.assertEquals(SectionStart.NEW_PAGE, doc.getSections().get(1).getPageSetup().getSectionStart());

        // 2 -  Starts the next section on the current page:
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.writeln("This text is in section 3.");

        Assert.assertEquals(SectionStart.CONTINUOUS, doc.getSections().get(2).getPageSetup().getSectionStart());

        // 3 -  Starts the next section on a new even page:
        builder.insertBreak(BreakType.SECTION_BREAK_EVEN_PAGE);
        builder.writeln("This text is in section 4.");

        Assert.assertEquals(SectionStart.EVEN_PAGE, doc.getSections().get(3).getPageSetup().getSectionStart());

        // 4 -  Starts the next section on a new odd page:
        builder.insertBreak(BreakType.SECTION_BREAK_ODD_PAGE);
        builder.writeln("This text is in section 5.");

        Assert.assertEquals(SectionStart.ODD_PAGE, doc.getSections().get(4).getPageSetup().getSectionStart());

        // 5 -  Starts the next section on a new column:
        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        columns.setCount(2);

        builder.insertBreak(BreakType.SECTION_BREAK_NEW_COLUMN);
        builder.writeln("This text is in section 6.");

        Assert.assertEquals(SectionStart.NEW_COLUMN, doc.getSections().get(5).getPageSetup().getSectionStart());

        doc.save(getArtifactsDir() + "PageSetup.SetSectionStart.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.SetSectionStart.docx");

        Assert.assertEquals(SectionStart.NEW_PAGE, doc.getSections().get(0).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.NEW_PAGE, doc.getSections().get(1).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.CONTINUOUS, doc.getSections().get(2).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.EVEN_PAGE, doc.getSections().get(3).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.ODD_PAGE, doc.getSections().get(4).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.NEW_COLUMN, doc.getSections().get(5).getPageSetup().getSectionStart());
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed")
    public void defaultPaperTray() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.FirstPageTray
        //ExFor:PageSetup.OtherPagesTray
        //ExSummary:Shows how to get all the sections in a document to use the default paper tray of the selected printer.
        Document doc = new Document();

        // Find the default printer that we will use for printing this document.
        // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
        PrinterSettings settings = new PrinterSettings();
        
        // The paper tray value stored in documents is printer specific.
        // This means the code below resets all page tray values to use the current printers default tray.
        // You can enumerate PrinterSettings.PaperSources to find the other valid paper tray values of the selected printer.
        for (Section section : doc.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            section.getPageSetup().setFirstPageTray(settings.getDefaultPageSettings().PaperSource.RawKind);
            section.getPageSetup().setOtherPagesTray(settings.getDefaultPageSettings().PaperSource.RawKind);
        }
        //ExEnd
        
        for (Section section : DocumentHelper.saveOpen(doc).getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            Assert.assertEquals(settings.getDefaultPageSettings().PaperSource.RawKind, section.getPageSetup().getFirstPageTray());
            Assert.assertEquals(settings.getDefaultPageSettings().PaperSource.RawKind, section.getPageSetup().getOtherPagesTray());
        }
    }

    @Test (enabled = false, description = "Run only when the printer driver is installed")
    public void paperTrayForDifferentPaperType() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.FirstPageTray
        //ExFor:PageSetup.OtherPagesTray
        //ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
        Document doc = new Document();

        // Find the default printer that we will use for printing this document.
        // You can define a specific printer using the "PrinterName" property of the PrinterSettings object.
        PrinterSettings settings = new PrinterSettings();

        // This is the tray we will use for pages in the "A4" paper size.
        int printerTrayForA4 = settings.getPaperSources().get(0).RawKind;

        // This is the tray we will use for pages in the "Letter" paper size.
        int printerTrayForLetter = settings.getPaperSources().get(1).RawKind;

        // Modify the PageSettings object of this section to get Microsoft Word to instruct the printer
        // to use one of the trays we identified above, depending on this section's paper size.
        for (Section section : doc.getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            if (section.getPageSetup().getPaperSize() == com.aspose.words.PaperSize.LETTER)
            {
                section.getPageSetup().setFirstPageTray(printerTrayForLetter);
                section.getPageSetup().setOtherPagesTray(printerTrayForLetter);
            }
            else if (section.getPageSetup().getPaperSize() == com.aspose.words.PaperSize.A4)
            {
                section.getPageSetup().setFirstPageTray(printerTrayForA4);
                section.getPageSetup().setOtherPagesTray(printerTrayForA4);
            }
        }
        //ExEnd

        for (Section section : DocumentHelper.saveOpen(doc).getSections().<Section>OfType() !!Autoporter error: Undefined expression type )
        {
            if (section.getPageSetup().getPaperSize() == com.aspose.words.PaperSize.LETTER)
            {
                Assert.assertEquals(printerTrayForLetter, section.getPageSetup().getFirstPageTray());
                Assert.assertEquals(printerTrayForLetter, section.getPageSetup().getOtherPagesTray());
            }
            else if (section.getPageSetup().getPaperSize() == com.aspose.words.PaperSize.A4)
            {
                Assert.assertEquals(printerTrayForA4, section.getPageSetup().getFirstPageTray());
                Assert.assertEquals(printerTrayForA4, section.getPageSetup().getOtherPagesTray());
            }
        }
    }

    @Test
    public void pageMargins() throws Exception
    {
        //ExStart
        //ExFor:ConvertUtil
        //ExFor:ConvertUtil.InchToPoint
        //ExFor:PaperSize
        //ExFor:PageSetup.PaperSize
        //ExFor:PageSetup.Orientation
        //ExFor:PageSetup.TopMargin
        //ExFor:PageSetup.BottomMargin
        //ExFor:PageSetup.LeftMargin
        //ExFor:PageSetup.RightMargin
        //ExFor:PageSetup.HeaderDistance
        //ExFor:PageSetup.FooterDistance
        //ExSummary:Shows how to adjust paper size, orientation, margins, along with other settings for a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getPageSetup().setPaperSize(PaperSize.LEGAL);
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setTopMargin(ConvertUtil.inchToPoint(1.0));
        builder.getPageSetup().setBottomMargin(ConvertUtil.inchToPoint(1.0));
        builder.getPageSetup().setLeftMargin(ConvertUtil.inchToPoint(1.5));
        builder.getPageSetup().setRightMargin(ConvertUtil.inchToPoint(1.5));
        builder.getPageSetup().setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        builder.getPageSetup().setFooterDistance(ConvertUtil.inchToPoint(0.2));

        builder.writeln("Hello world!");

        doc.save(getArtifactsDir() + "PageSetup.PageMargins.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageMargins.docx");

        Assert.assertEquals(PaperSize.LEGAL, doc.getFirstSection().getPageSetup().getPaperSize());
        Assert.assertEquals(1008.0d, doc.getFirstSection().getPageSetup().getPageWidth());
        Assert.assertEquals(612.0d, doc.getFirstSection().getPageSetup().getPageHeight());
        Assert.assertEquals(Orientation.LANDSCAPE, doc.getFirstSection().getPageSetup().getOrientation());
        Assert.assertEquals(72.0d, doc.getFirstSection().getPageSetup().getTopMargin());
        Assert.assertEquals(72.0d, doc.getFirstSection().getPageSetup().getBottomMargin());
        Assert.assertEquals(108.0d, doc.getFirstSection().getPageSetup().getLeftMargin());
        Assert.assertEquals(108.0d, doc.getFirstSection().getPageSetup().getRightMargin());
        Assert.assertEquals(14.4d, doc.getFirstSection().getPageSetup().getHeaderDistance());
        Assert.assertEquals(14.4d, doc.getFirstSection().getPageSetup().getFooterDistance());
    }

    @Test
    public void paperSizes() throws Exception
    {
        //ExStart
        //ExFor:PaperSize
        //ExFor:PageSetup.PaperSize
        //ExSummary:Shows how to set page sizes.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We can change the current page's size to a pre-defined size
        // by using the "PaperSize" property of this section's PageSetup object.
        builder.getPageSetup().setPaperSize(PaperSize.TABLOID);

        Assert.assertEquals(792.0d, builder.getPageSetup().getPageWidth());
        Assert.assertEquals(1224.0d, builder.getPageSetup().getPageHeight());

        builder.writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

        // Each section has its own PageSetup object. When we use a document builder to make a new section,
        // that section's PageSetup object inherits all the previous section's PageSetup object's values.
        builder.insertBreak(BreakType.SECTION_BREAK_EVEN_PAGE);

        Assert.assertEquals(PaperSize.TABLOID, builder.getPageSetup().getPaperSize());

        builder.getPageSetup().setPaperSize(PaperSize.A5);
        builder.writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

        Assert.assertEquals(419.55d, builder.getPageSetup().getPageWidth());
        Assert.assertEquals(595.30d, builder.getPageSetup().getPageHeight());

        builder.insertBreak(BreakType.SECTION_BREAK_EVEN_PAGE);

        // Set a custom size for this section's pages.
        builder.getPageSetup().setPageWidth(620.0);
        builder.getPageSetup().setPageHeight(480.0);

        Assert.assertEquals(PaperSize.CUSTOM, builder.getPageSetup().getPaperSize());

        builder.writeln($"This page is {builder.PageSetup.PageWidth}x{builder.PageSetup.PageHeight}.");

        doc.save(getArtifactsDir() + "PageSetup.PaperSizes.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PaperSizes.docx");

        Assert.assertEquals(PaperSize.TABLOID, doc.getSections().get(0).getPageSetup().getPaperSize());
        Assert.assertEquals(792.0d, doc.getSections().get(0).getPageSetup().getPageWidth());
        Assert.assertEquals(1224.0d, doc.getSections().get(0).getPageSetup().getPageHeight());
        Assert.assertEquals(PaperSize.A5, doc.getSections().get(1).getPageSetup().getPaperSize());
        Assert.assertEquals(419.55d, doc.getSections().get(1).getPageSetup().getPageWidth());
        Assert.assertEquals(595.30d, doc.getSections().get(1).getPageSetup().getPageHeight());
        Assert.assertEquals(PaperSize.CUSTOM, doc.getSections().get(2).getPageSetup().getPaperSize());
        Assert.assertEquals(620.0d, doc.getSections().get(2).getPageSetup().getPageWidth());
        Assert.assertEquals(480.0d, doc.getSections().get(2).getPageSetup().getPageHeight());
    }

    @Test
    public void columnsSameWidth() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.TextColumns
        //ExFor:TextColumnCollection
        //ExFor:TextColumnCollection.Spacing
        //ExFor:TextColumnCollection.SetCount
        //ExSummary:Shows how to create multiple evenly spaced columns in a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        columns.setSpacing(100.0);
        columns.setCount(2);

        builder.writeln("Column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Column 2.");

        doc.save(getArtifactsDir() + "PageSetup.ColumnsSameWidth.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.ColumnsSameWidth.docx");

        Assert.assertEquals(100.0d, doc.getFirstSection().getPageSetup().getTextColumns().getSpacing());
        Assert.assertEquals(2, doc.getFirstSection().getPageSetup().getTextColumns().getCount());
    }

    @Test
    public void customColumnWidth() throws Exception
    {
        //ExStart
        //ExFor:TextColumnCollection.EvenlySpaced
        //ExFor:TextColumnCollection.Item
        //ExFor:TextColumn
        //ExFor:TextColumn.Width
        //ExFor:TextColumn.SpaceAfter
        //ExSummary:Shows how to create unevenly spaced columns.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        PageSetup pageSetup = builder.getPageSetup();

        TextColumnCollection columns = pageSetup.getTextColumns();
        columns.setEvenlySpaced(false);
        columns.setCount(2);

        // Determine the amount of room that we have available for arranging columns.
        double contentWidth = pageSetup.getPageWidth() - pageSetup.getLeftMargin() - pageSetup.getRightMargin();

        Assert.assertEquals(470.30d, contentWidth, 0.01d);

        // Set the first column to be narrow.
        TextColumn column = columns.get(0);
        column.setWidth(100.0);
        column.setSpaceAfter(20.0);

        // Set the second column to take the rest of the space available within the margins of the page.
        column = columns.get(1);
        column.setWidth(contentWidth - column.getWidth() - column.getSpaceAfter());

        builder.writeln("Narrow column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Wide column 2.");

        doc.save(getArtifactsDir() + "PageSetup.CustomColumnWidth.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.CustomColumnWidth.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertFalse(pageSetup.getTextColumns().getEvenlySpaced());
        Assert.assertEquals(2, pageSetup.getTextColumns().getCount());
        Assert.assertEquals(100.0d, pageSetup.getTextColumns().get(0).getWidth());
        Assert.assertEquals(20.0d, pageSetup.getTextColumns().get(0).getSpaceAfter());
        Assert.assertEquals(470.3d, pageSetup.getTextColumns().get(1).getWidth());
        Assert.assertEquals(0.0d, pageSetup.getTextColumns().get(1).getSpaceAfter());
    }

    @Test (dataProvider = "verticalLineBetweenColumnsDataProvider")
    public void verticalLineBetweenColumns(boolean lineBetween) throws Exception
    {
        //ExStart
        //ExFor:TextColumnCollection.LineBetween
        //ExSummary:Shows how to separate columns with a vertical line.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure the current section's PageSetup object to divide the text into several columns.
        // Set the "LineBetween" property to "true" to put a dividing line between columns.
        // Set the "LineBetween" property to "false" to leave the space between columns blank.
        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        columns.setLineBetween(lineBetween);
        columns.setCount(3);

        builder.writeln("Column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Column 2.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Column 3.");

        doc.save(getArtifactsDir() + "PageSetup.VerticalLineBetweenColumns.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.VerticalLineBetweenColumns.docx");

        Assert.assertEquals(lineBetween, doc.getFirstSection().getPageSetup().getTextColumns().getLineBetween());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "verticalLineBetweenColumnsDataProvider")
	public static Object[][] verticalLineBetweenColumnsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void lineNumbers() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.LineStartingNumber
        //ExFor:PageSetup.LineNumberDistanceFromText
        //ExFor:PageSetup.LineNumberCountBy
        //ExFor:PageSetup.LineNumberRestartMode
        //ExFor:ParagraphFormat.SuppressLineNumbers
        //ExFor:LineNumberRestartMode
        //ExSummary:Shows how to enable line numbering for a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // We can use the section's PageSetup object to display numbers to the left of the section's text lines.
        // This is the same behavior as a List object,
        // but it covers the entire section and does not modify the text in any way.
        // Our section will restart the numbering on each new page from 1 and display the number,
        // if it is a multiple of 3, at 50pt to the left of the line.
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setLineStartingNumber(1);
        pageSetup.setLineNumberCountBy(3);
        pageSetup.setLineNumberRestartMode(LineNumberRestartMode.RESTART_PAGE);
        pageSetup.setLineNumberDistanceFromText(50.0d);

        for (int i = 1; i <= 25; i++)
            builder.writeln($"Line {i}.");

        // The line counter will skip any paragraph with the "SuppressLineNumbers" flag set to "true".
        // This paragraph is on the 15th line, which is a multiple of 3, and thus would normally display a line number.
        // The section's line counter will also ignore this line, treat the next line as the 15th,
        // and continue the count from that point onward.
        doc.getFirstSection().getBody().getParagraphs().get(14).getParagraphFormat().setSuppressLineNumbers(true);

        doc.save(getArtifactsDir() + "PageSetup.LineNumbers.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.LineNumbers.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(1, pageSetup.getLineStartingNumber());
        Assert.assertEquals(3, pageSetup.getLineNumberCountBy());
        Assert.assertEquals(LineNumberRestartMode.RESTART_PAGE, pageSetup.getLineNumberRestartMode());
        Assert.assertEquals(50.0d, pageSetup.getLineNumberDistanceFromText());
    }

    @Test
    public void pageBorderProperties() throws Exception
    {
        //ExStart
        //ExFor:Section.PageSetup
        //ExFor:PageSetup.BorderAlwaysInFront
        //ExFor:PageSetup.BorderDistanceFrom
        //ExFor:PageSetup.BorderAppliesTo
        //ExFor:PageBorderDistanceFrom
        //ExFor:PageBorderAppliesTo
        //ExFor:Border.DistanceFromText
        //ExSummary:Shows how to create a wide blue band border at the top of the first page.
        Document doc = new Document();

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setBorderAlwaysInFront(false);
        pageSetup.setBorderDistanceFrom(PageBorderDistanceFrom.PAGE_EDGE);
        pageSetup.setBorderAppliesTo(PageBorderAppliesTo.FIRST_PAGE);

        Border border = pageSetup.getBorders().getByBorderType(BorderType.TOP);
        border.setLineStyle(LineStyle.SINGLE);
        border.setLineWidth(30.0);
        border.setColor(Color.BLUE);
        border.setDistanceFromText(0.0);

        doc.save(getArtifactsDir() + "PageSetup.PageBorderProperties.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageBorderProperties.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertFalse(pageSetup.getBorderAlwaysInFront());
        Assert.assertEquals(PageBorderDistanceFrom.PAGE_EDGE, pageSetup.getBorderDistanceFrom());
        Assert.assertEquals(PageBorderAppliesTo.FIRST_PAGE, pageSetup.getBorderAppliesTo());

        border = pageSetup.getBorders().getByBorderType(BorderType.TOP);

        Assert.assertEquals(LineStyle.SINGLE, border.getLineStyle());
        Assert.assertEquals(30.0d, border.getLineWidth());
        Assert.assertEquals(Color.BLUE.getRGB(), border.getColor().getRGB());
        Assert.assertEquals(0.0d, border.getDistanceFromText());
    }

    @Test
    public void pageBorders() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.Borders
        //ExFor:Border.Shadow
        //ExFor:BorderCollection.LineStyle
        //ExFor:BorderCollection.LineWidth
        //ExFor:BorderCollection.Color
        //ExFor:BorderCollection.DistanceFromText
        //ExFor:BorderCollection.Shadow
        //ExSummary:Shows how to create green wavy page border with a shadow.
        Document doc = new Document();
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();

        pageSetup.getBorders().setLineStyle(LineStyle.DOUBLE_WAVE);
        pageSetup.getBorders().setLineWidth(2.0);
        pageSetup.getBorders().setColor(msColor.getGreen());
        pageSetup.getBorders().setDistanceFromText(24.0);
        pageSetup.getBorders().setShadow(true);

        doc.save(getArtifactsDir() + "PageSetup.PageBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageBorders.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        for (Border border : pageSetup.getBorders())
        {
            Assert.assertEquals(LineStyle.DOUBLE_WAVE, border.getLineStyle());
            Assert.assertEquals(2.0d, border.getLineWidth());
            Assert.assertEquals(msColor.getGreen().getRGB(), border.getColor().getRGB());
            Assert.assertEquals(24.0d, border.getDistanceFromText());
            Assert.assertTrue(border.getShadow());
        }
    }

    @Test
    public void pageNumbering() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.RestartPageNumbering
        //ExFor:PageSetup.PageStartingNumber
        //ExFor:PageSetup.PageNumberStyle
        //ExFor:DocumentBuilder.InsertField(String, String)
        //ExSummary:Shows how to set up page numbering in a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Section 1, page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Section 1, page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Section 1, page 3.");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.writeln("Section 2, page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Section 2, page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Section 2, page 3.");

        // Move the document builder to the first section's primary header,
        // which every page in that section will display.
        builder.moveToSection(0);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);

        // Insert a PAGE field, which will display the number of the current page.
        builder.write("Page ");
        builder.insertField("PAGE", "");

        // Configure the section to have the page count that PAGE fields display start from 5.
        // Also, configure all PAGE fields to display their page numbers using uppercase Roman numerals.
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setRestartPageNumbering(true);
        pageSetup.setPageStartingNumber(5);
        pageSetup.setPageNumberStyle(NumberStyle.UPPERCASE_ROMAN);

        // Create another primary header for the second section, with another PAGE field.
        builder.moveToSection(1);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write(" - ");
        builder.insertField("PAGE", "");
        builder.write(" - ");

        // Configure the section to have the page count that PAGE fields display start from 10.
        // Also, configure all PAGE fields to display their page numbers using Arabic numbers.
        pageSetup = doc.getSections().get(1).getPageSetup();
        pageSetup.setPageStartingNumber(10);
        pageSetup.setRestartPageNumbering(true);
        pageSetup.setPageNumberStyle(NumberStyle.ARABIC);

        doc.save(getArtifactsDir() + "PageSetup.PageNumbering.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageNumbering.docx");
        pageSetup = doc.getSections().get(0).getPageSetup();

        Assert.assertTrue(pageSetup.getRestartPageNumbering());
        Assert.assertEquals(5, pageSetup.getPageStartingNumber());
        Assert.assertEquals(NumberStyle.UPPERCASE_ROMAN, pageSetup.getPageNumberStyle());

        pageSetup = doc.getSections().get(1).getPageSetup();

        Assert.assertTrue(pageSetup.getRestartPageNumbering());
        Assert.assertEquals(10, pageSetup.getPageStartingNumber());
        Assert.assertEquals(NumberStyle.ARABIC, pageSetup.getPageNumberStyle());
    }

    @Test
    public void footnoteOptions() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.EndnoteOptions
        //ExFor:PageSetup.FootnoteOptions
        //ExSummary:Shows how to configure options affecting footnotes/endnotes in a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote reference text.");

        // Configure all footnotes in the first section to restart the numbering from 1
        // at each new page and display themselves directly beneath the text on every page.
        FootnoteOptions footnoteOptions = doc.getSections().get(0).getPageSetup().getFootnoteOptions();
        footnoteOptions.setPosition(FootnotePosition.BENEATH_TEXT);
        footnoteOptions.setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        footnoteOptions.setStartNumber(1);

        builder.write(" Hello again.");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Endnote reference text.");

        // Configure all endnotes in the first section to maintain a continuous count throughout the section,
        // starting from 1. Also, set them all to appear collected at the end of the document.
        EndnoteOptions endnoteOptions = doc.getSections().get(0).getPageSetup().getEndnoteOptions();
        endnoteOptions.setPosition(EndnotePosition.END_OF_DOCUMENT);
        endnoteOptions.setRestartRule(FootnoteNumberingRule.CONTINUOUS);
        endnoteOptions.setStartNumber(1);

        doc.save(getArtifactsDir() + "PageSetup.FootnoteOptions.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.FootnoteOptions.docx");
        footnoteOptions = doc.getFirstSection().getPageSetup().getFootnoteOptions();

        Assert.assertEquals(FootnotePosition.BENEATH_TEXT, footnoteOptions.getPosition());
        Assert.assertEquals(FootnoteNumberingRule.RESTART_PAGE, footnoteOptions.getRestartRule());
        Assert.assertEquals(1, footnoteOptions.getStartNumber());

        endnoteOptions = doc.getFirstSection().getPageSetup().getEndnoteOptions();

        Assert.assertEquals(EndnotePosition.END_OF_DOCUMENT, endnoteOptions.getPosition());
        Assert.assertEquals(FootnoteNumberingRule.CONTINUOUS, endnoteOptions.getRestartRule());
        Assert.assertEquals(1, endnoteOptions.getStartNumber());
    }

    @Test (dataProvider = "bidiDataProvider")
    public void bidi(boolean reverseColumns) throws Exception
    {
        //ExStart
        //ExFor:PageSetup.Bidi
        //ExSummary:Shows how to set the order of text columns in a section.
        Document doc = new Document();

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.getTextColumns().setCount(3);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.write("Column 2.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.write("Column 3.");

        // Set the "Bidi" property to "true" to arrange the columns starting from the page's right side.
        // The order of the columns will match the direction of the right-to-left text.
        // Set the "Bidi" property to "false" to arrange the columns starting from the page's left side.
        // The order of the columns will match the direction of the left-to-right text.
        pageSetup.setBidi(reverseColumns);

        doc.save(getArtifactsDir() + "PageSetup.Bidi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.Bidi.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(3, pageSetup.getTextColumns().getCount());
        Assert.assertEquals(reverseColumns, pageSetup.getBidi());
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "bidiDataProvider")
	public static Object[][] bidiDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void pageBorder() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.BorderSurroundsFooter
        //ExFor:PageSetup.BorderSurroundsHeader
        //ExSummary:Shows how to apply a border to the page and header/footer.
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world! This is the main body text.");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("This is the header.");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("This is the footer.");
        builder.moveToDocumentEnd();

        // Insert a blue double-line border.
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.getBorders().setLineStyle(LineStyle.DOUBLE);
        pageSetup.getBorders().setColor(Color.BLUE);

        // A section's PageSetup object has "BorderSurroundsHeader" and "BorderSurroundsFooter" flags that determine
        // whether a page border surrounds the main body text, also includes the header or footer, respectively.
        // Set the "BorderSurroundsHeader" flag to "true" to surround the header with our border,
        // and then set the "BorderSurroundsFooter" flag to leave the footer outside of the border.
        pageSetup.setBorderSurroundsHeader(true);
        pageSetup.setBorderSurroundsFooter(false);

        doc.save(getArtifactsDir() + "PageSetup.PageBorder.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageBorder.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertTrue(pageSetup.getBorderSurroundsHeader());
        Assert.assertFalse(pageSetup.getBorderSurroundsFooter());
    }

    @Test
    public void gutter() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.Gutter
        //ExFor:PageSetup.RtlGutter
        //ExFor:PageSetup.MultiplePages
        //ExSummary:Shows how to set gutter margins.
        Document doc = new Document();

        // Insert text that spans several pages.
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 6; i++)
        {
            builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                          "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.insertBreak(BreakType.PAGE_BREAK);
        }

        // A gutter adds whitespaces to either the left or right page margin,
        // which makes up for the center folding of pages in a book encroaching on the page's layout.
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();

        // Determine how much space our pages have for text within the margins and then add an amount to pad a margin. 
        Assert.assertEquals(470.30d, pageSetup.getPageWidth() - pageSetup.getLeftMargin() - pageSetup.getRightMargin(), 0.01d);
        
        pageSetup.setGutter(100.0d);

        // Set the "RtlGutter" property to "true" to place the gutter in a more suitable position for right-to-left text.
        pageSetup.setRtlGutter(true);

        // Set the "MultiplePages" property to "MultiplePagesType.MirrorMargins" to alternate
        // the left/right page side position of margins every page.
        pageSetup.setMultiplePages(MultiplePagesType.MIRROR_MARGINS);

        doc.save(getArtifactsDir() + "PageSetup.Gutter.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.Gutter.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(100.0d, pageSetup.getGutter());
        Assert.assertTrue(pageSetup.getRtlGutter());
        Assert.assertEquals(MultiplePagesType.MIRROR_MARGINS, pageSetup.getMultiplePages());
    }

    @Test
    public void booklet() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.Gutter
        //ExFor:PageSetup.MultiplePages
        //ExFor:PageSetup.SheetsPerBooklet
        //ExSummary:Shows how to configure a document that can be printed as a book fold.
        Document doc = new Document();

        // Insert text that spans 16 pages.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("My Booklet:");

        for (int i = 0; i < 15; i++)
        {
            builder.insertBreak(BreakType.PAGE_BREAK);
            builder.write($"Booklet face #{i}");
        }

        // Configure the first section's "PageSetup" property to print the document in the form of a book fold.
        // When we print this document on both sides, we can take the pages to stack them
        // and fold them all down the middle at once. The contents of the document will line up into a book fold.
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);

        // We can only specify the number of sheets in multiples of 4.
        pageSetup.setSheetsPerBooklet(4);

        doc.save(getArtifactsDir() + "PageSetup.Booklet.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.Booklet.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(MultiplePagesType.BOOK_FOLD_PRINTING, pageSetup.getMultiplePages());
        Assert.assertEquals(4, pageSetup.getSheetsPerBooklet());
    }

    @Test
    public void setTextOrientation() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.TextOrientation
        //ExSummary:Shows how to set text orientation.
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Set the "TextOrientation" property to "TextOrientation.Upward" to rotate all the text 90 degrees
        // to the right so that all left-to-right text now goes top-to-bottom.
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setTextOrientation(TextOrientation.UPWARD);

        doc.save(getArtifactsDir() + "PageSetup.SetTextOrientation.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.SetTextOrientation.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(TextOrientation.UPWARD, pageSetup.getTextOrientation());
    }

    //ExStart
    //ExFor:PageSetup.SuppressEndnotes
    //ExFor:Body.ParentSection
    //ExSummary:Shows how to store endnotes at the end of each section, and modify their positions.
    @Test //ExSkip
    public void suppressEndnotes() throws Exception
    {
        Document doc = new Document();
        doc.removeAllChildren();

        // By default, a document compiles all endnotes at its end. 
        Assert.assertEquals(EndnotePosition.END_OF_DOCUMENT, doc.getEndnoteOptions().getPosition());

        // We use the "Position" property of the document's "EndnoteOptions" object
        // to collect endnotes at the end of each section instead. 
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);

        insertSectionWithEndnote(doc, "Section 1", "Endnote 1, will stay in section 1");
        insertSectionWithEndnote(doc, "Section 2", "Endnote 2, will be pushed down to section 3");
        insertSectionWithEndnote(doc, "Section 3", "Endnote 3, will stay in section 3");

        // While getting sections to display their respective endnotes, we can set the "SuppressEndnotes" flag
        // of a section's "PageSetup" object to "true" to revert to the default behavior and pass its endnotes
        // onto the next section.
        PageSetup pageSetup = doc.getSections().get(1).getPageSetup();
        pageSetup.setSuppressEndnotes(true);

        doc.save(getArtifactsDir() + "PageSetup.SuppressEndnotes.docx");
        testSuppressEndnotes(new Document(getArtifactsDir() + "PageSetup.SuppressEndnotes.docx")); //ExSkip
    }

    /// <summary>
    /// Append a section with text and an endnote to a document.
    /// </summary>
    private static void insertSectionWithEndnote(Document doc, String sectionBodyText, String endnoteText)
    {
        Section section = new Section(doc);

        doc.appendChild(section);

        Body body = new Body(doc);
        section.appendChild(body);

        Assert.assertEquals(section, body.getParentNode());

        Paragraph para = new Paragraph(doc);
        body.appendChild(para);

        Assert.assertEquals(body, para.getParentNode());

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveTo(para);
        builder.write(sectionBodyText);
        builder.insertFootnote(FootnoteType.ENDNOTE, endnoteText);
    }
    //ExEnd

    private static void testSuppressEndnotes(Document doc)
    {
        PageSetup pageSetup = doc.getSections().get(1).getPageSetup();

        Assert.assertTrue(pageSetup.getSuppressEndnotes());
    }
}
