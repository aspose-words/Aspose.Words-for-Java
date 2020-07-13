package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.standard.Media;
import java.awt.*;
import java.text.MessageFormat;

public class ExPageSetup extends ApiExampleBase {
    @Test
    public void clearFormatting() throws Exception {
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
        //ExSummary:Shows how to insert sections using DocumentBuilder, specify page setup for a section and reset page setup to defaults.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Modify the first section in the document
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setVerticalAlignment(PageVerticalAlignment.CENTER);
        builder.writeln("Section 1, landscape oriented and text vertically centered.");

        // Start a new section and reset its formatting to defaults
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.getPageSetup().clearFormatting();
        builder.writeln("Section 2, back to default Letter paper size, portrait orientation and top alignment.");

        doc.save(getArtifactsDir() + "PageSetup.ClearFormatting.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.ClearFormatting.docx");

        Assert.assertEquals(Orientation.LANDSCAPE, doc.getSections().get(0).getPageSetup().getOrientation());
        Assert.assertEquals(PageVerticalAlignment.CENTER, doc.getSections().get(0).getPageSetup().getVerticalAlignment());

        Assert.assertEquals(Orientation.PORTRAIT, doc.getSections().get(1).getPageSetup().getOrientation());
        Assert.assertEquals(PageVerticalAlignment.TOP, doc.getSections().get(1).getPageSetup().getVerticalAlignment());
    }

    @Test
    public void differentHeaders() throws Exception {
        //ExStart
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExFor:PageSetup.LayoutMode
        //ExFor:PageSetup.CharactersPerLine
        //ExFor:PageSetup.LinesPerPage
        //ExFor:SectionLayoutMode
        //ExSummary:Shows how to create headers and footers different for first, even and odd pages using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setDifferentFirstPageHeaderFooter(true);
        pageSetup.setOddAndEvenPagesHeaderFooter(true);
        pageSetup.setLayoutMode(SectionLayoutMode.LINE_GRID);
        pageSetup.setCharactersPerLine(1);
        pageSetup.setLinesPerPage(1);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.writeln("First page header.");

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.writeln("Even pages header.");

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Odd pages header.");

        // Move back to the main story of the first section
        builder.moveToSection(0);
        builder.writeln("Text page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Text page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Text page 3.");

        doc.save(getArtifactsDir() + "PageSetup.DifferentHeaders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.DifferentHeaders.docx");

        Assert.assertTrue(pageSetup.getDifferentFirstPageHeaderFooter());
        Assert.assertTrue(pageSetup.getOddAndEvenPagesHeaderFooter());
        Assert.assertEquals(SectionLayoutMode.LINE_GRID, doc.getFirstSection().getPageSetup().getLayoutMode());
        Assert.assertEquals(1, doc.getFirstSection().getPageSetup().getCharactersPerLine());
        Assert.assertEquals(1, doc.getFirstSection().getPageSetup().getLinesPerPage());
    }

    @Test
    public void setSectionStart() throws Exception {
        //ExStart
        //ExFor:SectionStart
        //ExFor:PageSetup.SectionStart
        //ExFor:Document.Sections
        //ExSummary:Shows how to specify how the section starts, from a new page, on the same page or other.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text to the first section and that comes with a blank document,
        // then add a new section that starts a new page and give it text as well
        builder.writeln("This text is in section 1.");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.writeln("This text is in section 2.");

        // Section break types determine how a new section gets split from the previous section
        // By inserting a "SectionBreakNewPage" type section break, we've set this section's SectionStart value to "NewPage" 
        Assert.assertEquals(SectionStart.NEW_PAGE, doc.getSections().get(1).getPageSetup().getSectionStart());

        // Insert a new column section the same way
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_COLUMN);
        builder.writeln("This text is in section 3.");

        Assert.assertEquals(SectionStart.NEW_COLUMN, doc.getSections().get(2).getPageSetup().getSectionStart());

        // We can change the types of section breaks by assigning different values to each section's SectionStart
        // Setting their values to "Continuous" will put no visible breaks between sections
        // and will leave all the content of this document on one page
        doc.getSections().get(1).getPageSetup().setSectionStart(SectionStart.CONTINUOUS);
        doc.getSections().get(2).getPageSetup().setSectionStart(SectionStart.CONTINUOUS);

        doc.save(getArtifactsDir() + "PageSetup.SetSectionStart.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.SetSectionStart.docx");

        Assert.assertEquals(SectionStart.NEW_PAGE, doc.getSections().get(0).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.CONTINUOUS, doc.getSections().get(1).getPageSetup().getSectionStart());
        Assert.assertEquals(SectionStart.CONTINUOUS, doc.getSections().get(2).getPageSetup().getSectionStart());
    }

    @Test(enabled = false, description = "Run only when the printer driver is installed")
    public void paperTrayForDifferentPaperType() throws Exception {
        //ExStart
        //ExFor:PageSetup.FirstPageTray
        //ExFor:PageSetup.OtherPagesTray
        //ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
        Document doc = new Document();

        // Choose the default printer to be used for printing this document
        PrintService printService = PrintServiceLookup.lookupDefaultPrintService();
        Media[] trays = (Media[]) printService.getSupportedAttributeValues(Media.class, null, null);

        // This is the tray we will use for A4 paper size
        // This is the first tray in the media set
        int printerTrayForA4 = trays[0].getValue();
        // This is the tray we will use Letter paper size
        // This is the second tray in the media set
        int printerTrayForLetter = trays[1].getValue();

        // Set the tray used for each section based off the paper size used in the section
        for (Section section : doc.getSections()) {
            if (section.getPageSetup().getPaperSize() == PaperSize.LETTER) {
                section.getPageSetup().setFirstPageTray(printerTrayForLetter);
                section.getPageSetup().setOtherPagesTray(printerTrayForLetter);
            } else if (section.getPageSetup().getPaperSize() == PaperSize.A4) {
                section.getPageSetup().setFirstPageTray(printerTrayForA4);
                section.getPageSetup().setOtherPagesTray(printerTrayForA4);
            }
        }
        //ExEnd

        for (Section section : DocumentHelper.saveOpen(doc).getSections()) {
            if (section.getPageSetup().getPaperSize() == com.aspose.words.PaperSize.LETTER) {
                Assert.assertEquals(printerTrayForLetter, section.getPageSetup().getFirstPageTray());
                Assert.assertEquals(printerTrayForLetter, section.getPageSetup().getOtherPagesTray());
            } else if (section.getPageSetup().getPaperSize() == com.aspose.words.PaperSize.A4) {
                Assert.assertEquals(printerTrayForA4, section.getPageSetup().getFirstPageTray());
                Assert.assertEquals(printerTrayForA4, section.getPageSetup().getOtherPagesTray());
            }
        }
    }

    @Test
    public void pageMargins() throws Exception {
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
        //ExSummary:Shows how to adjust paper size, orientation, margins and other settings for a section.
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

        builder.writeln("Hello world.");

        doc.save(getArtifactsDir() + "PageSetup.PageMargins.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageMargins.docx");

        Assert.assertEquals(PaperSize.LEGAL, doc.getFirstSection().getPageSetup().getPaperSize());
        Assert.assertEquals(Orientation.LANDSCAPE, doc.getFirstSection().getPageSetup().getOrientation());
        Assert.assertEquals(72.0d, doc.getFirstSection().getPageSetup().getTopMargin());
        Assert.assertEquals(72.0d, doc.getFirstSection().getPageSetup().getBottomMargin());
        Assert.assertEquals(108.0d, doc.getFirstSection().getPageSetup().getLeftMargin());
        Assert.assertEquals(108.0d, doc.getFirstSection().getPageSetup().getRightMargin());
        Assert.assertEquals(14.4d, doc.getFirstSection().getPageSetup().getHeaderDistance());
        Assert.assertEquals(14.4d, doc.getFirstSection().getPageSetup().getFooterDistance());
    }

    @Test
    public void columnsSameWidth() throws Exception {
        //ExStart
        //ExFor:PageSetup.TextColumns
        //ExFor:TextColumnCollection
        //ExFor:TextColumnCollection.Spacing
        //ExFor:TextColumnCollection.SetCount
        //ExSummary:Shows how to create multiple evenly spaced columns in a section using DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        // Make spacing between columns wider
        columns.setSpacing(100.0);
        // This creates two columns of equal width
        columns.setCount(2);

        builder.writeln("Text in column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Text in column 2.");

        doc.save(getArtifactsDir() + "PageSetup.ColumnsSameWidth.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.ColumnsSameWidth.docx");

        Assert.assertEquals(100.0d, doc.getFirstSection().getPageSetup().getTextColumns().getSpacing());
        Assert.assertEquals(2, doc.getFirstSection().getPageSetup().getTextColumns().getCount());
    }

    @Test
    public void customColumnWidth() throws Exception {
        //ExStart
        //ExFor:TextColumnCollection.LineBetween
        //ExFor:TextColumnCollection.EvenlySpaced
        //ExFor:TextColumnCollection.Item
        //ExFor:TextColumn
        //ExFor:TextColumn.Width
        //ExFor:TextColumn.SpaceAfter
        //ExSummary:Shows how to set widths of columns.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        // Show vertical line between columns
        columns.setLineBetween(true);
        // Indicate we want to create column with different widths
        columns.setEvenlySpaced(false);
        // Create two columns, note they will be created with zero widths, need to set them
        columns.setCount(2);

        // Set the first column to be narrow
        TextColumn column = columns.get(0);
        column.setWidth(100.0);
        column.setSpaceAfter(20.0);

        // Set the second column to take the rest of the space available on the page
        column = columns.get(1);
        PageSetup pageSetup = builder.getPageSetup();
        double contentWidth = pageSetup.getPageWidth() - pageSetup.getLeftMargin() - pageSetup.getRightMargin();
        column.setWidth(contentWidth - column.getWidth() - column.getSpaceAfter());

        builder.writeln("Narrow column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Wide column 2.");

        doc.save(getArtifactsDir() + "PageSetup.CustomColumnWidth.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.CustomColumnWidth.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertTrue(pageSetup.getTextColumns().getLineBetween());
        Assert.assertFalse(pageSetup.getTextColumns().getEvenlySpaced());
        Assert.assertEquals(2, pageSetup.getTextColumns().getCount());
        Assert.assertEquals(100.0d, pageSetup.getTextColumns().get(0).getWidth());
        Assert.assertEquals(20.0d, pageSetup.getTextColumns().get(0).getSpaceAfter());
        Assert.assertEquals(468.0d, pageSetup.getTextColumns().get(1).getWidth());
        Assert.assertEquals(0.0d, pageSetup.getTextColumns().get(1).getSpaceAfter());
    }

    @Test
    public void lineNumbers() throws Exception {
        //ExStart
        //ExFor:PageSetup.LineStartingNumber
        //ExFor:PageSetup.LineNumberDistanceFromText
        //ExFor:PageSetup.LineNumberCountBy
        //ExFor:PageSetup.LineNumberRestartMode
        //ExFor:ParagraphFormat.SuppressLineNumbers
        //ExFor:LineNumberRestartMode
        //ExSummary:Shows how to enable Microsoft Word line numbering for a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Line numbering for each section can be configured via PageSetup
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setLineStartingNumber(1);
        pageSetup.setLineNumberCountBy(3);
        pageSetup.setLineNumberRestartMode(LineNumberRestartMode.RESTART_PAGE);
        pageSetup.setLineNumberDistanceFromText(50.0d);

        // LineNumberCountBy is set to 3, so every line that's a multiple of 3
        // will display that line number to the left of the text
        for (int i = 1; i <= 25; i++)
            builder.writeln(MessageFormat.format("Line {0}.", i));

        // The line counter will skip any paragraph with this flag set to true
        // Normally, the number "15" would normally appear next to this paragraph, which says "Line 15"
        // Since we set this flag to true and this paragraph is not counted by numbering,
        // number 15 will appear next to the next paragraph, "Line 16", and from then on counting will carry on as normal
        // until it will restart according to LineNumberRestartMode
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
    public void pageBorderProperties() throws Exception {
        //ExStart
        //ExFor:Section.PageSetup
        //ExFor:PageSetup.BorderAlwaysInFront
        //ExFor:PageSetup.BorderDistanceFrom
        //ExFor:PageSetup.BorderAppliesTo
        //ExFor:PageBorderDistanceFrom
        //ExFor:PageBorderAppliesTo
        //ExFor:Border.DistanceFromText
        //ExSummary:Shows how to create a page border that looks like a wide blue band at the top of the first page only.
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
    public void pageBorders() throws Exception {
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
        pageSetup.getBorders().setColor(Color.GREEN);
        pageSetup.getBorders().setDistanceFromText(24.0);
        pageSetup.getBorders().setShadow(true);

        doc.save(getArtifactsDir() + "PageSetup.PageBorders.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageBorders.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        for (Border border : pageSetup.getBorders()) {
            Assert.assertEquals(LineStyle.DOUBLE_WAVE, border.getLineStyle());
            Assert.assertEquals(2.0d, border.getLineWidth());
            Assert.assertEquals(Color.GREEN.getRGB(), border.getColor().getRGB());
            Assert.assertEquals(24.0d, border.getDistanceFromText());
            Assert.assertTrue(border.getShadow());
        }
    }

    @Test
    public void pageNumbering() throws Exception {
        //ExStart
        //ExFor:PageSetup.RestartPageNumbering
        //ExFor:PageSetup.PageStartingNumber
        //ExFor:PageSetup.PageNumberStyle
        //ExFor:DocumentBuilder.InsertField(String, String)
        //ExSummary:Shows how to control page numbering per section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.writeln("Section 2");

        // Use document builder to create a header with a page number field for the first section
        // The page number will look like "Page V"
        builder.moveToSection(0);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Page ");
        builder.insertField("PAGE", "");

        // Set first section page numbering
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setRestartPageNumbering(true);
        pageSetup.setPageStartingNumber(5);
        pageSetup.setPageNumberStyle(NumberStyle.UPPERCASE_ROMAN);

        // Create a header for the section
        // The page number will look like " - 10 - ".
        builder.moveToSection(1);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write(" - ");
        builder.insertField("PAGE", "");
        builder.write(" - ");

        // Set second section page numbering
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
    public void footnoteOptions() throws Exception {
        //ExStart
        //ExFor:PageSetup.EndnoteOptions
        //ExFor:PageSetup.FootnoteOptions
        //ExSummary:Shows how to set options for footnotes and endnotes in current section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text and a reference for it in the form of a footnote
        builder.write("Hello world!.");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Footnote reference text.");

        // Set options for footnote position and numbering
        FootnoteOptions footnoteOptions = doc.getSections().get(0).getPageSetup().getFootnoteOptions();
        footnoteOptions.setPosition(FootnotePosition.BENEATH_TEXT);
        footnoteOptions.setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        footnoteOptions.setStartNumber(1);

        // Endnotes also have a similar options object
        builder.write(" Hello again.");
        builder.insertFootnote(FootnoteType.FOOTNOTE, "Endnote reference text.");

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

    @Test
    public void bidi() throws Exception {
        //ExStart
        //ExFor:PageSetup.Bidi
        //ExSummary:Shows how to change the order of columns.
        Document doc = new Document();

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.getTextColumns().setCount(3);

        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.write("Column 2.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.write("Column 3.");

        // Reverse the order of the columns
        pageSetup.setBidi(true);

        doc.save(getArtifactsDir() + "PageSetup.Bidi.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.Bidi.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(3, pageSetup.getTextColumns().getCount());
        Assert.assertTrue(pageSetup.getBidi());
    }

    @Test
    public void pageBorder() throws Exception {
        //ExStart
        //ExFor:PageSetup.BorderSurroundsFooter
        //ExFor:PageSetup.BorderSurroundsHeader
        //ExSummary:Shows how to apply a border to the page and header/footer.
        Document doc = new Document();

        // Insert header and footer text
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Header");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Footer");
        builder.moveToDocumentEnd();

        // Insert a page border and set the color and line style
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.getBorders().setLineStyle(LineStyle.DOUBLE);
        pageSetup.getBorders().setColor(Color.BLUE);

        // By default, page borders don't surround headers and footers
        // We can change that by setting these flags
        pageSetup.setBorderSurroundsFooter(true);
        pageSetup.setBorderSurroundsHeader(true);

        doc.save(getArtifactsDir() + "PageSetup.PageBorder.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.PageBorder.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertTrue(pageSetup.getBorderSurroundsFooter());
        Assert.assertTrue(pageSetup.getBorderSurroundsHeader());
    }

    @Test
    public void gutter() throws Exception {
        //ExStart
        //ExFor:PageSetup.Gutter
        //ExFor:PageSetup.RtlGutter
        //ExFor:PageSetup.MultiplePages
        //ExSummary:Shows how to set gutter margins.
        Document doc = new Document();

        // Insert text spanning several pages
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 6; i++) {
            builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            builder.insertBreak(BreakType.PAGE_BREAK);
        }

        // We can access the gutter margin in the section's page options,
        // which is a margin which is added to the page margin at one side of the page
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setGutter(100.0d);

        // If our text is LTR, the gutter will appear on the left side of the page
        // Setting this flag will move it to the right side
        pageSetup.setRtlGutter(true);

        // Mirroring the margins will make the gutter alternate in position from page to page
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
    public void booklet() throws Exception {
        //ExStart
        //ExFor:PageSetup.SheetsPerBooklet
        //ExSummary:Shows how to create a booklet.
        Document doc = new Document();

        // Use a document builder to create 16 pages of content that will be compiled in a booklet
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("My Booklet:");

        for (int i = 0; i < 15; i++) {
            builder.insertBreak(BreakType.PAGE_BREAK);
            builder.write(MessageFormat.format("Booklet face #{0}", i));
        }

        // Set the number of sheets that will be used by the printer to create the booklet
        // After being printed on both sides, the sheets can be stacked and folded down the centre
        // The contents that we placed in such a way that they will be in order once the booklet is folded
        // We can only specify the number of sheets in multiples of 4
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
        pageSetup.setSheetsPerBooklet(4);

        doc.save(getArtifactsDir() + "PageSetup.Booklet.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.Booklet.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(MultiplePagesType.BOOK_FOLD_PRINTING, pageSetup.getMultiplePages());
        Assert.assertEquals(4, pageSetup.getSheetsPerBooklet());
    }

    @Test
    public void sectionTextOrientation() throws Exception {
        //ExStart
        //ExFor:PageSetup.TextOrientation
        //ExSummary:Shows how to set text orientation.
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Setting this value will rotate the section's text 90 degrees to the right
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setTextOrientation(TextOrientation.UPWARD);

        doc.save(getArtifactsDir() + "PageSetup.SectionTextOrientation.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "PageSetup.SectionTextOrientation.docx");
        pageSetup = doc.getFirstSection().getPageSetup();

        Assert.assertEquals(TextOrientation.UPWARD, pageSetup.getTextOrientation());
    }

    //ExStart
    //ExFor:PageSetup.SuppressEndnotes
    //ExFor:Body.ParentSection
    //ExSummary:Shows how to store endnotes at the end of each section instead of the document and manipulate their positions.
    @Test //ExSkip
    public void suppressEndnotes() throws Exception {
        // Create a new document and make it empty
        Document doc = new Document();
        doc.removeAllChildren();

        // Normally endnotes are all stored at the end of a document, but this option lets us store them at the end of each section
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);

        // Create 3 new sections, each having a paragraph and an endnote at the end
        insertSection(doc, "Section 1", "Endnote 1, will stay in section 1");
        insertSection(doc, "Section 2", "Endnote 2, will be pushed down to section 3");
        insertSection(doc, "Section 3", "Endnote 3, will stay in section 3");

        // Each section contains its own page setup object
        // Setting this value will push this section's endnotes down to the next section
        PageSetup pageSetup = doc.getSections().get(1).getPageSetup();
        pageSetup.setSuppressEndnotes(true);

        doc.save(getArtifactsDir() + "PageSetup.SuppressEndnotes.docx");
        testSuppressEndnotes(new Document(getArtifactsDir() + "PageSetup.SuppressEndnotes.docx")); //ExSkip
    }

    /// <summary>
    /// Add a section to the end of a document, give it a body and a paragraph, then add text and an endnote to that paragraph.
    /// </summary>
    private void insertSection(Document doc, String sectionBodyText, String endnoteText) {
        Section section = new Section(doc);

        doc.appendChild(section);

        Body body = new Body(doc);
        section.appendChild(body);

        Assert.assertEquals(body.getParentNode(), section);

        Paragraph para = new Paragraph(doc);
        body.appendChild(para);

        Assert.assertEquals(para.getParentNode(), body);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveTo(para);
        builder.write(sectionBodyText);
        builder.insertFootnote(FootnoteType.ENDNOTE, endnoteText);
    }
    //ExEnd

    private static void testSuppressEndnotes(Document doc) {
        PageSetup pageSetup = doc.getSections().get(1).getPageSetup();

        Assert.assertTrue(pageSetup.getSuppressEndnotes());
    }
}
