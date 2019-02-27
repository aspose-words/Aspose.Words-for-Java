//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.lang.SystemUtils;
import org.testng.annotations.Test;

import java.awt.Color;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.standard.Media;

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
        //ExSummary:Shows how to insert sections using DocumentBuilder, specify page setup for a section and reset page setup to defaults.
        DocumentBuilder builder = new DocumentBuilder();

        // Modify the first section in the document.
        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setVerticalAlignment(PageVerticalAlignment.CENTER);
        builder.writeln("Section 1, landscape oriented and text vertically centered.");

        // Start a new section and reset its formatting to defaults.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.getPageSetup().clearFormatting();
        builder.writeln("Section 2, back to default Letter paper size, portrait orientation and top alignment.");

        builder.getDocument().save(getArtifactsDir() + "PageSetup.ClearFormatting.doc");
        //ExEnd
    }

    @Test
    public void differentHeaders() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.DifferentFirstPageHeaderFooter
        //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
        //ExFor:PageSetup.LayoutMode
        //ExFor:PageSetup.CharactersPerLine
        //ExFor:PageSetup.LinesPerPage
        //ExSummary:Creates headers and footers different for first, even and odd pages using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        PageSetup ps = builder.getPageSetup();
        ps.setDifferentFirstPageHeaderFooter(true);
        ps.setOddAndEvenPagesHeaderFooter(true);
        ps.setLayoutMode(SectionLayoutMode.LINE_GRID);
        ps.setCharactersPerLine(1);
        ps.setLinesPerPage(1);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_FIRST);
        builder.writeln("First page header.");

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_EVEN);
        builder.writeln("Even pages header.");

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Odd pages header.");

        // Move back to the main story of the first section.
        builder.moveToSection(0);
        builder.writeln("Text page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Text page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Text page 3.");

        builder.getDocument().save(getArtifactsDir() + "PageSetup.DifferentHeaders.doc");
        //ExEnd
    }

    @Test
    public void sectionStart() throws Exception
    {
        //ExStart
        //ExFor:SectionStart
        //ExFor:PageSetup.SectionStart
        //ExFor:Document.Sections
        //ExSummary:Specifies how the section starts, from a new page, on the same page or other.
        Document doc = new Document();
        doc.getSections().get(0).getPageSetup().setSectionStart(com.aspose.words.SectionStart.CONTINUOUS);
        //ExEnd
    }

    @Test
    public void defaultPaperTray() throws Exception
    {
        if (!SystemUtils.IS_OS_LINUX) {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Changes all sections in a document to use the default paper tray of the selected printer.
            Document doc = new Document();

            // Find the printer that will be used for printing this document. In this case it is the default printer.
            // You can define a specific printer by using PrintServiceLookup.lookupPrintServices.
            PrintService printService = PrintServiceLookup.lookupDefaultPrintService();
            Media defaultTray = (Media) printService.getDefaultAttributeValue(Media.class);

            // The paper tray value stored in documents is completely printer specific. This means
            // The code below resets all page tray values to use the current printers default tray.
            // You can enumerate getSupportedAttributeValues for Media type to find the other valid
            // paper tray values of the selected printer.
            for (Section section : doc.getSections()) {
                section.getPageSetup().setFirstPageTray(defaultTray.getValue());
                section.getPageSetup().setOtherPagesTray(defaultTray.getValue());
            }
            //ExEnd
        }
    }

    @Test
    public void paperTrayForDifferentPaperType() throws Exception
    {
        if (!SystemUtils.IS_OS_LINUX) {
            //ExStart
            //ExFor:PageSetup.FirstPageTray
            //ExFor:PageSetup.OtherPagesTray
            //ExSummary:Shows how to set up printing using different printer trays for different paper sizes.
            Document doc = new Document();

            // Choose the default printer to be used for printing this document.
            PrintService printService = PrintServiceLookup.lookupDefaultPrintService();
            Media[] trays = (Media[]) printService.getSupportedAttributeValues(Media.class, null, null);

            // This is the tray we will use for A4 paper size. This is the first tray in the media set.
            int printerTrayForA4 = trays[0].getValue();
            // This is the tray we will use Letter paper size. This is the second tray in the media set.
            int printerTrayForLetter = trays[1].getValue();

            // Set the tray used for each section based off the paper size used in the section.
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
        //ExSummary:Specifies paper size, orientation, margins and other settings for a section.
        DocumentBuilder builder = new DocumentBuilder();

        PageSetup ps = builder.getPageSetup();
        ps.setPaperSize(PaperSize.LEGAL);
        ps.setOrientation(Orientation.LANDSCAPE);
        ps.setTopMargin(ConvertUtil.inchToPoint(1.0));
        ps.setBottomMargin(ConvertUtil.inchToPoint(1.0));
        ps.setLeftMargin(ConvertUtil.inchToPoint(1.5));
        ps.setRightMargin(ConvertUtil.inchToPoint(1.5));
        ps.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        ps.setFooterDistance(ConvertUtil.inchToPoint(0.2));

        builder.writeln("Hello world.");

        builder.getDocument().save(getArtifactsDir() + "PageSetup.PageMargins.doc");
        //ExEnd
    }

    @Test
    public void columnsSameWidth() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.TextColumns
        //ExFor:TextColumnCollection
        //ExFor:TextColumnCollection.Spacing
        //ExFor:TextColumnCollection.SetCount
        //ExSummary:Creates multiple evenly spaced columns in a section using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        // Make spacing between columns wider.
        columns.setSpacing(100);
        // This creates two columns of equal width.
        columns.setCount(2);

        builder.writeln("Text in column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Text in column 2.");

        builder.getDocument().save(getArtifactsDir() + "PageSetup.ColumnsSameWidth.doc");
        //ExEnd
    }

    @Test
    public void columnsCustomWidth() throws Exception
    {
        //ExStart
        //ExFor:TextColumnCollection.LineBetween
        //ExFor:TextColumnCollection.EvenlySpaced
        //ExFor:TextColumnCollection.Item
        //ExFor:TextColumn
        //ExFor:TextColumn.Width
        //ExFor:TextColumn.SpaceAfter
        //ExSummary:Creates multiple columns of different widths in a section using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder();

        TextColumnCollection columns = builder.getPageSetup().getTextColumns();
        // Show vertical line between columns.
        columns.setLineBetween(true);
        // Indicate we want to create column with different widths.
        columns.setEvenlySpaced(false);
        // Create two columns, note they will be created with zero widths, need to set them.
        columns.setCount(2);

        // Set the first column to be narrow.
        TextColumn c1 = columns.get(0);
        c1.setWidth(100);
        c1.setSpaceAfter(20);

        // Set the second column to take the rest of the space available on the page.
        TextColumn c2 = columns.get(1);
        PageSetup ps = builder.getPageSetup();
        double contentWidth = ps.getPageWidth() - ps.getLeftMargin() - ps.getRightMargin();
        c2.setWidth(contentWidth - c1.getWidth() - c1.getSpaceAfter());

        builder.writeln("Narrow column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Wide column 2.");

        builder.getDocument().save(getArtifactsDir() + "PageSetup.ColumnsCustomWidth.doc");
        //ExEnd
    }

    @Test
    public void lineNumbers() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.LineStartingNumber
        //ExFor:PageSetup.LineNumberCountBy
        //ExFor:PageSetup.LineNumberRestartMode
        //ExFor:LineNumberRestartMode
        //ExSummary:Turns on Microsoft Word line numbering for a section.
        DocumentBuilder builder = new DocumentBuilder();

        PageSetup ps = builder.getPageSetup();
        ps.setLineStartingNumber(1);
        ps.setLineNumberCountBy(5);
        ps.setLineNumberRestartMode(LineNumberRestartMode.RESTART_PAGE);

        for (int i = 1; i <= 20; i++)
            builder.writeln(java.text.MessageFormat.format("Line {0}.", i));

        builder.getDocument().save(getArtifactsDir() + "PageSetup.LineNumbers.doc");
        //ExEnd
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
        //ExSummary:Creates a page border that looks like a wide blue band at the top of the first page only.
        Document doc = new Document();

        PageSetup ps = doc.getSections().get(0).getPageSetup();
        ps.setBorderAlwaysInFront(false);
        ps.setBorderDistanceFrom(PageBorderDistanceFrom.PAGE_EDGE);
        ps.setBorderAppliesTo(PageBorderAppliesTo.FIRST_PAGE);

        Border border = ps.getBorders().getByBorderType(BorderType.TOP);
        border.setLineStyle(LineStyle.SINGLE);
        border.setLineWidth(30);
        border.setColor(Color.BLUE);
        border.setDistanceFromText(0);

        doc.save(getArtifactsDir() + "PageSetup.PageBorderTop.doc");
        //ExEnd
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
        //ExSummary:Creates a fancy looking green wavy page border with a shadow.
        Document doc = new Document();
        PageSetup ps = doc.getSections().get(0).getPageSetup();

        ps.getBorders().setLineStyle(LineStyle.DOUBLE_WAVE);
        ps.getBorders().setLineWidth(2);
        ps.getBorders().setColor(Color.GREEN);
        ps.getBorders().setDistanceFromText(24);
        ps.getBorders().setShadow(true);

        doc.save(getArtifactsDir() + "PageSetup.PageBorders.doc");
        //ExEnd
    }

    @Test
    public void pageNumbering() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.RestartPageNumbering
        //ExFor:PageSetup.PageStartingNumber
        //ExFor:PageSetup.PageNumberStyle
        //ExFor:DocumentBuilder.InsertField(String, String)
        //ExSummary:Shows how to control page numbering per section.
        // This document has two sections, but no page numbers yet.
        Document doc = new Document(getMyDir() + "PageSetup.PageNumbering.doc");

        // Use document builder to create a header with a page number field for the first section.
        // The page number will look like "Page V".
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.moveToSection(0);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Page ");
        builder.insertField("PAGE", "");

        // Set first section page numbering.
        Section section = doc.getSections().get(0);
        section.getPageSetup().setRestartPageNumbering(true);
        section.getPageSetup().setPageStartingNumber(5);
        section.getPageSetup().setPageNumberStyle(NumberStyle.UPPERCASE_ROMAN);

        // Create a header for the section. 
        // The page number will look like " - 10 - ".
        builder.moveToSection(1);
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write(" - ");
        builder.insertField("PAGE", "");
        builder.write(" - ");

        // Set second section page numbering.
        section = doc.getSections().get(1);
        section.getPageSetup().setPageStartingNumber(10);
        section.getPageSetup().setRestartPageNumbering(true);
        section.getPageSetup().setPageNumberStyle(NumberStyle.ARABIC);

        doc.save(getArtifactsDir() + "PageSetup.PageNumbering.doc");
        //ExEnd
    }

    @Test
    public void footnoteOptions() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.FootnoteOptions
        //ExSummary: Shows how to set footnote options for current section
        Document doc = new Document();

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();

        pageSetup.getFootnoteOptions().setPosition(FootnotePosition.BOTTOM_OF_PAGE);
        pageSetup.getFootnoteOptions().setNumberStyle(NumberStyle.BULLET);
        pageSetup.getFootnoteOptions().setStartNumber(1);
        pageSetup.getFootnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        pageSetup.getFootnoteOptions().setColumns(0);
        //ExEnd
    }

    @Test
    public void endnoteOptions() throws Exception
    {
        //ExStart
        //ExFor:PageSetup.EndnoteOptions
        //ExSummary: Shows how to set endnote options for current section
        Document doc = new Document();

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();

        pageSetup.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        pageSetup.getEndnoteOptions().setNumberStyle(NumberStyle.BULLET);
        pageSetup.getEndnoteOptions().setStartNumber(1);
        pageSetup.getEndnoteOptions().setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        //ExEnd
    }
}

