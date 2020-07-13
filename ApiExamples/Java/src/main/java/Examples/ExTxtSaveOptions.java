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
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.nio.charset.Charset;

@Test
public class ExTxtSaveOptions extends ApiExampleBase {
    @Test(dataProvider = "pageBreaksDataProvider")
    public void pageBreaks(boolean forcePageBreaks) throws Exception {
        //ExStart
        //ExFor:TxtSaveOptionsBase.ForcePageBreaks
        //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3");

        // If ForcePageBreaks is set to true then the output document will have form feed characters in place of page breaks
        // Otherwise, they will be line breaks
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        {
            saveOptions.setForcePageBreaks(forcePageBreaks);
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.PageBreaks.txt", saveOptions);

        // If we load the document using Aspose.Words again, the page breaks will be preserved/lost depending on ForcePageBreaks
        doc = new Document(getArtifactsDir() + "TxtSaveOptions.PageBreaks.txt");

        Assert.assertEquals(forcePageBreaks ? 3 : 1, doc.getPageCount());
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "pageBreaksDataProvider")
    public static Object[][] pageBreaksDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "addBidiMarksDataProvider")
    public void addBidiMarks(boolean addBidiMarks) throws Exception {
        //ExStart
        //ExFor:TxtSaveOptions.AddBidiMarks
        //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.getParagraphFormat().setBidi(true);
        builder.writeln("שלום עולם!");
        builder.writeln("مرحبا بالعالم!");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        {
            saveOptions.setAddBidiMarks(addBidiMarks);
            saveOptions.setEncoding(Charset.defaultCharset());
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.AddBidiMarks.txt", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "addBidiMarksDataProvider")
    public static Object[][] addBidiMarksDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "exportHeadersFootersDataProvider")
    public void exportHeadersFooters(/*TxtExportHeadersFootersMode*/int txtExportHeadersFootersMode) throws Exception {
        //ExStart
        //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
        //ExFor:TxtExportHeadersFootersMode
        //ExSummary:Shows how to specifies the way headers and footers are exported to plain text format.
        Document doc = new Document();

        // Insert even and primary headers/footers into the document
        // The primary header/footers should override the even ones 
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.HEADER_EVEN));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN).appendParagraph("Even header");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.FOOTER_EVEN));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN).appendParagraph("Even footer");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.HEADER_PRIMARY));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).appendParagraph("Primary header");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.FOOTER_PRIMARY));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).appendParagraph("Primary footer");

        // Insert pages that would display these headers and footers
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.write("Page 3");

        // Three values are available in TxtExportHeadersFootersMode enum:
        // "None" - No headers and footers are exported
        // "AllAtEnd" - All headers and footers are placed after all section bodies at the very end of a document
        // "PrimaryOnly" - Only primary headers and footers are exported at the beginning and end of each section (default value)
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        {
            saveOptions.setExportHeadersFootersMode(txtExportHeadersFootersMode);
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportHeadersFootersDataProvider")
    public static Object[][] exportHeadersFootersDataProvider() throws Exception {
        return new Object[][]
                {
                        {TxtExportHeadersFootersMode.ALL_AT_END},
                        {TxtExportHeadersFootersMode.PRIMARY_ONLY},
                        {TxtExportHeadersFootersMode.NONE},
                };
    }

    @Test
    public void txtListIndentation() throws Exception {
        //ExStart
        //ExFor:TxtListIndentation
        //ExFor:TxtListIndentation.Count
        //ExFor:TxtListIndentation.Character
        //ExFor:TxtSaveOptions.ListIndentation
        //ExSummary:Shows how to configure list indenting when converting to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list with three levels of indentation
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent();
        builder.write("Item 3");

        // Microsoft Word list objects get lost when converting to plaintext
        // We can create a custom representation for list indentation using pure plaintext with a SaveOptions object
        // In this case, each list item will be left-padded by 3 space characters times its list indent level
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        txtSaveOptions.getListIndentation().setCount(3);
        txtSaveOptions.getListIndentation().setCharacter(' ');

        doc.save(getArtifactsDir() + "TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);
        //ExEnd
    }

    @Test(dataProvider = "simplifyListLabelsDataProvider")
    public void simplifyListLabels(boolean simplifyListLabels) throws Exception {
        //ExStart
        //ExFor:TxtSaveOptions.SimplifyListLabels
        //ExSummary:Shows how to change the appearance of lists when converting to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list with five levels of indentation
        builder.getListFormat().applyBulletDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent();
        builder.writeln("Item 3");
        builder.getListFormat().listIndent();
        builder.writeln("Item 4");
        builder.getListFormat().listIndent();
        builder.write("Item 5");

        // The SimplifyListLabels flag will convert some list symbols
        // into ASCII characters such as *, o, +, > etc, depending on list level
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        {
            txtSaveOptions.setSimplifyListLabels(simplifyListLabels);
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.SimplifyListLabels.txt", txtSaveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "simplifyListLabelsDataProvider")
    public static Object[][] simplifyListLabelsDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void paragraphBreak() throws Exception {
        //ExStart
        //ExFor:TxtSaveOptions
        //ExFor:TxtSaveOptions.SaveFormat
        //ExFor:TxtSaveOptionsBase
        //ExFor:TxtSaveOptionsBase.ParagraphBreak
        //ExSummary:Shows how to save a .txt document with a custom paragraph break.
        // Create a new document and add some paragraphs
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Paragraph 1.");
        builder.writeln("Paragraph 2.");
        builder.write("Paragraph 3.");

        // When saved to plain text, the paragraphs we created can be separated by a custom string
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        {
            txtSaveOptions.setSaveFormat(SaveFormat.TEXT);
            txtSaveOptions.setParagraphBreak(" End of paragraph.\n\n\t");
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.ParagraphBreak.txt", txtSaveOptions);
        //ExEnd
    }

    @Test
    public void encoding() throws Exception {
        //ExStart
        //ExFor:TxtSaveOptionsBase.Encoding
        //ExSummary:Shows how to set encoding for a .txt output document.
        // Create a new document and add some text from outside the ASCII character set
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("À È Ì Ò Ù.");

        // We can use a SaveOptions object to make sure the encoding we save the .txt document in supports our content
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        {
            txtSaveOptions.setEncoding(Charset.forName("UTF-8"));
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.Encoding.txt", txtSaveOptions);
        //ExEnd
    }

    @Test(dataProvider = "tableLayoutDataProvider")
    public void tableLayout(boolean preserveTableLayout) throws Exception {
        //ExStart
        //ExFor:TxtSaveOptions.PreserveTableLayout
        //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table
        builder.startTable();
        builder.insertCell();
        builder.write("Row 1, cell 1");
        builder.insertCell();
        builder.write("Row 1, cell 2");
        builder.endRow();
        builder.insertCell();
        builder.write("Row 2, cell 1");
        builder.insertCell();
        builder.write("Row 2, cell 2");
        builder.endTable();

        // Tables, with their borders and widths do not translate to plaintext
        // However, we can configure a SaveOptions object to arrange table contents to preserve some of the table's appearance
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        {
            txtSaveOptions.setPreserveTableLayout(preserveTableLayout);
        }

        doc.save(getArtifactsDir() + "TxtSaveOptions.TableLayout.txt", txtSaveOptions);
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "tableLayoutDataProvider")
    public static Object[][] tableLayoutDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }
}
