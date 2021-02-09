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
import com.aspose.words.BreakType;
import com.aspose.words.TxtSaveOptions;
import org.testng.Assert;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.ms.System.IO.File;
import com.aspose.words.TxtExportHeadersFootersMode;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.SaveFormat;
import org.testng.annotations.DataProvider;


@Test
public class ExTxtSaveOptions extends ApiExampleBase
{
    @Test (dataProvider = "pageBreaksDataProvider")
    public void pageBreaks(boolean forcePageBreaks) throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptionsBase.ForcePageBreaks
        //ExSummary:Shows how to specify whether to preserve page breaks when exporting a document to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 3");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save"
        // method to modify how we save the document to plaintext.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // The Aspose.Words "Document" objects have page breaks, just like Microsoft Word documents.
        // Save formats such as ".txt" are one continuous body of text without page breaks.
        // Set the "ForcePageBreaks" property to "true" to preserve all page breaks in the form of '\f' characters.
        // Set the "ForcePageBreaks" property to "false" to discard all page breaks.
        saveOptions.setForcePageBreaks(forcePageBreaks);

        doc.save(getArtifactsDir() + "TxtSaveOptions.PageBreaks.txt", saveOptions);
        
        // If we load a plaintext document with page breaks,
        // the "Document" object will use them to split the body into pages.
        doc = new Document(getArtifactsDir() + "TxtSaveOptions.PageBreaks.txt");

        Assert.assertEquals(forcePageBreaks ? 3 : 1, doc.getPageCount());
        //ExEnd

        TestUtil.fileContainsString(
            forcePageBreaks ? "Page 1\r\n\fPage 2\r\n\fPage 3\r\n\r\n" : "Page 1\r\nPage 2\r\nPage 3\r\n\r\n",
            getArtifactsDir() + "TxtSaveOptions.PageBreaks.txt");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "pageBreaksDataProvider")
	public static Object[][] pageBreaksDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "addBidiMarksDataProvider")
    public void addBidiMarks(boolean addBidiMarks) throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptions.AddBidiMarks
        //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.getParagraphFormat().setBidi(true);
        builder.writeln("שלום עולם!");
        builder.writeln("مرحبا بالعالم!");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions saveOptions = new TxtSaveOptions(); { saveOptions.setEncoding(com.aspose.ms.System.Text.Encoding.getUnicode());}

        // Set the "AddBidiMarks" property to "true" to add marks before runs
        // with right-to-left text to indicate the fact.
        // Set the "AddBidiMarks" property to "false" to write all left-to-right
        // and right-to-left run equally with nothing to indicate which is which.
        saveOptions.setAddBidiMarks(addBidiMarks);

        doc.save(getArtifactsDir() + "TxtSaveOptions.AddBidiMarks.txt", saveOptions);

        String docText = com.aspose.ms.System.Text.Encoding.getUnicode().getString(File.readAllBytes(getArtifactsDir() + "TxtSaveOptions.AddBidiMarks.txt"));

        if (addBidiMarks)
        {
            Assert.assertEquals("\uFEFFHello world!‎\r\nשלום עולם!‏\r\nمرحبا بالعالم!‏\r\n\r\n", docText);
            Assert.assertTrue(docText.contains("\u200f"));
        }
        else
        {
            Assert.assertEquals("\uFEFFHello world!\r\nשלום עולם!\r\nمرحبا بالعالم!\r\n\r\n", docText);
            Assert.assertFalse(docText.contains("\u200f"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "addBidiMarksDataProvider")
	public static Object[][] addBidiMarksDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "exportHeadersFootersDataProvider")
    public void exportHeadersFooters(/*TxtExportHeadersFootersMode*/int txtExportHeadersFootersMode) throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
        //ExFor:TxtExportHeadersFootersMode
        //ExSummary:Shows how to specify how to export headers and footers to plain text format.
        Document doc = new Document();

        // Insert even and primary headers/footers into the document.
        // The primary header/footers will override the even headers/footers.
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.HEADER_EVEN));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN).appendParagraph("Even header");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.FOOTER_EVEN));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN).appendParagraph("Even footer");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.HEADER_PRIMARY));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).appendParagraph("Primary header");
        doc.getFirstSection().getHeadersFooters().add(new HeaderFooter(doc, HeaderFooterType.FOOTER_PRIMARY));
        doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).appendParagraph("Primary footer");

        // Insert pages to display these headers and footers.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK); 
        builder.write("Page 3");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.None"
        // to not export any headers/footers.
        // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.PrimaryOnly"
        // to only export primary headers/footers.
        // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.AllAtEnd"
        // to place all headers and footers for all section bodies at the end of the document.
        saveOptions.setExportHeadersFootersMode(txtExportHeadersFootersMode);

        doc.save(getArtifactsDir() + "TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);

        String docText = File.readAllText(getArtifactsDir() + "TxtSaveOptions.ExportHeadersFooters.txt");

        switch (txtExportHeadersFootersMode)
        {
            case TxtExportHeadersFootersMode.ALL_AT_END:
                Assert.assertEquals("Page 1\r\n" +
                                "Page 2\r\n" +
                                "Page 3\r\n" +
                                "Even header\r\n\r\n" +
                                "Primary header\r\n\r\n" +
                                "Even footer\r\n\r\n" +
                                "Primary footer\r\n\r\n", docText);
                break;
            case TxtExportHeadersFootersMode.PRIMARY_ONLY:
                Assert.assertEquals("Primary header\r\n" +
                                "Page 1\r\n" +
                                "Page 2\r\n" +
                                "Page 3\r\n" +
                                "Primary footer\r\n", docText);
                break;
            case TxtExportHeadersFootersMode.NONE:
                Assert.assertEquals("Page 1\r\n" +
                                "Page 2\r\n" +
                                "Page 3\r\n", docText);
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportHeadersFootersDataProvider")
	public static Object[][] exportHeadersFootersDataProvider() throws Exception
	{
		return new Object[][]
		{
			{TxtExportHeadersFootersMode.ALL_AT_END},
			{TxtExportHeadersFootersMode.PRIMARY_ONLY},
			{TxtExportHeadersFootersMode.NONE},
		};
	}

    @Test
    public void txtListIndentation() throws Exception
    {
        //ExStart
        //ExFor:TxtListIndentation
        //ExFor:TxtListIndentation.Count
        //ExFor:TxtListIndentation.Character
        //ExFor:TxtSaveOptions.ListIndentation
        //ExSummary:Shows how to configure list indenting when saving a document to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list with three levels of indentation.
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent(); 
        builder.write("Item 3");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

        // Set the "Character" property to assign a character to use
        // for padding that simulates list indentation in plaintext.
        txtSaveOptions.getListIndentation().setCharacter(' ');

        // Set the "Count" property to specify the number of times
        // to place the padding character for each list indent level.
        txtSaveOptions.getListIndentation().setCount(3);

        doc.save(getArtifactsDir() + "TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);

        String docText = File.readAllText(getArtifactsDir() + "TxtSaveOptions.TxtListIndentation.txt");

        Assert.assertEquals("1. Item 1\r\n" +
                        "   a. Item 2\r\n" +
                        "      i. Item 3\r\n", docText);
        //ExEnd
    }

    @Test (dataProvider = "simplifyListLabelsDataProvider")
    public void simplifyListLabels(boolean simplifyListLabels) throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptions.SimplifyListLabels
        //ExSummary:Shows how to change the appearance of lists when saving a document to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list with five levels of indentation.
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

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

        // Set the "SimplifyListLabels" property to "true" to convert some list
        // symbols into simpler ASCII characters, such as '*', 'o', '+', '>', etc.
        // Set the "SimplifyListLabels" property to "false" to preserve as many original list symbols as possible.
        txtSaveOptions.setSimplifyListLabels(simplifyListLabels);

        doc.save(getArtifactsDir() + "TxtSaveOptions.SimplifyListLabels.txt", txtSaveOptions);

        String docText = File.readAllText(getArtifactsDir() + "TxtSaveOptions.SimplifyListLabels.txt");

        if (simplifyListLabels)
            Assert.assertEquals("* Item 1\r\n" +
                            "  > Item 2\r\n" +
                            "    + Item 3\r\n" +
                            "      - Item 4\r\n" +
                            "        o Item 5\r\n", docText);
        else
            Assert.assertEquals("· Item 1\r\n" +
                            "o Item 2\r\n" +
                            "§ Item 3\r\n" +
                            "· Item 4\r\n" +
                            "o Item 5\r\n", docText);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "simplifyListLabelsDataProvider")
	public static Object[][] simplifyListLabelsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void paragraphBreak() throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptions
        //ExFor:TxtSaveOptions.SaveFormat
        //ExFor:TxtSaveOptionsBase
        //ExFor:TxtSaveOptionsBase.ParagraphBreak
        //ExSummary:Shows how to save a .txt document with a custom paragraph break.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Paragraph 1.");
        builder.writeln("Paragraph 2.");
        builder.write("Paragraph 3.");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

        Assert.assertEquals(SaveFormat.TEXT, txtSaveOptions.getSaveFormat());

        // Set the "ParagraphBreak" to a custom value that we wish to put at the end of every paragraph.
        txtSaveOptions.setParagraphBreak(" End of paragraph.\n\n\t");

        doc.save(getArtifactsDir() + "TxtSaveOptions.ParagraphBreak.txt", txtSaveOptions);

        String docText = File.readAllText(getArtifactsDir() + "TxtSaveOptions.ParagraphBreak.txt");

        Assert.assertEquals("Paragraph 1. End of paragraph.\n\n\t" +
                        "Paragraph 2. End of paragraph.\n\n\t" +
                        "Paragraph 3. End of paragraph.\n\n\t", docText);
        //ExEnd
    }

    @Test
    public void encoding() throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptionsBase.Encoding
        //ExSummary:Shows how to set encoding for a .txt output document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text with characters from outside the ASCII character set.
        builder.write("À È Ì Ò Ù.");

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
        
        // Verify that the "Encoding" property contains the appropriate encoding for our document's contents.
        Assert.assertEquals(com.aspose.ms.System.Text.Encoding.getUTF8(), txtSaveOptions.getEncodingInternal());

        doc.save(getArtifactsDir() + "TxtSaveOptions.Encoding.UTF8.txt", txtSaveOptions);

        String docText = com.aspose.ms.System.Text.Encoding.getUTF8().getString(File.readAllBytes(getArtifactsDir() + "TxtSaveOptions.Encoding.UTF8.txt"));
        
        Assert.assertEquals("\uFEFFÀ È Ì Ò Ù.\r\n", docText);

        // Using an unsuitable encoding may result in a loss of document contents.
        txtSaveOptions.setEncodingInternal(com.aspose.ms.System.Text.Encoding.getASCII());
        doc.save(getArtifactsDir() + "TxtSaveOptions.Encoding.ASCII.txt", txtSaveOptions);
        docText = com.aspose.ms.System.Text.Encoding.getASCII().getString(File.readAllBytes(getArtifactsDir() + "TxtSaveOptions.Encoding.ASCII.txt"));

        Assert.assertEquals("? ? ? ? ?.\r\n", docText);
        //ExEnd
    }

    @Test (dataProvider = "preserveTableLayoutDataProvider")
    public void preserveTableLayout(boolean preserveTableLayout) throws Exception
    {
        //ExStart
        //ExFor:TxtSaveOptions.PreserveTableLayout
        //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

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

        // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how we save the document to plaintext.
        TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

        // Set the "PreserveTableLayout" property to "true" to apply whitespace padding to the contents
        // of the output plaintext document to preserve as much of the table's layout as possible.
        // Set the "PreserveTableLayout" property to "false" to save all tables' contents
        // as a continuous body of text, with just a new line for each row.
        txtSaveOptions.setPreserveTableLayout(preserveTableLayout);

        doc.save(getArtifactsDir() + "TxtSaveOptions.PreserveTableLayout.txt", txtSaveOptions);

        String docText = File.readAllText(getArtifactsDir() + "TxtSaveOptions.PreserveTableLayout.txt");

        if (preserveTableLayout)
            Assert.assertEquals("Row 1, cell 1                Row 1, cell 2\r\n" +
                            "Row 2, cell 1                Row 2, cell 2\r\n\r\n", docText);
        else
            Assert.assertEquals("Row 1, cell 1\r\n" +
                            "Row 1, cell 2\r\n" +
                            "Row 2, cell 1\r\n" +
                            "Row 2, cell 2\r\n\r\n", docText);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "preserveTableLayoutDataProvider")
	public static Object[][] preserveTableLayoutDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
