// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.StyleIdentifier;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.XpsSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.Section;
import com.aspose.words.MultiplePagesType;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.words.BreakType;
import com.aspose.words.PageSet;
import com.aspose.words.CertificateHolder;
import com.aspose.words.SignOptions;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.DigitalSignatureDetails;
import org.testng.annotations.DataProvider;


@Test
public class ExXpsSaveOptions extends ApiExampleBase
{
    @Test
    public void outlineLevels() throws Exception
    {
        //ExStart
        //ExFor:XpsSaveOptions
        //ExFor:XpsSaveOptions.#ctor
        //ExFor:XpsSaveOptions.OutlineOptions
        //ExFor:XpsSaveOptions.SaveFormat
        //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved XPS document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

        Assert.assertTrue(builder.getParagraphFormat().isHeading());

        builder.writeln("Heading 1");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

        builder.writeln("Heading 1.1");
        builder.writeln("Heading 1.2");

        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);

        builder.writeln("Heading 1.2.1");
        builder.writeln("Heading 1.2.2");

        // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        Assert.assertEquals(SaveFormat.XPS, saveOptions.getSaveFormat());

        // The output XPS document will contain an outline, a table of contents that lists headings in the document body.
        // Clicking on an entry in this outline will take us to the location of its respective heading.
        // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
        // The last two headings we have inserted above will not appear.
        saveOptions.getOutlineOptions().setHeadingsOutlineLevels(2);

        doc.save(getArtifactsDir() + "XpsSaveOptions.OutlineLevels.xps", saveOptions);
        //ExEnd
    }

    @Test (dataProvider = "bookFoldDataProvider")
    public void bookFold(boolean renderTextAsBookFold) throws Exception
    {
        //ExStart
        //ExFor:XpsSaveOptions.#ctor(SaveFormat)
        //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.XPS);

        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output XPS in a way that helps us use it to make a booklet.
        // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
        xpsOptions.setUseBookFoldPrintingSettings(renderTextAsBookFold);

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        if (renderTextAsBookFold)
            for (Section s : (Iterable<Section>) doc.getSections())
            {
                s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
            }

        // Once we print this document, we can turn it into a booklet by stacking the pages
        // to come out of the printer and folding down the middle.
        doc.save(getArtifactsDir() + "XpsSaveOptions.BookFold.xps", xpsOptions);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "bookFoldDataProvider")
	public static Object[][] bookFoldDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "optimizeOutputDataProvider")
    public void optimizeOutput(boolean optimizeOutput) throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to xps.
        Document doc = new Document(getMyDir() + "Unoptimized document.docx");

        // Create an "XpsSaveOptions" object to pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        XpsSaveOptions saveOptions = new XpsSaveOptions();
        // Set the "OptimizeOutput" property to "true" to take measures such as removing nested or empty canvases
        // and concatenating adjacent runs with identical formatting to optimize the output document's content.
        // This may affect the appearance of the document.
        // Set the "OptimizeOutput" property to "false" to save the document normally.
        saveOptions.setOptimizeOutput(optimizeOutput);

        doc.save(getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps", saveOptions);
        //ExEnd

        long testedFileLength = new FileInfo(getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps").getLength();
        if (optimizeOutput)
            Assert.assertTrue(testedFileLength < 43000);
        else
            Assert.assertTrue(testedFileLength < 64000);

        TestUtil.docPackageFileContainsString(
            optimizeOutput
                ? "Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" " +
                  "UnicodeString=\"This document contains complex content which can be optimized to save space when \""
                : "<Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" UnicodeString=\"This\"",
            getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps", "1.fpage");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "optimizeOutputDataProvider")
	public static Object[][] optimizeOutputDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void exportExactPages() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.PageSet
        //ExFor:PageSet.#ctor(int[])
        //ExSummary:Shows how to extract pages based on exact page indices.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add five pages to the document.
        for (int i = 1; i < 6; i++)
        {
            builder.write("Page " + i);
            builder.insertBreak(BreakType.PAGE_BREAK);
        }

        // Create an "XpsSaveOptions" object, which we can pass to the document's "Save" method
        // to modify how that method converts the document to .XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Use the "PageSet" property to select a set of the document's pages to save to output XPS.
        // In this case, we will choose, via a zero-based index, only three pages: page 1, page 2, and page 4.
        xpsOptions.setPageSet(new PageSet(0, 1, 3));

        doc.save(getArtifactsDir() + "XpsSaveOptions.ExportExactPages.xps", xpsOptions);
        //ExEnd
    }

    @Test
    public void xpsDigitalSignature() throws Exception
    {
        //ExStart:XpsDigitalSignature
        //GistId:708ce40a68fac5003d46f6b4acfd5ff1
        //ExFor:XpsSaveOptions.DigitalSignatureDetails
        //ExSummary:Shows how to sign XPS document.
        Document doc = new Document(getMyDir() + "Document.docx");

        CertificateHolder certificateHolder = CertificateHolder.create(getMyDir() + "morzal.pfx", "aw");
        SignOptions options = new SignOptions();
        options.setSignTimeInternal(new Date);
        options.setComments("Some comments");

        DigitalSignatureDetails digitalSignatureDetails = new DigitalSignatureDetails(certificateHolder, options);

        XpsSaveOptions saveOptions = new XpsSaveOptions();
        saveOptions.setDigitalSignatureDetails(digitalSignatureDetails);

        Assert.assertEquals(certificateHolder, digitalSignatureDetails.getCertificateHolder());
        Assert.assertEquals("Some comments", digitalSignatureDetails.getSignOptions().getComments());

        doc.save(getArtifactsDir() + "XpsSaveOptions.XpsDigitalSignature.docx", saveOptions);
        //ExEnd:XpsDigitalSignature
    }
}
