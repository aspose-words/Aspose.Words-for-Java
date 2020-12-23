package Examples;

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
        //ExSummary:Shows how to limit the level of headings that will appear in the outline of a saved XPS document.
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

        // Create an "XpsSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method converts the document to .XPS.
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        Assert.assertEquals(SaveFormat.XPS, saveOptions.getSaveFormat());

        // The output XPS document will contain an outline, which is a table of contents that lists headings in the document body.
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

        // Create an "XpsSaveOptions" object which we can pass to the document's "Save" method
        // to modify the way in which that method converts the document to .XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.XPS);

        // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
        // in the output XPS in a way that helps us use it to make a booklet.
        // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
        xpsOptions.setUseBookFoldPrintingSettings(renderTextAsBookFold);

        // If we are rendering the document as a booklet, we must set the "MultiplePages"
        // properties of all page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
        if (renderTextAsBookFold)
            for (Section s : (Iterable<Section>) doc.getSections())
            {
                s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
            }

        // Once we print this document, we can turn it into a booklet by stacking the pages
        // in the order they come out of the printer and then folding down the middle
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

        // When saving to .xps, we can use SaveOptions to optimize the output in some cases
        XpsSaveOptions saveOptions = new XpsSaveOptions(); { saveOptions.setOptimizeOutput(optimizeOutput); }

        doc.save(getArtifactsDir() + "XpsSaveOptions.OptimizeOutput.xps", saveOptions);
        //ExEnd
    }

    @DataProvider(name = "optimizeOutputDataProvider")
    public static Object[][] optimizeOutputDataProvider() throws Exception {
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
        Document doc = new Document(getMyDir() + "Images.docx");

        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        xpsOptions.setPageSet(new PageSet(0, 1, 2, 4, 1, 3, 2, 3));

        doc.save(getArtifactsDir() + "XpsSaveOptions.ExportExactPages.xps", xpsOptions);
        //ExEnd
    }
}
