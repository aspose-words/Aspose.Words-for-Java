// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ViewType;
import org.testng.Assert;
import com.aspose.words.ZoomType;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.WordML2003SaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.File;
import org.testng.annotations.DataProvider;


@Test
public class ExViewOptions extends ApiExampleBase
{
    @Test
    public void setZoom() throws Exception
    {
        //ExStart
        //ExFor:Document.ViewOptions
        //ExFor:ViewOptions
        //ExFor:ViewOptions.ViewType
        //ExFor:ViewOptions.ZoomType
        //ExFor:ViewOptions.ZoomPercent
        //ExFor:ViewType
        //ExSummary:Shows how to make sure the document is displayed at 50% zoom when opened in Microsoft Word.
        Document doc = new Document(getMyDir() + "Document.docx");

        // We can set the zoom factor to a percentage
        doc.getViewOptions().setViewType(ViewType.PAGE_LAYOUT);
        doc.getViewOptions().setZoomPercent(50);

        // Or we can set the ZoomType to a different value to avoid using percentages 
        Assert.assertEquals(ZoomType.NONE, doc.getViewOptions().getZoomType());

        doc.save(getArtifactsDir() + "ViewOptions.SetZoom.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.SetZoom.docx");

        Assert.assertEquals(ViewType.PAGE_LAYOUT, doc.getViewOptions().getViewType());
        Assert.assertEquals(50.0d, doc.getViewOptions().getZoomPercent());
        Assert.assertEquals(ZoomType.NONE, doc.getViewOptions().getZoomType());
    }

    @Test
    public void displayBackgroundShape() throws Exception
    {
        //ExStart
        //ExFor:ViewOptions.DisplayBackgroundShape
        //ExSummary:Shows how to hide/display document background images in view options.
        // Create a new document from an html string with a flat background color
        final String HTML = "<html>\r\n                <body style='background-color: blue'>\r\n                    <p>Hello world!</p>\r\n                </body>\r\n            </html>";

        Document doc = new Document(new MemoryStream(Encoding.getUnicode().getBytes(HTML)));

        // The source for the document has a flat color background, the presence of which will turn on the DisplayBackgroundShape flag
        // We can disable it like this
        doc.getViewOptions().setDisplayBackgroundShape(false);

        doc.save(getArtifactsDir() + "ViewOptions.DisplayBackgroundShape.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.DisplayBackgroundShape.docx");

        Assert.assertFalse(doc.getViewOptions().getDisplayBackgroundShape());
    }

    @Test
    public void displayPageBoundaries() throws Exception
    {
        //ExStart
        //ExFor:ViewOptions.DoNotDisplayPageBoundaries
        //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert content spanning 3 pages
        builder.writeln("Paragraph 1, Page 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Paragraph 2, Page 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Paragraph 3, Page 3");

        // Insert a header and a footer
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Header");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.writeln("Footer");

        // In this case we have a lot of space taken up by quite a little amount of content
        // In older versions of Microsoft Word, we can hide headers/footers and compact vertical whitespace of pages
        // to give the document's main body content some flow by setting this flag
        doc.getViewOptions().setDoNotDisplayPageBoundaries(true);

        doc.save(getArtifactsDir() + "ViewOptions.DisplayPageBoundaries.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.DisplayPageBoundaries.docx");

        Assert.assertTrue(doc.getViewOptions().getDoNotDisplayPageBoundaries());
    }

    @Test (dataProvider = "formsDesignDataProvider")
    public void formsDesign(boolean useFormsDesign) throws Exception
    {
        //ExStart
        //ExFor:ViewOptions.FormsDesign
        //ExFor:WordML2003SaveOptions
        //ExFor:WordML2003SaveOptions.SaveFormat
        //ExSummary:Shows how to save to a .wml document while applying save options.
        Document doc = new Document(getMyDir() + "Document.docx");

        WordML2003SaveOptions options = new WordML2003SaveOptions();
        {
            options.setSaveFormat(SaveFormat.WORD_ML);
            options.setMemoryOptimization(true);
            options.setPrettyFormat(true);
        }

        // Enables forms design mode in WordML documents
        doc.getViewOptions().setFormsDesign(useFormsDesign);

        doc.save(getArtifactsDir() + "ViewOptions.FormsDesign.xml", options);

        Assert.assertEquals(useFormsDesign,
            File.readAllText(getArtifactsDir() + "ViewOptions.FormsDesign.xml").contains("<w:formsDesign />"));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "formsDesignDataProvider")
	public static Object[][] formsDesignDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}
