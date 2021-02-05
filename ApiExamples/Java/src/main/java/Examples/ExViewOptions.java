package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;

@Test
public class ExViewOptions extends ApiExampleBase {
    @Test
    public void setZoom() throws Exception {
        //ExStart
        //ExFor:Document.ViewOptions
        //ExFor:ViewOptions
        //ExFor:ViewOptions.ViewType
        //ExFor:ViewOptions.ZoomPercent
        //ExFor:ViewOptions.ZoomType
        //ExFor:ViewType
        //ExSummary:Shows how to set a custom zoom factor, which older versions of Microsoft Word will apply to a document upon loading.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        doc.getViewOptions().setViewType(ViewType.PAGE_LAYOUT);
        doc.getViewOptions().setZoomPercent(50);

        Assert.assertEquals(ZoomType.CUSTOM, doc.getViewOptions().getZoomType());
        Assert.assertEquals(ZoomType.NONE, doc.getViewOptions().getZoomType());

        doc.save(getArtifactsDir() + "ViewOptions.SetZoomPercentage.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.SetZoomPercentage.doc");

        Assert.assertEquals(ViewType.PAGE_LAYOUT, doc.getViewOptions().getViewType());
        Assert.assertEquals(50, doc.getViewOptions().getZoomPercent());
        Assert.assertEquals(ZoomType.NONE, doc.getViewOptions().getZoomType());
    }

    @Test(dataProvider = "setZoomTypeDataProvider")
    public void setZoomType(int zoomType) throws Exception {
        //ExStart
        //ExFor:Document.ViewOptions
        //ExFor:ViewOptions
        //ExFor:ViewOptions.ZoomType
        //ExSummary:Shows how to set a custom zoom type, which older versions of Microsoft Word will apply to a document upon loading.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Set the "ZoomType" property to "ZoomType.PageWidth" to get Microsoft Word
        // to automatically zoom the document to fit the width of the page.
        // Set the "ZoomType" property to "ZoomType.FullPage" to get Microsoft Word
        // to automatically zoom the document to make the entire first page visible.
        // Set the "ZoomType" property to "ZoomType.TextFit" to get Microsoft Word
        // to automatically zoom the document to fit the inner text margins of the first page.
        doc.getViewOptions().setZoomType(zoomType);

        doc.save(getArtifactsDir() + "ViewOptions.SetZoomType.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.SetZoomType.doc");

        Assert.assertEquals(zoomType, doc.getViewOptions().getZoomType());
    }

    @DataProvider(name = "setZoomTypeDataProvider")
    public static Object[][] setZoomTypeDataProvider() {
        return new Object[][]
                {
                        {ZoomType.PAGE_WIDTH},
                        {ZoomType.FULL_PAGE},
                        {ZoomType.TEXT_FIT},
                };
    }

    @Test(dataProvider = "displayBackgroundShapeDataProvider")
    public void displayBackgroundShape(boolean displayBackgroundShape) throws Exception {
        //ExStart
        //ExFor:ViewOptions.DisplayBackgroundShape
        //ExSummary:Shows how to hide/display document background images in view options.
        // Use an HTML string to create a new document with a flat background color.
        final String HTML =
                "<html>\r\n                <body style='background-color: blue'>\r\n                    <p>Hello world!</p>\r\n                </body>\r\n            </html>";

        Document doc = new Document(new ByteArrayInputStream(HTML.getBytes()));

        // The source for the document has a flat color background,
        // the presence of which will set the "DisplayBackgroundShape" flag to "true".
        Assert.assertTrue(doc.getViewOptions().getDisplayBackgroundShape());

        // Keep the "DisplayBackgroundShape" as "true" to get the document to display the background color.
        // This may affect some text colors to improve visibility.
        // Set the "DisplayBackgroundShape" to "false" to not display the background color.
        doc.getViewOptions().setDisplayBackgroundShape(displayBackgroundShape);

        doc.save(getArtifactsDir() + "ViewOptions.DisplayBackgroundShape.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.DisplayBackgroundShape.docx");

        Assert.assertEquals(displayBackgroundShape, doc.getViewOptions().getDisplayBackgroundShape());
    }

    @DataProvider(name = "displayBackgroundShapeDataProvider")
    public static Object[][] displayBackgroundShapeDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "displayPageBoundariesDataProvider")
    public void displayPageBoundaries(boolean doNotDisplayPageBoundaries) throws Exception {
        //ExStart
        //ExFor:ViewOptions.DoNotDisplayPageBoundaries
        //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert content that spans across 3 pages.
        builder.writeln("Paragraph 1, Page 1.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Paragraph 2, Page 2.");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Paragraph 3, Page 3.");

        // Insert a header and a footer.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("This is the header.");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.writeln("This is the footer.");

        // This document contains a small amount of content that takes up a few full pages worth of space.
        // Set the "DoNotDisplayPageBoundaries" flag to "true" to get older versions of Microsoft Word to omit headers,
        // footers, and much of the vertical whitespace when displaying our document.
        // Set the "DoNotDisplayPageBoundaries" flag to "false" to get older versions of Microsoft Word
        // to normally display our document.
        doc.getViewOptions().setDoNotDisplayPageBoundaries(doNotDisplayPageBoundaries);

        doc.save(getArtifactsDir() + "ViewOptions.DisplayPageBoundaries.doc");
        //ExEnd

        doc = new Document(getArtifactsDir() + "ViewOptions.DisplayPageBoundaries.doc");

        Assert.assertEquals(doNotDisplayPageBoundaries, doc.getViewOptions().getDoNotDisplayPageBoundaries());
    }

    @DataProvider(name = "displayPageBoundariesDataProvider")
    public static Object[][] displayPageBoundariesDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "formsDesignDataProvider")
    public void formsDesign(boolean useFormsDesign) throws Exception {
        //ExStart
        //ExFor:ViewOptions.FormsDesign
        //ExSummary:Shows how to enable/disable forms design mode.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Set the "FormsDesign" property to "false" to keep forms design mode disabled.
        // Set the "FormsDesign" property to "true" to enable forms design mode.
        doc.getViewOptions().setFormsDesign(useFormsDesign);

        doc.save(getArtifactsDir() + "ViewOptions.FormsDesign.xml");
        //ExEnd
    }

    @DataProvider(name = "formsDesignDataProvider")
    public static Object[][] formsDesignDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }
}
