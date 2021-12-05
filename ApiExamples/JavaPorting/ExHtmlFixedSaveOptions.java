// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HtmlFixedSaveOptions;
import com.aspose.ms.System.Text.Encoding;
import org.testng.Assert;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.Directory;
import com.aspose.words.HtmlFixedPageHorizontalAlignment;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.words.ExportFontFormat;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IResourceSavingCallback;
import com.aspose.words.ResourceSavingArgs;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import org.testng.annotations.DataProvider;


@Test
class ExHtmlFixedSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void useEncoding() throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.Encoding
        //ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("Hello World!");

        // The default encoding is UTF-8. If we want to represent our document using a different encoding,
        // we can use a SaveOptions object to set a specific encoding.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setEncoding(Encoding.getEncoding("ASCII"));
        }

        Assert.assertEquals("US-ASCII", htmlFixedSaveOptions.getEncodingInternal().getEncodingName());

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
        //ExEnd

        Assert.assertTrue(Regex.match(File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.UseEncoding.html"), 
            "content=\"text/html; charset=us-ascii\"").getSuccess());
    }

    @Test
    public void getEncoding() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setEncoding(Encoding.getEncoding("utf-16"));
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.GetEncoding.html", htmlFixedSaveOptions);
    }

    @Test (dataProvider = "exportEmbeddedCssDataProvider")
    public void exportEmbeddedCss(boolean exportEmbeddedCss) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
        //ExSummary:Shows how to determine where to store CSS stylesheets when exporting a document to Html.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When we export a document to html, Aspose.Words will also create a CSS stylesheet to format the document with.
        // Setting the "ExportEmbeddedCss" flag to "true" save the CSS stylesheet to a .css file,
        // and link to the file from the html document using a <link> element.
        // Setting the flag to "false" will embed the CSS stylesheet within the Html document,
        // which will create only one file instead of two.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedCss(exportEmbeddedCss);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss.html");

        if (exportEmbeddedCss)
        {
            Assert.assertTrue(Regex.match(outDocContents, "<style type=\"text/css\">").getSuccess());
            Assert.assertFalse(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css"));
        }
        else
        {
            Assert.assertTrue(Regex.match(outDocContents,
                "<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCss/styles[.]css\" media=\"all\" />").getSuccess());
            Assert.assertTrue(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportEmbeddedCssDataProvider")
	public static Object[][] exportEmbeddedCssDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "exportEmbeddedFontsDataProvider")
    public void exportEmbeddedFonts(boolean exportEmbeddedFonts) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
        //ExSummary:Shows how to determine where to store embedded fonts when exporting a document to Html.
        Document doc = new Document(getMyDir() + "Embedded font.docx");

        // When we export a document with embedded fonts to .html,
        // Aspose.Words can place the fonts in two possible locations.
        // Setting the "ExportEmbeddedFonts" flag to "true" will store the raw data for embedded fonts within the CSS stylesheet,
        // in the "url" property of the "@font-face" rule. This may create a huge CSS stylesheet file
        // and reduce the number of external files that this HTML conversion will create.
        // Setting this flag to "false" will create a file for each font.
        // The CSS stylesheet will link to each font file using the "url" property of the "@font-face" rule.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedFonts(exportEmbeddedFonts);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

        if (exportEmbeddedFonts)
        {
            Assert.assertTrue(Regex.match(outDocContents,
                "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }").getSuccess());
            Assert.AreEqual(0, Directory.getFiles(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts").Count(f => f.EndsWith(".woff")));
        }
        else
        {
            Assert.assertTrue(Regex.match(outDocContents,
                "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }").getSuccess());
            Assert.AreEqual(2, Directory.getFiles(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts").Count(f => f.EndsWith(".woff")));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportEmbeddedFontsDataProvider")
	public static Object[][] exportEmbeddedFontsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "exportEmbeddedImagesDataProvider")
    public void exportEmbeddedImages(boolean exportImages) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
        //ExSummary:Shows how to determine where to store images when exporting a document to Html.
        Document doc = new Document(getMyDir() + "Images.docx");

        // When we export a document with embedded images to .html,
        // Aspose.Words can place the images in two possible locations.
        // Setting the "ExportEmbeddedImages" flag to "true" will store the raw data
        // for all images within the output HTML document, in the "src" attribute of <image> tags.
        // Setting this flag to "false" will create an image file in the local file system for every image,
        // and store all these files in a separate folder.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedImages(exportImages);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages.html");

        if (exportImages)
        {
            Assert.assertFalse(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
            Assert.assertTrue(Regex.match(outDocContents,
                "<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" src=\".+\" />").getSuccess());
        }
        else
        {
            Assert.assertTrue(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
            Assert.assertTrue(Regex.match(outDocContents,
                "<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" " +
                "src=\"HtmlFixedSaveOptions[.]ExportEmbeddedImages/image001[.]jpeg\" />").getSuccess());
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportEmbeddedImagesDataProvider")
	public static Object[][] exportEmbeddedImagesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "exportEmbeddedSvgsDataProvider")
    public void exportEmbeddedSvgs(boolean exportSvgs) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
        //ExSummary:Shows how to determine where to store SVG objects when exporting a document to Html.
        Document doc = new Document(getMyDir() + "Images.docx");

        // When we export a document with SVG objects to .html,
        // Aspose.Words can place these objects in two possible locations.
        // Setting the "ExportEmbeddedSvg" flag to "true" will embed all SVG object raw data
        // within the output HTML, inside <image> tags.
        // Setting this flag to "false" will create a file in the local file system for each SVG object.
        // The HTML will link to each file using the "data" attribute of an <object> tag.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedSvg(exportSvgs);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

        if (exportSvgs)
        {
            Assert.assertFalse(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
            Assert.assertTrue(Regex.match(outDocContents,
                "<image id=\"image004\" xlink:href=.+/>").getSuccess());
        }
        else
        {
            Assert.assertTrue(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
            Assert.assertTrue(Regex.match(outDocContents,
                "<object type=\"image/svg[+]xml\" data=\"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001[.]svg\"></object>").getSuccess());
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportEmbeddedSvgsDataProvider")
	public static Object[][] exportEmbeddedSvgsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "exportFormFieldsDataProvider")
    public void exportFormFields(boolean exportFormFields) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportFormFields
        //ExSummary:Shows how to export form fields to Html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCheckBox("CheckBox", false, 15);

        // When we export a document with form fields to .html,
        // there are two ways in which Aspose.Words can export form fields.
        // Setting the "ExportFormFields" flag to "true" will export them as interactive objects.
        // Setting this flag to "false" will display form fields as plain text.
        // This will freeze them at their current value, and prevent the reader of our HTML document
        // from being able to interact with them.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportFormFields(exportFormFields);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportFormFields.html");

        if (exportFormFields)
        {
            Assert.assertTrue(Regex.match(outDocContents,
                "<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
                "<input style=\"position:absolute; left:0pt; top:0pt;\" type=\"checkbox\" name=\"CheckBox\" />").getSuccess());
        }
        else
        {
            Assert.assertTrue(Regex.match(outDocContents, 
                "<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
                "<div class=\"awdiv\" style=\"left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;\"").getSuccess());
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportFormFieldsDataProvider")
	public static Object[][] exportFormFieldsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void addCssClassNamesPrefix() throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.CssClassNamesPrefix
        //ExFor:HtmlFixedSaveOptions.SaveFontFaceCssSeparately
        //ExSummary:Shows how to place CSS into a separate file and add a prefix to all of its CSS class names.
        Document doc = new Document(getMyDir() + "Bookmarks.docx");

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setCssClassNamesPrefix("myprefix");
            htmlFixedSaveOptions.setSaveFontFaceCssSeparately(true);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html");

        Assert.assertTrue(Regex.match(outDocContents,
            "<div class=\"myprefixdiv myprefixpage\" style=\"width:595[.]3pt; height:841[.]9pt;\">" +
            "<div class=\"myprefixdiv\" style=\"left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];\">" +
            "<span class=\"myprefixspan myprefixtext001\" style=\"font-size:11pt; left:294[.]73pt; top:0[.]36pt;\">").getSuccess());

        outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css");

        Assert.assertTrue(Regex.match(outDocContents,
            ".myprefixdiv { position:absolute; } " +
            ".myprefixspan { position:absolute; white-space:pre; color:#000000; font-size:12pt; }").getSuccess());
        //ExEnd
    }

    @Test (dataProvider = "horizontalAlignmentDataProvider")
    public void horizontalAlignment(/*HtmlFixedPageHorizontalAlignment*/int pageHorizontalAlignment) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
        //ExFor:HtmlFixedPageHorizontalAlignment
        //ExSummary:Shows how to set the horizontal alignment of pages when saving a document to HTML.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setPageHorizontalAlignment(pageHorizontalAlignment);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.HorizontalAlignment.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.HorizontalAlignment/styles.css");

        switch (pageHorizontalAlignment)
        {
            case HtmlFixedPageHorizontalAlignment.CENTER:
                Assert.assertTrue(Regex.match(outDocContents,
                    "[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; }").getSuccess());
                break;
            case HtmlFixedPageHorizontalAlignment.LEFT:
                Assert.assertTrue(Regex.match(outDocContents, 
                    "[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; }").getSuccess());
                break;
            case HtmlFixedPageHorizontalAlignment.RIGHT:
                Assert.assertTrue(Regex.match(outDocContents, 
                    "[.]awpage { position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; }").getSuccess());
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "horizontalAlignmentDataProvider")
	public static Object[][] horizontalAlignmentDataProvider() throws Exception
	{
		return new Object[][]
		{
			{HtmlFixedPageHorizontalAlignment.CENTER},
			{HtmlFixedPageHorizontalAlignment.LEFT},
			{HtmlFixedPageHorizontalAlignment.RIGHT},
		};
	}

    @Test
    public void pageMargins() throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageMargins
        //ExSummary:Shows how to adjust page margins when saving a document to HTML.
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setPageMargins(15.0);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.PageMargins.html", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.PageMargins/styles.css");

        Assert.assertTrue(Regex.match(outDocContents,
            "[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }").getSuccess());
        //ExEnd
    }

    @Test
    public void pageMarginsException()
    {
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        Assert.That(() => saveOptions.setPageMargins(-1), Throws.<IllegalArgumentException>TypeOf());
    }

    @Test (dataProvider = "optimizeGraphicsOutputDataProvider")
    public void optimizeGraphicsOutput(boolean optimizeOutput) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.OptimizeOutput
        //ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions(); { saveOptions.setOptimizeOutput(optimizeOutput); }
        
        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html", saveOptions);

        // The size of the optimized version of the document is almost a third of the size of the unoptimized document.
        Assert.assertEquals(optimizeOutput ? 57220 : 159000,
            new FileInfo(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html").getLength(), 200.0);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "optimizeGraphicsOutputDataProvider")
	public static Object[][] optimizeGraphicsOutputDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}


    @Test (dataProvider = "usingMachineFontsDataProvider")
    public void usingMachineFonts(boolean useTargetMachineFonts) throws Exception
    {
        //ExStart
        //ExFor:ExportFontFormat
        //ExFor:HtmlFixedSaveOptions.FontFormat
        //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
        //ExSummary:Shows how use fonts only from the target machine when saving a document to HTML.
        Document doc = new Document(getMyDir() + "Bullet points with alternative font.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setExportEmbeddedCss(true);
            saveOptions.setUseTargetMachineFonts(useTargetMachineFonts);
            saveOptions.setFontFormat(ExportFontFormat.TTF);
            saveOptions.setExportEmbeddedFonts(false);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.UsingMachineFonts.html");

        if (useTargetMachineFonts)
            Assert.assertFalse(Regex.match(outDocContents, "@font-face").getSuccess());
        else
            Assert.assertTrue(Regex.match(outDocContents,
                "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], " +
                "url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; }").getSuccess());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "usingMachineFontsDataProvider")
	public static Object[][] usingMachineFontsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    //ExStart
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs
    //ExFor:ResourceSavingArgs.Document
    //ExFor:ResourceSavingArgs.ResourceFileName
    //ExFor:ResourceSavingArgs.ResourceFileUri
    //ExSummary:Shows how to use a callback to track external resources created while converting a document to HTML.
    @Test //ExSkip
    public void resourceSavingCallback() throws Exception
    {
        Document doc = new Document(getMyDir() + "Bullet points with alternative font.docx");

        FontSavingCallback callback = new FontSavingCallback();

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setResourceSavingCallback(callback);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

        System.out.println(callback.getText());
        testResourceSavingCallback(callback); //ExSkip
    }

    private static class FontSavingCallback implements IResourceSavingCallback
    {
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        public void resourceSaving(ResourceSavingArgs args)
        {
            msStringBuilder.appendLine(mText, $"Original document URI:\t{args.Document.OriginalFileName}");
            msStringBuilder.appendLine(mText, $"Resource being saved:\t{args.ResourceFileName}");
            msStringBuilder.appendLine(mText, $"Full uri after saving:\t{args.ResourceFileUri}\n");
        }

        public String getText()
        {
            return mText.toString();
        }

        private /*final*/ StringBuilder mText = new StringBuilder();
    }
    //ExEnd

    private void testResourceSavingCallback(FontSavingCallback callback)
    {
        Assert.assertTrue(callback.getText().contains("font001.woff")); 
        Assert.assertTrue(callback.getText().contains("styles.css"));
    }

    //ExStart
    //ExFor:HtmlFixedSaveOptions
    //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
    //ExFor:HtmlFixedSaveOptions.ResourcesFolder
    //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:HtmlFixedSaveOptions.SaveFormat
    //ExFor:HtmlFixedSaveOptions.ShowPageBorder
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
    //ExFor:ResourceSavingArgs.ResourceStream
    //ExSummary:Shows how to use a callback to print the URIs of external resources created while converting a document to HTML.
    @Test //ExSkip
    public void htmlFixedResourceFolder() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        ResourceUriPrinter callback = new ResourceUriPrinter();

        HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
        {
            options.setSaveFormat(SaveFormat.HTML_FIXED);
            options.setExportEmbeddedImages(false);
            options.setResourcesFolder(getArtifactsDir() + "HtmlFixedResourceFolder");
            options.setResourcesFolderAlias(getArtifactsDir() + "HtmlFixedResourceFolderAlias");
            options.setShowPageBorder(false);
            options.setResourceSavingCallback(callback);
        }

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder.
        // We must ensure the folder exists before the streams can put their resources into it.
        Directory.createDirectory(options.getResourcesFolderAlias());

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

        System.out.println(callback.getText());

        String[] resourceFiles = Directory.getFiles(getArtifactsDir() + "HtmlFixedResourceFolderAlias");

        Assert.assertFalse(Directory.exists(getArtifactsDir() + "HtmlFixedResourceFolder"));
        Assert.AreEqual(6, resourceFiles.Count(f => f.EndsWith(".jpeg") || f.EndsWith(".png") || f.EndsWith(".css")));
        testHtmlFixedResourceFolder(callback); //ExSkip
    }
    
    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed HTML.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback
    {
        public void /*IResourceSavingCallback.*/resourceSaving(ResourceSavingArgs args) throws Exception
        {
            // If we set a folder alias in the SaveOptions object, we will be able to print it from here.
            msStringBuilder.appendLine(mText, $"Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");

            String extension = Path.getExtension(args.getResourceFileName());
            switch (gStringSwitchMap.of(extension))
            {
                case /*".ttf"*/0:
                case /*".woff"*/1:
                {
                    // By default, 'ResourceFileUri' uses system folder for fonts.
                    // To avoid problems in other platforms you must explicitly specify the path for the fonts.
                    args.setResourceFileUri(getArtifactsDir() + Path.DirectorySeparatorChar + args.getResourceFileName());
                    break;
                }
            }

            msStringBuilder.appendLine(mText, "\t" + args.getResourceFileUri());

            // If we have specified a folder in the "ResourcesFolderAlias" property,
            // we will also need to redirect each stream to put its resource in that folder.
            args.ResourceStream = new FileStream(args.getResourceFileUri(), FileMode.CREATE);
            args.setKeepResourceStreamOpen(false);
        }

        public String getText()
        {
            return mText.toString();
        }

        private int mSavedResourceCount;
        private /*final*/ StringBuilder mText = new StringBuilder();
    }
    //ExEnd

    private void testHtmlFixedResourceFolder(ResourceUriPrinter callback)
    {
        Assert.assertEquals(16, Regex.matches(callback.getText(), "Resource #").getCount());
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		".ttf",
		".woff"
	);

}
