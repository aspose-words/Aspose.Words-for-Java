// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.IResourceSavingCallback;
import com.aspose.words.ResourceSavingArgs;
import com.aspose.ms.System.msConsole;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.IO.Path;
import com.aspose.words.SaveFormat;
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
        //ExSummary:Shows how to set encoding while exporting to HTML.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("Hello World!");

        // The default encoding is UTF-8
        // If we want to represent our document using a different encoding, we can set one explicitly using a SaveOptions object
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

    // Note: Test doesn't contain validation result, because it's may take a lot of time for assert result
    // For validation result, you can save the document to HTML file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
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

    @Test (dataProvider = "exportEmbeddedCSSDataProvider")
    public void exportEmbeddedCSS(boolean doExportEmbeddedCss) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
        //ExSummary:Shows how to export embedded stylesheets into an HTML file.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedCss(doExportEmbeddedCss);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCSS.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCSS.html");

        if (doExportEmbeddedCss)
        {
            Assert.assertTrue(Regex.match(outDocContents, "<style type=\"text/css\">").getSuccess());
            Assert.assertFalse(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
        }
        else
        {
            Assert.assertTrue(Regex.match(outDocContents,
                "<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCSS/styles[.]css\" media=\"all\" />").getSuccess());
            Assert.assertTrue(File.exists(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportEmbeddedCSSDataProvider")
	public static Object[][] exportEmbeddedCSSDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test (dataProvider = "exportEmbeddedFontsDataProvider")
    public void exportEmbeddedFonts(boolean doExportEmbeddedFonts) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
        //ExSummary:Shows how to export embedded fonts into an HTML file.
        Document doc = new Document(getMyDir() + "Embedded font.docx");

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedFonts(doExportEmbeddedFonts);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

        if (doExportEmbeddedFonts)
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
    public void exportEmbeddedImages(boolean doExportImages) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
        //ExSummary:Shows how to export embedded images into an HTML file.
        Document doc = new Document(getMyDir() + "Images.docx");

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedImages(doExportImages);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages.html");

        if (doExportImages)
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
    public void exportEmbeddedSvgs(boolean doExportSvgs) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
        //ExSummary:Shows how to export embedded SVG objects into an HTML file.
        Document doc = new Document(getMyDir() + "Images.docx");

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportEmbeddedSvg(doExportSvgs);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

        if (doExportSvgs)
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
    public void exportFormFields(boolean doExportFormFields) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.ExportFormFields
        //ExSummary:Show how to exporting form fields from a document into HTML file.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCheckBox("CheckBox", false, 15);

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        {
            htmlFixedSaveOptions.setExportFormFields(doExportFormFields);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.ExportFormFields.html");

        if (doExportFormFields)
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
        //ExSummary:Shows how to add prefix to all class names in css file.
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
        //ExEnd
    }

    @Test (dataProvider = "horizontalAlignmentDataProvider")
    public void horizontalAlignment(/*HtmlFixedPageHorizontalAlignment*/int pageHorizontalAlignment) throws Exception
    {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
        //ExFor:HtmlFixedPageHorizontalAlignment
        //ExSummary:Shows how to set the horizontal alignment of pages in HTML file.
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
        //ExSummary:Shows how to set the margins around pages in HTML file.
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

    @Test
    public void optimizeGraphicsOutput() throws Exception
    {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExFor:HtmlFixedSaveOptions.OptimizeOutput
        //ExSummary:Shows how to optimize document objects while saving to html.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions(); { saveOptions.setOptimizeOutput(false); }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html", saveOptions);

        saveOptions.setOptimizeOutput(true);

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html", saveOptions);

        Assert.assertTrue(new FileInfo(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html").getLength() > 
                        new FileInfo(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html").getLength());
        //ExEnd
    }

    //ExStart
    //ExFor:ExportFontFormat
    //ExFor:HtmlFixedSaveOptions.FontFormat
    //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
    //ExFor:IResourceSavingCallback
    //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
    //ExFor:ResourceSavingArgs
    //ExFor:ResourceSavingArgs.Document
    //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
    //ExFor:ResourceSavingArgs.ResourceFileName
    //ExFor:ResourceSavingArgs.ResourceFileUri
    //ExFor:ResourceSavingArgs.ResourceStream
    //ExSummary:Shows how use target machine fonts to display the document.
    @Test //ExSkip
    public void usingMachineFonts() throws Exception
    {
        Document doc = new Document(getMyDir() + "Bullet points with alternative font.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setExportEmbeddedCss(true);
            saveOptions.setUseTargetMachineFonts(true);
            saveOptions.setFontFormat(ExportFontFormat.TTF);
            saveOptions.setExportEmbeddedFonts(false);
            saveOptions.setResourceSavingCallback(new ResourceSavingCallback());
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlFixedSaveOptions.UsingMachineFonts.html");

        if (saveOptions.getUseTargetMachineFonts())
            Assert.assertFalse(Regex.match(outDocContents, "@font-face").getSuccess());
        else
            Assert.assertTrue(Regex.match(outDocContents,
                "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], " +
                "url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; }").getSuccess());
    }

    private static class ResourceSavingCallback implements IResourceSavingCallback
    {
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        public void resourceSaving(ResourceSavingArgs args) throws Exception
        {
            System.out.println("Original document URI:\t{args.Document.OriginalFileName}");
            System.out.println("Resource being saved:\t{args.ResourceFileName}");
            System.out.println("Full uri after saving:\t{args.ResourceFileUri}");

            args.ResourceStream = new MemoryStream();
            args.setKeepResourceStreamOpen(true);

            String extension = Path.getExtension(args.getResourceFileName());
            switch (gStringSwitchMap.of(extension))
            {
                case /*".ttf"*/0:
                case /*".woff"*/1:
                {
                    Assert.fail("'ResourceSavingCallback' is not fired for fonts when 'UseTargetMachineFonts' is true");
                    break;
                }
            }
        }
    }
    //ExEnd

    //ExStart
    //ExFor:HtmlFixedSaveOptions
    //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
    //ExFor:HtmlFixedSaveOptions.ResourcesFolder
    //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
    //ExFor:HtmlFixedSaveOptions.SaveFormat
    //ExFor:HtmlFixedSaveOptions.ShowPageBorder
    //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .html.
    @Test //ExSkip
    public void htmlFixedResourceFolder() throws Exception
    {
        // Open a document which contains images
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions options = new HtmlFixedSaveOptions();
        {
            options.setSaveFormat(SaveFormat.HTML_FIXED);
            options.setExportEmbeddedImages(false);
            options.setResourcesFolder(getArtifactsDir() + "HtmlFixedResourceFolder");
            options.setResourcesFolderAlias(getArtifactsDir() + "HtmlFixedResourceFolderAlias");
            options.setShowPageBorder(false);
            options.setResourceSavingCallback(new ResourceUriPrinter());
        }

        // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
        // We must ensure the folder exists before the streams can put their resources into it
        Directory.createDirectory(options.getResourcesFolderAlias());

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

        String[] resourceFiles = Directory.getFiles(getArtifactsDir() + "HtmlFixedResourceFolderAlias");

        Assert.assertFalse(Directory.exists(getArtifactsDir() + "HtmlFixedResourceFolder"));
        Assert.AreEqual(6, resourceFiles.Count(f => f.EndsWith(".jpeg") || f.EndsWith(".png") || f.EndsWith(".css")));
    }

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed .Html
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback
    {
        public void /*IResourceSavingCallback.*/resourceSaving(ResourceSavingArgs args) throws Exception
        {
            // If we set a folder alias in the SaveOptions object, it will be printed here
            System.out.println("Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");

            String extension = Path.getExtension(args.getResourceFileName());
            switch (gStringSwitchMap.of(extension))
            {
                case /*".ttf"*/0:
                case /*".woff"*/1:
                {
                    // By default 'ResourceFileUri' used system folder for fonts
                    // To avoid problems across platforms you must explicitly specify the path for the fonts
                    args.setResourceFileUri(getArtifactsDir() + Path.DirectorySeparatorChar + args.getResourceFileName());
                    break;
                }
            }
            System.out.println("\t" + args.getResourceFileUri());

            // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
            args.ResourceStream = new FileStream(args.getResourceFileUri(), FileMode.CREATE);
            args.setKeepResourceStreamOpen(false);
        }

        private int mSavedResourceCount;
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		".ttf",
		".woff"
	);

    //ExEnd
}
