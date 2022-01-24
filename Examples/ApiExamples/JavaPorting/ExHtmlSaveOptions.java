// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.SaveFormat;
import com.aspose.words.Document;
import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.HtmlOfficeMathOutputMode;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.IO.SearchOption;
import org.testng.Assert;
import com.aspose.words.ExportListLabels;
import com.aspose.words.List;
import com.aspose.words.ListTemplate;
import com.aspose.words.CssStyleSheetType;
import com.aspose.words.HtmlVersion;
import com.aspose.words.HtmlMetafileFormat;
import com.aspose.ms.System.IO.File;
import com.aspose.words.FontSettings;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.Table;
import com.aspose.words.PreferredWidth;
import com.aspose.words.BreakType;
import com.aspose.words.HtmlElementSizeOutputMode;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IFontSavingCallback;
import com.aspose.words.FontSavingArgs;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.ms.System.Globalization.msCultureInfo;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.PageSetup;
import com.aspose.words.PaperSize;
import com.aspose.words.FieldToc;
import com.aspose.words.FieldType;
import com.aspose.ms.System.IO.FileInfo;
import com.aspose.words.HtmlLoadOptions;
import com.aspose.ms.System.IO.MemoryStream;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import com.aspose.ms.System.Drawing.msSize;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.words.LayoutCollector;
import org.testng.annotations.DataProvider;


@Test
class ExHtmlSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "exportPageMarginsEpubDataProvider")
    public void exportPageMarginsEpub(/*SaveFormat*/int saveFormat) throws Exception
    {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setSaveFormat(saveFormat);
            saveOptions.setExportPageMargins(true);
        }

        doc.save(
            getArtifactsDir() + "HtmlSaveOptions.ExportPageMarginsEpub" +
            FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportPageMarginsEpubDataProvider")
	public static Object[][] exportPageMarginsEpubDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.HTML},
			{SaveFormat.MHTML},
			{SaveFormat.EPUB},
		};
	}

    @Test (dataProvider = "exportOfficeMathEpubDataProvider")
    public void exportOfficeMathEpub(/*SaveFormat*/int saveFormat, /*HtmlOfficeMathOutputMode*/int outputMode) throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setOfficeMathOutputMode(outputMode); }

        doc.save(
            getArtifactsDir() + "HtmlSaveOptions.ExportOfficeMathEpub" +
            FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportOfficeMathEpubDataProvider")
	public static Object[][] exportOfficeMathEpubDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.HTML,  HtmlOfficeMathOutputMode.IMAGE},
			{SaveFormat.MHTML,  HtmlOfficeMathOutputMode.MATH_ML},
			{SaveFormat.EPUB,  HtmlOfficeMathOutputMode.TEXT},
		};
	}

    @Test (dataProvider = "exportTextBoxAsSvgEpubDataProvider")
    public void exportTextBoxAsSvgEpub(/*SaveFormat*/int saveFormat, boolean isTextBoxAsSvg) throws Exception
    {
        String[] dirFiles;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textbox = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 100.0);
        builder.moveTo(textbox.getFirstParagraph());
        builder.write("Hello world!");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat);
        saveOptions.setExportTextBoxAsSvg(isTextBoxAsSvg);
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTextBoxAsSvgEpub" + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);

        switch (saveFormat)
        {
            case SaveFormat.HTML:

                dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                    SearchOption.ALL_DIRECTORIES);
                Assert.That(dirFiles, Is.Empty);
                return;

            case SaveFormat.EPUB:

                dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                    SearchOption.ALL_DIRECTORIES);
                Assert.That(dirFiles, Is.Empty);
                return;

            case SaveFormat.MHTML:

                dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png",
                    SearchOption.ALL_DIRECTORIES);
                Assert.That(dirFiles, Is.Empty);
                return;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportTextBoxAsSvgEpubDataProvider")
	public static Object[][] exportTextBoxAsSvgEpubDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.HTML,  true},
			{SaveFormat.EPUB,  true},
			{SaveFormat.MHTML,  false},
		};
	}

    @Test (dataProvider = "controlListLabelsExportDataProvider")
    public void controlListLabelsExport(/*ExportListLabels*/int howExportListLabels) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        List bulletedList = doc.getLists().add(ListTemplate.BULLET_DEFAULT);
        builder.getListFormat().setList(bulletedList);
        builder.getParagraphFormat().setLeftIndent(72.0);
        builder.writeln("Bulleted list item 1.");
        builder.writeln("Bulleted list item 2.");
        builder.getParagraphFormat().clearFormatting();

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it does not cause formatting loss, 
            // otherwise HTML <p> tag is used. This is also the default value.
            // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
            // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
            saveOptions.setExportListLabels(howExportListLabels);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "controlListLabelsExportDataProvider")
	public static Object[][] controlListLabelsExportDataProvider() throws Exception
	{
		return new Object[][]
		{
			{ExportListLabels.AUTO},
			{ExportListLabels.AS_INLINE_TEXT},
			{ExportListLabels.BY_HTML_TAGS},
		};
	}

    @Test (dataProvider = "exportUrlForLinkedImageDataProvider")
    public void exportUrlForLinkedImage(boolean export) throws Exception
    {
        Document doc = new Document(getMyDir() + "Linked image.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setExportOriginalUrlForLinkedImages(export); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        String[] dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportUrlForLinkedImage.001.png",
            SearchOption.ALL_DIRECTORIES);

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html",
            dirFiles.length == 0
                ? "<img src=\"http://www.aspose.com/images/aspose-logo.gif\""
                : "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportUrlForLinkedImageDataProvider")
	public static Object[][] exportUrlForLinkedImageDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}

    @Test
    public void exportRoundtripInformation() throws Exception
    {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setExportRoundtripInformation(true); }
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html", saveOptions);
    }

    @Test
    public void roundtripInformationDefaulValue()
    {
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        Assert.assertEquals(true, saveOptions.getExportRoundtripInformation());

        saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        Assert.assertEquals(false, saveOptions.getExportRoundtripInformation());

        saveOptions = new HtmlSaveOptions(SaveFormat.EPUB);
        Assert.assertEquals(false, saveOptions.getExportRoundtripInformation());
    }

    @Test
    public void externalResourceSavingConfig() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setExportFontResources(true);
            saveOptions.setResourceFolder("Resources");
            saveOptions.setResourceFolderAlias("https://www.aspose.com/");
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExternalResourceSavingConfig.html", saveOptions);

        String[] imageFiles = Directory.getFiles(getArtifactsDir() + "Resources/",
            "HtmlSaveOptions.ExternalResourceSavingConfig*.png", SearchOption.ALL_DIRECTORIES);
        Assert.assertEquals(8, imageFiles.length);

        String[] fontFiles = Directory.getFiles(getArtifactsDir() + "Resources/",
            "HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", SearchOption.ALL_DIRECTORIES);
        Assert.assertEquals(10, fontFiles.length);

        String[] cssFiles = Directory.getFiles(getArtifactsDir() + "Resources/",
            "HtmlSaveOptions.ExternalResourceSavingConfig*.css", SearchOption.ALL_DIRECTORIES);
        Assert.assertEquals(1, cssFiles.length);

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExternalResourceSavingConfig.html",
            "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
    }

    @Test
    public void convertFontsAsBase64() throws Exception
    {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setResourceFolder("Resources");
            saveOptions.setExportFontResources(true);
            saveOptions.setExportFontsAsBase64(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
	}

    @Test (dataProvider = "html5SupportDataProvider")
    public void html5Support(/*HtmlVersion*/int htmlVersion) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setHtmlVersion(htmlVersion); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.Html5Support.html", saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "html5SupportDataProvider")
	public static Object[][] html5SupportDataProvider() throws Exception
	{
		return new Object[][]
		{
			{HtmlVersion.HTML_5},
			{HtmlVersion.XHTML},
		};
	}

    @Test (dataProvider = "exportFontsDataProvider")
    public void exportFonts(boolean exportAsBase64) throws Exception
    {
        String fontsFolder = getArtifactsDir() + "HtmlSaveOptions.ExportFonts.Resources";
        
        Document doc = new Document(getMyDir() + "Document.docx");
        
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setExportFontResources(true);
            saveOptions.setFontsFolder(fontsFolder);
            saveOptions.setExportFontsAsBase64(exportAsBase64);
        }

        switch (exportAsBase64)
        {
            case false:

                doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFonts.False.html", saveOptions);

                Assert.IsNotEmpty(Directory.getFiles(fontsFolder, "HtmlSaveOptions.ExportFonts.False.times.ttf",
                    SearchOption.ALL_DIRECTORIES));

                Directory.delete(fontsFolder, true);
                break;

            case true:

                doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFonts.True.html", saveOptions);
                Assert.assertFalse(Directory.exists(fontsFolder));
                break;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportFontsDataProvider")
	public static Object[][] exportFontsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void resourceFolderPriority() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setExportFontResources(true);
            saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
            saveOptions.setResourceFolderAlias("http://example.com/resources");
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.001.png", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.002.png", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.arial.ttf", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.css", SearchOption.ALL_DIRECTORIES));
    }

    @Test
    public void resourceFolderLowPriority() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setExportFontResources(true);
            saveOptions.setFontsFolder(getArtifactsDir() + "Fonts");
            saveOptions.setImagesFolder(getArtifactsDir() + "Images");
            saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
            saveOptions.setResourceFolderAlias("http://example.com/resources");
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolderLowPriority.html", saveOptions);

        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Images",
            "HtmlSaveOptions.ResourceFolderLowPriority.001.png", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolderLowPriority.002.png",
            SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Fonts",
            "HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderLowPriority.css",
            SearchOption.ALL_DIRECTORIES));
    }

    @Test
    public void svgMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an SVG image: ");
        builder.insertHtml(
            "<svg height='210' width='500'>\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\n                  </svg> ");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.SvgMetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.PNG); });
    }

    @Test
    public void pngMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an Png image: ");
        builder.insertHtml(
            "<svg height='210' width='500'>\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\n                  </svg> ");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.PngMetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.PNG); });
    }

    @Test
    public void emfOrWmfMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an image as is: ");
        builder.insertHtml(
            "<img src=\"data:image/png;base64,\n                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\n                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\n                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\n                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\n                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\n                    vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.EmfOrWmfMetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF); });
    }

    @Test
    public void cssClassNamesPrefix() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.CssClassNamePrefix
        //ExSummary:Shows how to save a document to HTML, and add a prefix to all of its CSS class names.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setCssClassNamePrefix("myprefix-");
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html");

        Assert.assertTrue(outDocContents.contains("<p class=\"myprefix-Header\">"));
        Assert.assertTrue(outDocContents.contains("<p class=\"myprefix-Footer\">"));

        outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.css");

        Assert.assertTrue(outDocContents.contains(".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n" +
                                            ".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n"));
        //ExEnd
    }

    @Test
    public void cssClassNamesNotValidPrefix()
    {
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        Assert.<IllegalArgumentException>Throws(() => saveOptions.setCssClassNamePrefix("@%-"),
            "The class name prefix must be a valid CSS identifier.");
    }

    @Test
    public void cssClassNamesNullPrefix() throws Exception
    {
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
            saveOptions.setCssClassNamePrefix(null);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
    }

    @Test
    public void contentIdScheme() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        {
            saveOptions.setPrettyFormat(true);
            saveOptions.setExportCidUrlsForMhtmlResources(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
    }

    @Test (enabled = false, description = "Bug", dataProvider = "resolveFontNamesDataProvider")
    public void resolveFontNames(boolean resolveFontNames) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ResolveFontNames
        //ExSummary:Shows how to resolve all font names before writing them to HTML.
        Document doc = new Document(getMyDir() + "Missing font.docx");

        // This document contains text that names a font that we do not have.
        Assert.assertNotNull(doc.getFontInfos().get("28 Days Later"));

        // If we have no way of getting this font, and we want to be able to display all the text
        // in this document in an output HTML, we can substitute it with another font.
        FontSettings fontSettings = new FontSettings();
        {
            fontSettings.setSubstitutionSettings({
                fontSettings.getSubstitutionSettings().setDefaultFontSubstitution({
                    fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
                    fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setEnabled(true);
                });
            });
        }

        doc.setFontSettings(fontSettings);
        
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            // By default, this option is set to 'False' and Aspose.Words writes font names as specified in the source document
            saveOptions.setResolveFontNames(resolveFontNames);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html");

        Assert.assertTrue(resolveFontNames
            ? Regex.match(outDocContents, "<span style=\"font-family:Arial\">").getSuccess()
            : Regex.match(outDocContents, "<span style=\"font-family:\'28 Days Later\'\">").getSuccess());
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "resolveFontNamesDataProvider")
	public static Object[][] resolveFontNamesDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void headingLevels() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.DocumentSplitHeadingLevel
        //ExSummary:Shows how to split an output HTML document by headings into several parts.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Every paragraph that we format using a "Heading" style can serve as a heading.
        // Each heading may also have a heading level, determined by the number of its heading style.
        // The headings below are of levels 1-3.
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 1"));
        builder.writeln("Heading #1");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 2"));
        builder.writeln("Heading #2");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 3"));
        builder.writeln("Heading #3");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 1"));
        builder.writeln("Heading #4");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 2"));
        builder.writeln("Heading #5");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 3"));
        builder.writeln("Heading #6");

        // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph".
        // These criteria will split the document at paragraphs with "Heading" styles into several smaller documents,
        // and save each document in a separate HTML file in the local file system.
        // We will also set the maximum heading level, which splits the document to 2.
        // Saving the document will split it at headings of levels 1 and 2, but not at 3 to 9.
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);
            options.setDocumentSplitHeadingLevel(2);
        }
        
        // Our document has four headings of levels 1 - 2. One of those headings will not be
        // a split point since it is at the beginning of the document.
        // The saving operation will split our document at three places, into four smaller documents.
        doc.save(getArtifactsDir() + "HtmlSaveOptions.HeadingLevels.html", options);

        doc = new Document(getArtifactsDir() + "HtmlSaveOptions.HeadingLevels.html");

        Assert.assertEquals("Heading #1", doc.getText().trim());

        doc = new Document(getArtifactsDir() + "HtmlSaveOptions.HeadingLevels-01.html");

        Assert.assertEquals("Heading #2\r" +
                        "Heading #3", doc.getText().trim());

        doc = new Document(getArtifactsDir() + "HtmlSaveOptions.HeadingLevels-02.html");

        Assert.assertEquals("Heading #4", doc.getText().trim());

        doc = new Document(getArtifactsDir() + "HtmlSaveOptions.HeadingLevels-03.html");

        Assert.assertEquals("Heading #5\r" +
                        "Heading #6", doc.getText().trim());
        //ExEnd
    }

    @Test (dataProvider = "negativeIndentDataProvider")
    public void negativeIndent(boolean allowNegativeIndent) throws Exception
    {
        //ExStart
        //ExFor:HtmlElementSizeOutputMode
        //ExFor:HtmlSaveOptions.AllowNegativeIndent
        //ExFor:HtmlSaveOptions.TableWidthOutputMode
        //ExSummary:Shows how to preserve negative indents in the output .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with a negative indent, which will push it to the left past the left page boundary.
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, Cell 1");
        builder.insertCell();
        builder.write("Row 1, Cell 2");
        builder.endTable();
        table.setLeftIndent(-36);
        table.setPreferredWidth(PreferredWidth.fromPoints(144.0));

        builder.insertBreak(BreakType.PARAGRAPH_BREAK);

        // Insert a table with a positive indent, which will push the table to the right.
        table = builder.startTable();
        builder.insertCell();
        builder.write("Row 1, Cell 1");
        builder.insertCell();
        builder.write("Row 1, Cell 2");
        builder.endTable();
        table.setLeftIndent(36.0);
        table.setPreferredWidth(PreferredWidth.fromPoints(144.0));

        // When we save a document to HTML, Aspose.Words will only preserve negative indents
        // such as the one we have applied to the first table if we set the "AllowNegativeIndent" flag
        // in a SaveOptions object that we will pass to "true".
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        {
            options.setAllowNegativeIndent(allowNegativeIndent);
            options.setTableWidthOutputMode(HtmlElementSizeOutputMode.RELATIVE_ONLY);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.NegativeIndent.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.NegativeIndent.html");

        if (allowNegativeIndent)
        {
            Assert.assertTrue(outDocContents.contains(
                "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
            Assert.assertTrue(outDocContents.contains(
                "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains(
                "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
            Assert.assertTrue(outDocContents.contains(
                "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "negativeIndentDataProvider")
	public static Object[][] negativeIndentDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void folderAlias() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportOriginalUrlForLinkedImages
        //ExFor:HtmlSaveOptions.FontsFolder
        //ExFor:HtmlSaveOptions.FontsFolderAlias
        //ExFor:HtmlSaveOptions.ImageResolution
        //ExFor:HtmlSaveOptions.ImagesFolderAlias
        //ExFor:HtmlSaveOptions.ResourceFolder
        //ExFor:HtmlSaveOptions.ResourceFolderAlias
        //ExSummary:Shows how to set folders and folder aliases for externally saved resources that Aspose.Words will create when saving a document to HTML.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            options.setExportFontResources(true);
            options.setImageResolution(72);
            options.setFontResourcesSubsettingSizeThreshold(0);
            options.setFontsFolder(getArtifactsDir() + "Fonts");
            options.setImagesFolder(getArtifactsDir() + "Images");
            options.setResourceFolder(getArtifactsDir() + "Resources");
            options.setFontsFolderAlias("http://example.com/fonts");
            options.setImagesFolderAlias("http://example.com/images");
            options.setResourceFolderAlias("http://example.com/resources");
            options.setExportOriginalUrlForLinkedImages(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.FolderAlias.html", options);
        //ExEnd
    }

    //ExStart
    //ExFor:HtmlSaveOptions.ExportFontResources
    //ExFor:HtmlSaveOptions.FontSavingCallback
    //ExFor:IFontSavingCallback
    //ExFor:IFontSavingCallback.FontSaving
    //ExFor:FontSavingArgs
    //ExFor:FontSavingArgs.Bold
    //ExFor:FontSavingArgs.Document
    //ExFor:FontSavingArgs.FontFamilyName
    //ExFor:FontSavingArgs.FontFileName
    //ExFor:FontSavingArgs.FontStream
    //ExFor:FontSavingArgs.IsExportNeeded
    //ExFor:FontSavingArgs.IsSubsettingNeeded
    //ExFor:FontSavingArgs.Italic
    //ExFor:FontSavingArgs.KeepFontStreamOpen
    //ExFor:FontSavingArgs.OriginalFileName
    //ExFor:FontSavingArgs.OriginalFileSize
    //ExSummary:Shows how to define custom logic for exporting fonts when saving to HTML.
    @Test //ExSkip
    public void saveExportedFonts() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Configure a SaveOptions object to export fonts to separate files.
        // Set a callback that will handle font saving in a custom manner.
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportFontResources(true);
            options.setFontSavingCallback(new HandleFontSaving());
        }

        // The callback will export .ttf files and save them alongside the output document.
        doc.save(getArtifactsDir() + "HtmlSaveOptions.SaveExportedFonts.html", options);

        for (String fontFilename : Object[].FindAll(Directory.getFiles(getArtifactsDir()), s => s.endsWith(".ttf")))
        {
            System.out.println(fontFilename);
        }

        Assert.assertEquals(10, Object[].FindAll(Directory.getFiles(getArtifactsDir()), s => s.endsWith(".ttf")).length); //ExSkip
    }

    /// <summary>
    /// Prints information about exported fonts and saves them in the same local system folder as their output .html.
    /// </summary>
    public static class HandleFontSaving implements IFontSavingCallback
    {
        public void /*IFontSavingCallback.*/fontSaving(FontSavingArgs args) throws Exception
        {
            msConsole.write($"Font:\t{args.FontFamilyName}");
            if (args.getBold()) msConsole.write(", bold");
            if (args.getItalic()) msConsole.write(", italic");
            System.out.println("\nSource:\t{args.OriginalFileName}, {args.OriginalFileSize} bytes\n");

            // We can also access the source document from here.
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.docx"));

            Assert.assertTrue(args.isExportNeeded());
            Assert.assertTrue(args.isSubsettingNeeded());

            // There are two ways of saving an exported font.
            // 1 -  Save it to a local file system location:
            args.setFontFileName(msString.split(args.getOriginalFileName(), Path.DirectorySeparatorChar).Last());

            // 2 -  Save it to a stream:
            args.FontStream =
                new FileStream(getArtifactsDir() + msString.split(args.getOriginalFileName(), Path.DirectorySeparatorChar).Last(), FileMode.CREATE);
            Assert.assertFalse(args.getKeepFontStreamOpen());
        }
    }
    //ExEnd

    @Test (dataProvider = "htmlVersionsDataProvider")
    public void htmlVersions(/*HtmlVersion*/int htmlVersion) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
        //ExFor:HtmlSaveOptions.HtmlVersion
        //ExFor:HtmlVersion
        //ExSummary:Shows how to save a document to a specific version of HTML.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        {
            options.setHtmlVersion(htmlVersion);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.HtmlVersions.html", options);

        // Our HTML documents will have minor differences to be compatible with different HTML versions.
        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.HtmlVersions.html");

        switch (htmlVersion)
        {
            case HtmlVersion.HTML_5:
                Assert.assertTrue(outDocContents.contains("<a id=\"_Toc76372689\"></a>"));
                Assert.assertTrue(outDocContents.contains("<a id=\"_Toc76372689\"></a>"));
                Assert.assertTrue(outDocContents.contains("<table style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
                break;
            case HtmlVersion.XHTML:
                Assert.assertTrue(outDocContents.contains("<a name=\"_Toc76372689\"></a>"));
                Assert.assertTrue(outDocContents.contains("<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">"));
                Assert.assertTrue(outDocContents.contains("<table cellspacing=\"0\" cellpadding=\"0\" style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\""));
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "htmlVersionsDataProvider")
	public static Object[][] htmlVersionsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{HtmlVersion.HTML_5},
			{HtmlVersion.XHTML},
		};
	}

    @Test (dataProvider = "exportXhtmlTransitionalDataProvider")
    public void exportXhtmlTransitional(boolean showDoctypeDeclaration) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
        //ExFor:HtmlSaveOptions.HtmlVersion
        //ExFor:HtmlVersion
        //ExSummary:Shows how to display a DOCTYPE heading when converting documents to the Xhtml 1.0 transitional standard.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        {
            options.setHtmlVersion(HtmlVersion.XHTML);
            options.setExportXhtmlTransitional(showDoctypeDeclaration);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportXhtmlTransitional.html", options);

        // Our document will only contain a DOCTYPE declaration heading if we have set the "ExportXhtmlTransitional" flag to "true".
        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportXhtmlTransitional.html");

        if (showDoctypeDeclaration)
            Assert.assertTrue(outDocContents.contains(
                "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>\r\n" +
                "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
                "<html xmlns=\"http://www.w3.org/1999/xhtml\">"));
        else
            Assert.assertTrue(outDocContents.contains("<html>"));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportXhtmlTransitionalDataProvider")
	public static Object[][] exportXhtmlTransitionalDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void epubHeadings() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.EpubNavigationMapLevel
        //ExSummary:Shows how to filter headings that appear in the navigation panel of a saved Epub document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Every paragraph that we format using a "Heading" style can serve as a heading.
        // Each heading may also have a heading level, determined by the number of its heading style.
        // The headings below are of levels 1-3.
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 1"));
        builder.writeln("Heading #1");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 2"));
        builder.writeln("Heading #2");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 3"));
        builder.writeln("Heading #3");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 1"));
        builder.writeln("Heading #4");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 2"));
        builder.writeln("Heading #5");
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 3"));
        builder.writeln("Heading #6");

        // Epub readers typically create a table of contents for their documents.
        // Each paragraph with a "Heading" style in the document will create an entry in this table of contents.
        // We can use the "EpubNavigationMapLevel" property to set a maximum heading level. 
        // The Epub reader will not add headings with a level above the one we specify to the contents table.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.EPUB);
        options.setEpubNavigationMapLevel(2);

        // Our document has six headings, two of which are above level 2.
        // The table of contents for this document will have four entries.
        doc.save(getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", options);
        //ExEnd

        TestUtil.docPackageFileContainsString("<navLabel><text>Heading #1</text></navLabel>", 
            getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
        TestUtil.docPackageFileContainsString("<navLabel><text>Heading #2</text></navLabel>", 
            getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
        TestUtil.docPackageFileContainsString("<navLabel><text>Heading #4</text></navLabel>", 
            getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
        TestUtil.docPackageFileContainsString("<navLabel><text>Heading #5</text></navLabel>", 
            getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");

        Assert.<AssertionError>Throws(() =>
        {
            TestUtil.docPackageFileContainsString("<navLabel><text>Heading #3</text></navLabel>", 
                getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
        });

        Assert.<AssertionError>Throws(() =>
        {
            TestUtil.docPackageFileContainsString("<navLabel><text>Heading #6</text></navLabel>", 
                getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", "HtmlSaveOptions.EpubHeadings.ncx");
        });
    }

    @Test
    public void doc2EpubSaveOptions() throws Exception
    {
        //ExStart
        //ExFor:DocumentSplitCriteria
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.#ctor
        //ExFor:HtmlSaveOptions.Encoding
        //ExFor:HtmlSaveOptions.DocumentSplitCriteria
        //ExFor:HtmlSaveOptions.ExportDocumentProperties
        //ExFor:HtmlSaveOptions.SaveFormat
        //ExFor:SaveOptions
        //ExFor:SaveOptions.SaveFormat
        //ExSummary:Shows how to use a specific encoding when saving a document to .epub.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Use a SaveOptions object to specify the encoding for a document that we will save.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setSaveFormat(SaveFormat.EPUB);
        saveOptions.setEncodingInternal(Encoding.getUTF8());

        // By default, an output .epub document will have all its contents in one HTML part.
        // A split criterion allows us to segment the document into several HTML parts.
        // We will set the criteria to split the document into heading paragraphs.
        // This is useful for readers who cannot read HTML files more significant than a specific size.
        saveOptions.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);

        // Specify that we want to export document properties.
        saveOptions.setExportDocumentProperties(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
        //ExEnd
    }

    @Test (dataProvider = "contentIdUrlsDataProvider")
    public void contentIdUrls(boolean exportCidUrlsForMhtmlResources) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
        //ExSummary:Shows how to enable content IDs for output MHTML documents.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Setting this flag will replace "Content-Location" tags
        // with "Content-ID" tags for each resource from the input document.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.MHTML);
        {
            options.setExportCidUrlsForMhtmlResources(exportCidUrlsForMhtmlResources);
            options.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            options.setExportFontResources(true);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ContentIdUrls.mht", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ContentIdUrls.mht");

        if (exportCidUrlsForMhtmlResources)
        {
            Assert.assertTrue(outDocContents.contains("Content-ID: <document.html>"));
            Assert.assertTrue(outDocContents.contains("<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
            Assert.assertTrue(outDocContents.contains("@font-face { font-family:'Arial Black'; src:url('cid:ariblk.ttf') }"));
            Assert.assertTrue(outDocContents.contains("<img src=3D\"cid:image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />"));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains("Content-Location: document.html"));
            Assert.assertTrue(outDocContents.contains("<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
            Assert.assertTrue(outDocContents.contains("@font-face { font-family:'Arial Black'; src:url('ariblk.ttf') }"));
            Assert.assertTrue(outDocContents.contains("<img src=3D\"image.003.jpeg\" width=3D\"350\" height=3D\"180\" alt=3D\"\" />"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "contentIdUrlsDataProvider")
	public static Object[][] contentIdUrlsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "dropDownFormFieldDataProvider")
    public void dropDownFormField(boolean exportDropDownFormFieldAsText) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
        //ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a combo box with the value "Two" selected.
        builder.insertComboBox("MyComboBox", new String[] { "One", "Two", "Three" }, 1);

        // The "ExportDropDownFormFieldAsText" flag of this SaveOptions object allows us to
        // control how saving the document to HTML treats drop-down combo boxes.
        // Setting it to "true" will convert each combo box into simple text
        // that displays the combo box's currently selected value, effectively freezing it.
        // Setting it to "false" will preserve the functionality of the combo box using <select> and <option> tags.
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportDropDownFormFieldAsText(exportDropDownFormFieldAsText);    

        doc.save(getArtifactsDir() + "HtmlSaveOptions.DropDownFormField.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.DropDownFormField.html");

        if (exportDropDownFormFieldAsText)
            Assert.assertTrue(outDocContents.contains(
                "<span>Two</span>"));
        else
            Assert.assertTrue(outDocContents.contains(
                "<select name=\"MyComboBox\">" +
                    "<option>One</option>" +
                    "<option selected=\"selected\">Two</option>" +
                    "<option>Three</option>" +
                "</select>"));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "dropDownFormFieldDataProvider")
	public static Object[][] dropDownFormFieldDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "exportImagesAsBase64DataProvider")
    public void exportImagesAsBase64(boolean exportItemsAsBase64) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontsAsBase64
        //ExFor:HtmlSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to save a .html document with images embedded inside it.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportImagesAsBase64(exportItemsAsBase64);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportImagesAsBase64.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportImagesAsBase64.html");

        Assert.assertTrue(exportItemsAsBase64
            ? outDocContents.contains("<img src=\"data:image/png;base64")
            : outDocContents.contains("<img src=\"HtmlSaveOptions.ExportImagesAsBase64.001.png\""));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportImagesAsBase64DataProvider")
	public static Object[][] exportImagesAsBase64DataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}


    @Test
    public void exportFontsAsBase64() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontsAsBase64
        //ExFor:HtmlSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to embed fonts inside a saved HTML document.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportFontsAsBase64(true);
            options.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFontsAsBase64.html", options);
        //ExEnd
    }

    @Test (dataProvider = "exportLanguageInformationDataProvider")
    public void exportLanguageInformation(boolean exportLanguageInformation) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportLanguageInformation
        //ExSummary:Shows how to preserve language information when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use the builder to write text while formatting it in different locales.
        builder.getFont().setLocaleId(new msCultureInfo("en-US").getLCID());
        builder.writeln("Hello world!");

        builder.getFont().setLocaleId(new msCultureInfo("en-GB").getLCID());
        builder.writeln("Hello again!");

        builder.getFont().setLocaleId(new msCultureInfo("ru-RU").getLCID());
        builder.write("Привет, мир!");

        // When saving the document to HTML, we can pass a SaveOptions object
        // to either preserve or discard each formatted text's locale.
        // If we set the "ExportLanguageInformation" flag to "true",
        // the output HTML document will contain the locales in "lang" attributes of <span> tags.
        // If we set the "ExportLanguageInformation" flag to "false',
        // the text in the output HTML document will not contain any locale information.
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportLanguageInformation(exportLanguageInformation);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportLanguageInformation.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportLanguageInformation.html");

        if (exportLanguageInformation)
        {
            Assert.assertTrue(outDocContents.contains("<span>Hello world!</span>"));
            Assert.assertTrue(outDocContents.contains("<span lang=\"en-GB\">Hello again!</span>"));
            Assert.assertTrue(outDocContents.contains("<span lang=\"ru-RU\">Привет, мир!</span>"));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains("<span>Hello world!</span>"));
            Assert.assertTrue(outDocContents.contains("<span>Hello again!</span>"));
            Assert.assertTrue(outDocContents.contains("<span>Привет, мир!</span>"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportLanguageInformationDataProvider")
	public static Object[][] exportLanguageInformationDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "listDataProvider")
    public void list(/*ExportListLabels*/int exportListLabels) throws Exception
    {
        //ExStart
        //ExFor:ExportListLabels
        //ExFor:HtmlSaveOptions.ExportListLabels
        //ExSummary:Shows how to configure list exporting to HTML.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        List list = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);
        builder.getListFormat().setList(list);
        
        builder.writeln("Default numbered list item 1.");
        builder.writeln("Default numbered list item 2.");
        builder.getListFormat().listIndent();
        builder.writeln("Default numbered list item 3.");
        builder.getListFormat().removeNumbers();

        list = doc.getLists().add(ListTemplate.OUTLINE_HEADINGS_LEGAL);
        builder.getListFormat().setList(list);

        builder.writeln("Outline legal heading list item 1.");
        builder.writeln("Outline legal heading list item 2.");
        builder.getListFormat().listIndent();
        builder.writeln("Outline legal heading list item 3.");
        builder.getListFormat().listIndent();
        builder.writeln("Outline legal heading list item 4.");
        builder.getListFormat().listIndent();
        builder.writeln("Outline legal heading list item 5.");
        builder.getListFormat().removeNumbers();

        // When saving the document to HTML, we can pass a SaveOptions object
        // to decide which HTML elements the document will use to represent lists.
        // Setting the "ExportListLabels" property to "ExportListLabels.AsInlineText"
        // will create lists by formatting spans.
        // Setting the "ExportListLabels" property to "ExportListLabels.Auto" will use the <p> tag
        // to build lists in cases when using the <ol> and <li> tags may cause loss of formatting.
        // Setting the "ExportListLabels" property to "ExportListLabels.ByHtmlTags"
        // will use <ol> and <li> tags to build all lists.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportListLabels(exportListLabels); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.List.html", options);
        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.List.html");

        switch (exportListLabels)
        {
            case ExportListLabels.AS_INLINE_TEXT:
                Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-left:72pt; margin-bottom:0pt; text-indent:-18pt; -aw-import:list-item; -aw-list-level-number:1; -aw-list-number-format:'%1.'; -aw-list-number-styles:'lowerLetter'; -aw-list-number-values:'1'; -aw-list-padding-sml:9.67pt\">" +
                        "<span style=\"-aw-import:ignore\">" +
                            "<span>a.</span>" +
                            "<span style=\"width:9.67pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                        "</span>" +
                        "<span>Default numbered list item 3.</span>" +
                    "</p>"));

                Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; -aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; -aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                        "<span style=\"-aw-import:ignore\">" +
                            "<span>2.1.1.1</span>" +
                            "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                        "</span>" +
                        "<span>Outline legal heading list item 5.</span>" +
                    "</p>"));
                break;
            case ExportListLabels.AUTO:
                Assert.assertTrue(outDocContents.contains(
                    "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                        "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                            "<span>Default numbered list item 3.</span>" +
                        "</li>" +
                    "</ol>"));

                Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-left:43.2pt; margin-bottom:0pt; text-indent:-43.2pt; -aw-import:list-item; -aw-list-level-number:3; " +
                    "-aw-list-number-format:'%0.%1.%2.%3'; -aw-list-number-styles:'decimal decimal decimal decimal'; " +
                    "-aw-list-number-values:'2 1 1 1'; -aw-list-padding-sml:10.2pt\">" +
                        "<span style=\"-aw-import:ignore\">" +
                            "<span>2.1.1.1</span>" +
                            "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:spaces\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                        "</span>" +
                        "<span>Outline legal heading list item 5.</span>" +
                    "</p>"));
                break;
            case ExportListLabels.BY_HTML_TAGS:
                Assert.assertTrue(outDocContents.contains(
                    "<ol type=\"a\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                        "<li style=\"margin-left:31.33pt; padding-left:4.67pt\">" +
                            "<span>Default numbered list item 3.</span>" +
                        "</li>" +
                    "</ol>"));

                Assert.assertTrue(outDocContents.contains(
                    "<ol type=\"1\" class=\"awlist3\" style=\"margin-right:0pt; margin-left:0pt; padding-left:0pt\">" +
                        "<li style=\"margin-left:7.2pt; text-indent:-43.2pt; -aw-list-padding-sml:10.2pt\">" +
                            "<span style=\"width:10.2pt; font:7pt 'Times New Roman'; display:inline-block; -aw-import:ignore\">&#xa0;&#xa0;&#xa0;&#xa0;&#xa0;&#xa0; </span>" +
                            "<span>Outline legal heading list item 5.</span>" +
                        "</li>" +
                    "</ol>"));
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "listDataProvider")
	public static Object[][] listDataProvider() throws Exception
	{
		return new Object[][]
		{
			{ExportListLabels.AS_INLINE_TEXT},
			{ExportListLabels.AUTO},
			{ExportListLabels.BY_HTML_TAGS},
		};
	}

    @Test (dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(boolean exportPageMargins) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageMargins
        //ExSummary:Shows how to show out-of-bounds objects in output HTML documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a builder to insert a shape with no wrapping.
        Shape shape = builder.insertShape(ShapeType.CUBE, 200.0, 200.0);

        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setWrapType(WrapType.NONE);

        // Negative shape position values may place the shape outside of page boundaries.
        // If we export this to HTML, the shape will appear truncated.
        shape.setLeft(-150);

        // When saving the document to HTML, we can pass a SaveOptions object
        // to decide whether to adjust the page to display out-of-bounds objects fully.
        // If we set the "ExportPageMargins" flag to "true", the shape will be fully visible in the output HTML.
        // If we set the "ExportPageMargins" flag to "false",
        // our document will display the shape truncated as we would see it in Microsoft Word.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportPageMargins(exportPageMargins); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html");

        if (exportPageMargins)
        {
            Assert.assertTrue(outDocContents.contains("<style type=\"text/css\">div.Section1 { margin:70.85pt }</style>"));
            Assert.assertTrue(outDocContents.contains("<div class=\"Section1\"><p style=\"margin-top:0pt; margin-left:151pt; margin-bottom:0pt\">"));
        }
        else
        {
            Assert.assertFalse(outDocContents.contains("style type=\"text/css\">"));
            Assert.assertTrue(outDocContents.contains("<div><p style=\"margin-top:0pt; margin-left:221.85pt; margin-bottom:0pt\">"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportPageMarginsDataProvider")
	public static Object[][] exportPageMarginsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "exportPageSetupDataProvider")
    public void exportPageSetup(boolean exportPageSetup) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageSetup
        //ExSummary:Shows how decide whether to preserve section structure/page setup information when saving to HTML.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setTopMargin(36.0);
        pageSetup.setBottomMargin(36.0);
        pageSetup.setPaperSize(PaperSize.A5);

        // When saving the document to HTML, we can pass a SaveOptions object
        // to decide whether to preserve or discard page setup settings.
        // If we set the "ExportPageSetup" flag to "true", the output HTML document will contain our page setup configuration.
        // If we set the "ExportPageSetup" flag to "false", the save operation will discard our page setup settings
        // for the first section, and both sections will look identical.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportPageSetup(exportPageSetup); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageSetup.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportPageSetup.html");

        if (exportPageSetup)
        {
            Assert.assertTrue(outDocContents.contains(
                "<style type=\"text/css\">" +
                    "@page Section1 { size:419.55pt 595.3pt; margin:36pt 70.85pt }" +
                    "@page Section2 { size:612pt 792pt; margin:70.85pt }" +
                    "div.Section1 { page:Section1 }div.Section2 { page:Section2 }" +
                "</style>"));

            Assert.assertTrue(outDocContents.contains(
                "<div class=\"Section1\">" +
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<span>Section 1</span>" +
                    "</p>" +
                "</div>"));
        }
        else
        {
            Assert.assertFalse(outDocContents.contains("style type=\"text/css\">"));

            Assert.assertTrue(outDocContents.contains(
                "<div>" +
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<span>Section 1</span>" +
                    "</p>" +
                "</div>"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportPageSetupDataProvider")
	public static Object[][] exportPageSetupDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "relativeFontSizeDataProvider")
    public void relativeFontSize(boolean exportRelativeFontSize) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRelativeFontSize
        //ExSummary:Shows how to use relative font sizes when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Default font size, ");
        builder.getFont().setSize(24.0);
        builder.writeln("2x default font size,");
        builder.getFont().setSize(96.0);
        builder.write("8x default font size");

        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine whether to use relative or absolute font sizes.
        // Set the "ExportRelativeFontSize" flag to "true" to declare font sizes
        // using the "em" measurement unit, which is a factor that multiplies the current font size. 
        // Set the "ExportRelativeFontSize" flag to "false" to declare font sizes
        // using the "pt" measurement unit, which is the font's absolute size in points.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportRelativeFontSize(exportRelativeFontSize); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RelativeFontSize.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.RelativeFontSize.html");

        if (exportRelativeFontSize)
        {
            Assert.assertTrue(outDocContents.contains(
                "<body style=\"font-family:'Times New Roman'\">" +
                    "<div>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Default font size, </span>" +
                        "</p>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:2em\">" +
                            "<span>2x default font size,</span>" +
                        "</p>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:8em\">" +
                            "<span>8x default font size</span>" +
                        "</p>" +
                    "</div>" +
                "</body>"));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains(
                "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                    "<div>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Default font size, </span>" +
                        "</p>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:24pt\">" +
                            "<span>2x default font size,</span>" +
                        "</p>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:96pt\">" +
                            "<span>8x default font size</span>" +
                        "</p>" +
                    "</div>" +
                "</body>"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "relativeFontSizeDataProvider")
	public static Object[][] relativeFontSizeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "exportTextBoxDataProvider")
    public void exportTextBox(boolean exportTextBoxAsSvg) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTextBoxAsSvg
        //ExSummary:Shows how to export text boxes as scalable vector graphics.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 60.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("My text box");

        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation will export text box shapes.
        // If we set the "ExportTextBoxAsSvg" flag to "true",
        // the save operation will convert shapes with text into SVG objects.
        // If we set the "ExportTextBoxAsSvg" flag to "false",
        // the save operation will convert shapes with text into images.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportTextBoxAsSvg(exportTextBoxAsSvg); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTextBox.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportTextBox.html");

        if (exportTextBoxAsSvg)
        {
            Assert.assertTrue(outDocContents.contains(
                "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">"));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains(
                "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                    "<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
                    "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                "</p>"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportTextBoxDataProvider")
	public static Object[][] exportTextBoxDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "roundTripInformationDataProvider")
    public void roundTripInformation(boolean exportRoundtripInformation) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRoundtripInformation
        //ExSummary:Shows how to preserve hidden elements when converting to .html.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
        // or footnotes will be either removed or converted to plain text and effectively be lost.
        // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements.

        // When we save the document to HTML, we can pass a SaveOptions object to determine
        // how the saving operation will export document elements that HTML does not support or use,
        // such as hidden bookmarks and original shape positions.
        // If we set the "ExportRoundtripInformation" flag to "true", the save operation will preserve these elements.
        // If we set the "ExportRoundTripInformation" flag to "false", the save operation will discard these elements.
        // We will want to preserve such elements if we intend to load the saved HTML using Aspose.Words,
        // as they could be of use once again.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportRoundtripInformation(exportRoundtripInformation); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html");
        doc = new Document(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html");

        if (exportRoundtripInformation)
        {
            Assert.assertTrue(outDocContents.contains("<div style=\"-aw-headerfooter-type:header-primary; clear:both\">"));
            Assert.assertTrue(outDocContents.contains("<span style=\"-aw-import:ignore\">&#xa0;</span>"));
            
            Assert.assertTrue(outDocContents.contains(
                "td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " +
                "-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">"));
            
            Assert.assertTrue(outDocContents.contains(
                "<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">"));
            
            Assert.assertTrue(outDocContents.contains(
                "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" " +
                "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />"));


            Assert.assertTrue(outDocContents.contains(
                "<span>Page number </span>" +
                "<span style=\"-aw-field-start:true\"></span>" +
                "<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                "<span style=\"-aw-field-separator:true\"></span>" +
                "<span>1</span>" +
                "<span style=\"-aw-field-end:true\"></span>"));

            Assert.AreEqual(1, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldPage));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains("<div style=\"clear:both\">"));
            Assert.assertTrue(outDocContents.contains("<span>&#xa0;</span>"));
            
            Assert.assertTrue(outDocContents.contains(
                "<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">"));
            
            Assert.assertTrue(outDocContents.contains(
                "<li style=\"margin-left:30.2pt; padding-left:5.8pt\">"));
            
            Assert.assertTrue(outDocContents.contains(
                "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"350\" height=\"180\" alt=\"\" />"));

            Assert.assertTrue(outDocContents.contains(
                "<span>Page number 1</span>"));

            Assert.AreEqual(0, doc.getRange().getFields().Count(f => f.Type == FieldType.FieldPage));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "roundTripInformationDataProvider")
	public static Object[][] roundTripInformationDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "exportTocPageNumbersDataProvider")
    public void exportTocPageNumbers(boolean exportTocPageNumbers) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTocPageNumbers
        //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
        // style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
        // and the page number that contains the heading on the right.
        FieldToc fieldToc = (FieldToc)builder.insertField(FieldType.FIELD_TOC, true);

        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 1"));
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Entry 1");
        builder.writeln("Entry 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Entry 3");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Entry 4");
        fieldToc.updatePageNumbers();
        doc.updateFields();

        // HTML documents do not have pages. If we save this document to HTML,
        // the page numbers that our TOC displays will have no meaning.
        // When we save the document to HTML, we can pass a SaveOptions object to omit these page numbers from the TOC.
        // If we set the "ExportTocPageNumbers" flag to "true",
        // each TOC entry will display the heading, separator, and page number, preserving its appearance in Microsoft Word.
        // If we set the "ExportTocPageNumbers" flag to "false",
        // the save operation will omit both the separator and page number and leave the heading for each entry intact.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setExportTocPageNumbers(exportTocPageNumbers); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTocPageNumbers.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ExportTocPageNumbers.html");

        if (exportTocPageNumbers)
        {
            Assert.assertTrue(outDocContents.contains(
                "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                "<span>Entry 1</span>" +
                "<span style=\"width:428.14pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                "-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:469.8pt\">.......................................................................</span>" +
                "<span>2</span>" +
                "</p>"));
        }
        else
        {
            Assert.assertTrue(outDocContents.contains(
                "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                "<span>Entry 1</span>" +
                "</p>"));
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportTocPageNumbersDataProvider")
	public static Object[][] exportTocPageNumbersDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "fontSubsettingDataProvider")
    public void fontSubsetting(int fontResourcesSubsettingSizeThreshold) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
        //ExSummary:Shows how to work with font subsetting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Times New Roman");
        builder.writeln("Hello world!");
        builder.getFont().setName("Courier New");
        builder.writeln("Hello world!");

        // When we save the document to HTML, we can pass a SaveOptions object configure font subsetting.
        // Suppose we set the "ExportFontResources" flag to "true" and also name a folder in the "FontsFolder" property.
        // In that case, the saving operation will create that folder and place a .ttf file inside
        // that folder for each font that our document uses.
        // Each .ttf file will contain that font's entire glyph set,
        // which may potentially result in a very large file that accompanies the document.
        // When we apply subsetting to a font, its exported raw data will only contain the glyphs that the document is
        // using instead of the entire glyph set. If the text in our document only uses a small fraction of a font's
        // glyph set, then subsetting will significantly reduce our output documents' size.
        // We can use the "FontResourcesSubsettingSizeThreshold" property to define a .ttf file size, in bytes.
        // If an exported font creates a size bigger file than that, then the save operation will apply subsetting to that font. 
        // Setting a threshold of 0 applies subsetting to all fonts,
        // and setting it to "int.MaxValue" effectively disables subsetting.
        String fontsFolder = getArtifactsDir() + "HtmlSaveOptions.FontSubsetting.Fonts";

        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportFontResources(true);
            options.setFontsFolder(fontsFolder);
            options.setFontResourcesSubsettingSizeThreshold(fontResourcesSubsettingSizeThreshold);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.FontSubsetting.html", options);

        String[] fontFileNames = Directory.getFiles(fontsFolder).Where(s => s.EndsWith(".ttf")).ToArray();

        Assert.assertEquals(3, fontFileNames.length);

        for (String filename : fontFileNames)
        {
            // By default, the .ttf files for each of our three fonts will be over 700MB.
            // Subsetting will reduce them all to under 30MB.
            FileInfo fontFileInfo = new FileInfo(filename);

            Assert.assertTrue(fontFileInfo.getLength() > 700000 || fontFileInfo.getLength() < 30000);
            Assert.assertTrue(Math.max(fontResourcesSubsettingSizeThreshold, 30000) > new FileInfo(filename).getLength());
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "fontSubsettingDataProvider")
	public static Object[][] fontSubsettingDataProvider() throws Exception
	{
		return new Object[][]
		{
			{0},
			{1000000},
			{Integer.MAX_VALUE},
		};
	}

    @Test (dataProvider = "metafileFormatDataProvider")
    public void metafileFormat(/*HtmlMetafileFormat*/int htmlMetafileFormat) throws Exception
    {
        //ExStart
        //ExFor:HtmlMetafileFormat
        //ExFor:HtmlSaveOptions.MetafileFormat
        //ExFor:HtmlLoadOptions.ConvertSvgToEmf
        //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
        String html = 
            "<html>\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\n                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\n                    </svg>\n                </html>";

        // Use 'ConvertSvgToEmf' to turn back the legacy behavior
        // where all SVG images loaded from an HTML document were converted to EMF.
        // Now SVG images are loaded without conversion
        // if the MS Word version specified in load options supports SVG images natively.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(); { loadOptions.setConvertSvgToEmf(true); }

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(html)), loadOptions);

        // This document contains a <svg> element in the form of text.
        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation handles this object.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Png" to convert it to a PNG image.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Svg" preserve it as a SVG object.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.EmfOrWmf" to convert it to a metafile.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setMetafileFormat(htmlMetafileFormat); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html", options);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html");

        switch (htmlMetafileFormat)
        {
            case HtmlMetafileFormat.PNG:
                Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<img src=\"HtmlSaveOptions.MetafileFormat.001.png\" width=\"500\" height=\"40\" alt=\"\" " +
                        "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                    "</p>"));
                break;
            case HtmlMetafileFormat.SVG:
                Assert.assertTrue(outDocContents.contains(
                    "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"499\" height=\"40\">"));
                break;
            case HtmlMetafileFormat.EMF_OR_WMF:
                Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                        "<img src=\"HtmlSaveOptions.MetafileFormat.001.emf\" width=\"500\" height=\"40\" alt=\"\" " +
                        "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                    "</p>"));
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "metafileFormatDataProvider")
	public static Object[][] metafileFormatDataProvider() throws Exception
	{
		return new Object[][]
		{
			{HtmlMetafileFormat.PNG},
			{HtmlMetafileFormat.SVG},
			{HtmlMetafileFormat.EMF_OR_WMF},
		};
	}

    @Test (dataProvider = "officeMathOutputModeDataProvider")
    public void officeMathOutputMode(/*HtmlOfficeMathOutputMode*/int htmlOfficeMathOutputMode) throws Exception
    {
        //ExStart
        //ExFor:HtmlOfficeMathOutputMode
        //ExFor:HtmlSaveOptions.OfficeMathOutputMode
        //ExSummary:Shows how to specify how to export Microsoft OfficeMath objects to HTML.
        Document doc = new Document(getMyDir() + "Office math.docx");

        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation handles OfficeMath objects.
        // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Image"
        // will render each OfficeMath object into an image.
        // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.MathML"
        // will convert each OfficeMath object into MathML.
        // Setting the "OfficeMathOutputMode" property to "HtmlOfficeMathOutputMode.Text"
        // will represent each OfficeMath formula using plain HTML text.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setOfficeMathOutputMode(htmlOfficeMathOutputMode); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.OfficeMathOutputMode.html", options);
        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.OfficeMathOutputMode.html");

        switch (htmlOfficeMathOutputMode)
        {
            case HtmlOfficeMathOutputMode.IMAGE:
                Assert.assertTrue(Regex.match(outDocContents, 
                    "<p style=\"margin-top:0pt; margin-bottom:10pt\">" +
                        "<img src=\"HtmlSaveOptions.OfficeMathOutputMode.001.png\" width=\"159\" height=\"19\" alt=\"\" style=\"vertical-align:middle; " +
                        "-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                    "</p>").getSuccess());
                break;
            case HtmlOfficeMathOutputMode.MATH_ML:
                Assert.assertTrue(Regex.match(outDocContents,
                    "<p style=\"margin-top:0pt; margin-bottom:10pt; text-align:center\">" +
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" +
                            "<mi>i</mi>" +
                            "<mo>[+]</mo>" +
                            "<mi>b</mi>" +
                            "<mo>-</mo>" +
                            "<mi>c</mi>" +
                            "<mo>≥</mo>" +
                            ".*" +
                        "</math>" +
                    "</p>").getSuccess());
                break;
            case HtmlOfficeMathOutputMode.TEXT:
                Assert.assertTrue(Regex.match(outDocContents,
                    "<p style=\\\"margin-top:0pt; margin-bottom:10pt; text-align:center\\\">" +
                        "<span style=\\\"font-family:'Cambria Math'\\\">i[+]b-c≥iM[+]bM-cM </span>" +
                    "</p>").getSuccess());
                break;
        }
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "officeMathOutputModeDataProvider")
	public static Object[][] officeMathOutputModeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{HtmlOfficeMathOutputMode.IMAGE},
			{HtmlOfficeMathOutputMode.MATH_ML},
			{HtmlOfficeMathOutputMode.TEXT},
		};
	}

    @Test (dataProvider = "scaleImageToShapeSizeDataProvider")
    public void scaleImageToShapeSize(boolean scaleImageToShapeSize) throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
        //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape which contains an image, and then make that shape considerably smaller than the image.
        BufferedImage image = ImageIO.read(getImageDir() + "Transparent background logo.png");

        Assert.assertEquals(400, msSize.getWidth(image.Size));
        Assert.assertEquals(400, msSize.getHeight(image.Size));

        Shape imageShape = builder.insertImage(image);
        imageShape.setWidth(50.0);
        imageShape.setHeight(50.0);

        // Saving a document that contains shapes with images to HTML will create an image file in the local file system
        // for each such shape. The output HTML document will use <image> tags to link to and display these images.
        // When we save the document to HTML, we can pass a SaveOptions object to determine
        // whether to scale all images that are inside shapes to the sizes of their shapes.
        // Setting the "ScaleImageToShapeSize" flag to "true" will shrink every image
        // to the size of the shape that contains it, so that no saved images will be larger than the document requires them to be.
        // Setting the "ScaleImageToShapeSize" flag to "false" will preserve these images' original sizes,
        // which will take up more space in exchange for preserving image quality.
        HtmlSaveOptions options = new HtmlSaveOptions(); { options.setScaleImageToShapeSize(scaleImageToShapeSize); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ScaleImageToShapeSize.html", options);

        FileInfo fileInfo = new FileInfo(getArtifactsDir() + "HtmlSaveOptions.ScaleImageToShapeSize.001.png");

    if (scaleImageToShapeSize)
        Assert.That(3000, Is.AtLeast(fileInfo.getLength()));
    else
        Assert.That(20000, Is.LessThan(fileInfo.getLength()));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "scaleImageToShapeSizeDataProvider")
	public static Object[][] scaleImageToShapeSizeDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test
    public void imageFolder() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
        //ExFor:HtmlSaveOptions.ImagesFolder
        //ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        String imagesDir = Path.combine(getArtifactsDir(), "SaveHtmlWithOptions");

        if (Directory.exists(imagesDir))
            Directory.delete(imagesDir, true);

        Directory.createDirectory(imagesDir);

        // Set an option to export form fields as plain text instead of HTML input elements.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        {
            options.setExportTextInputFormFieldAsText(true); 
            options.setImagesFolder(imagesDir);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.SaveHtmlWithOptions.html", options);
        //ExEnd

        Assert.assertTrue(File.exists(getArtifactsDir() + "HtmlSaveOptions.SaveHtmlWithOptions.html"));
        Assert.assertEquals(9, Directory.getFiles(imagesDir).length);

        Directory.delete(imagesDir, true);
    }

    //ExStart
    //ExFor:ImageSavingArgs.CurrentShape
    //ExFor:ImageSavingArgs.Document
    //ExFor:ImageSavingArgs.ImageStream
    //ExFor:ImageSavingArgs.IsImageAvailable
    //ExFor:ImageSavingArgs.KeepImageStreamOpen
    //ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
    @Test //ExSkip
    public void imageSavingCallback() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When we save the document to HTML, we can pass a SaveOptions object to designate a callback
        // to customize the image saving process.
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setImageSavingCallback(new ImageShapePrinter());
       
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ImageSavingCallback.html", options);
    }

    /// <summary>
    /// Prints the properties of each image as the saving process saves it to an image file in the local file system
    /// during the exporting of a document to HTML.
    /// </summary>
    private static class ImageShapePrinter implements IImageSavingCallback
    {
        public void /*IImageSavingCallback.*/imageSaving(ImageSavingArgs args)
        {
            args.setKeepImageStreamOpen(false);
            Assert.assertTrue(args.isImageAvailable());

            System.out.println("{args.Document.OriginalFileName.Split('\\').Last()} Image #{++mImageCount}");

            LayoutCollector layoutCollector = new LayoutCollector(args.getDocument());

            System.out.println("\tOn page:\t{layoutCollector.GetStartPageIndex(args.CurrentShape)}");
            System.out.println("\tDimensions:\t{args.CurrentShape.Bounds}");
            System.out.println("\tAlignment:\t{args.CurrentShape.VerticalAlignment}");
            System.out.println("\tWrap type:\t{args.CurrentShape.WrapType}");
            System.out.println("Output filename:\t{args.ImageFileName}\n");
        }

        private int mImageCount;
    }
    //ExEnd

    @Test (dataProvider = "prettyFormatDataProvider")
    public void prettyFormat(boolean usePrettyFormat) throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.HTML); { htmlOptions.setPrettyFormat(usePrettyFormat); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.PrettyFormat.html", htmlOptions);

        // Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
        String html = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.PrettyFormat.html");

        if (usePrettyFormat)
            Assert.assertEquals(
                "<html>\r\n" +
                            "\t<head>\r\n" +
                                "\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n" +
                                "\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />\r\n" +
                                $"\t\t<meta name=\"generator\" content=\"{BuildVersionInfo.Product} {BuildVersionInfo.Version}\" />\r\n" +
                                "\t\t<title>\r\n" +
                                "\t\t</title>\r\n" +
                            "\t</head>\r\n" +
                            "\t<body style=\"font-family:'Times New Roman'; font-size:12pt\">\r\n" +
                                "\t\t<div>\r\n" +
                                    "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                                        "\t\t\t\t<span>Hello world!</span>\r\n" +
                                    "\t\t\t</p>\r\n" +
                                    "\t\t\t<p style=\"margin-top:0pt; margin-bottom:0pt\">\r\n" +
                                        "\t\t\t\t<span style=\"-aw-import:ignore\">&#xa0;</span>\r\n" +
                                    "\t\t\t</p>\r\n" +
                                "\t\t</div>\r\n" +
                            "\t</body>\r\n</html>", 
                html);
        else
            Assert.assertEquals(
                "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                        "<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />" +
                        $"<meta name=\"generator\" content=\"{BuildVersionInfo.Product} {BuildVersionInfo.Version}\" /><title></title></head>" +
                        "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                        "<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
                        "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>", 
                html);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "prettyFormatDataProvider")
	public static Object[][] prettyFormatDataProvider() throws Exception
	{
		return new Object[][]
		{
			{true},
			{false},
		};
	}
}
