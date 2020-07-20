// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.HtmlMetafileFormat;
import com.aspose.words.FontSettings;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.DocumentSplitCriteria;
import com.aspose.words.Table;
import com.aspose.words.PreferredWidth;
import com.aspose.words.HtmlElementSizeOutputMode;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;
import com.aspose.words.BreakType;
import com.aspose.words.PageSetup;
import com.aspose.words.PaperSize;
import com.aspose.words.FieldToc;
import com.aspose.words.FieldType;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.ms.System.Text.Encoding;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.ms.System.msConsole;
import com.aspose.words.LayoutCollector;
import org.testng.annotations.DataProvider;


@Test
class ExHtmlSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(/*SaveFormat*/int saveFormat) throws Exception
    {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setSaveFormat(saveFormat);
            saveOptions.setExportPageMargins(true);
        }

        doc.save(getArtifactsDir() +"HtmlSaveOptions.ExportPageMargins" + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportPageMarginsDataProvider")
	public static Object[][] exportPageMarginsDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.HTML},
			{SaveFormat.MHTML},
			{SaveFormat.EPUB},
		};
	}

    @Test (dataProvider = "exportOfficeMathDataProvider")
    public void exportOfficeMath(/*SaveFormat*/int saveFormat, /*HtmlOfficeMathOutputMode*/int outputMode) throws Exception
    {
        Document doc = new Document(getMyDir() + "Office math.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setOfficeMathOutputMode(outputMode);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportOfficeMath" + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportOfficeMathDataProvider")
	public static Object[][] exportOfficeMathDataProvider() throws Exception
	{
		return new Object[][]
		{
			{SaveFormat.HTML,  HtmlOfficeMathOutputMode.IMAGE},
			{SaveFormat.MHTML,  HtmlOfficeMathOutputMode.MATH_ML},
			{SaveFormat.EPUB,  HtmlOfficeMathOutputMode.TEXT},
		};
	}

    @Test (dataProvider = "exportTextBoxAsSvgDataProvider")
    public void exportTextBoxAsSvg(/*SaveFormat*/int saveFormat, boolean isTextBoxAsSvg) throws Exception
    {
        String[] dirFiles;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textbox = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 100.0);
        builder.moveTo(textbox.getFirstParagraph());
        builder.write("Hello world!");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat);
        saveOptions.setExportTextBoxAsSvg(isTextBoxAsSvg);
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTextBoxAsSvg" + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);

        switch (saveFormat)
        {
            case SaveFormat.HTML:
                
                dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.ALL_DIRECTORIES);
                Assert.That(dirFiles, Is.Empty);
                return;

            case SaveFormat.EPUB:

                dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.ALL_DIRECTORIES);
                Assert.That(dirFiles, Is.Empty);
                return;

            case SaveFormat.MHTML:

                dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvg.001.png", SearchOption.ALL_DIRECTORIES);
                Assert.That(dirFiles, Is.Empty);
                return;
        }
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "exportTextBoxAsSvgDataProvider")
	public static Object[][] exportTextBoxAsSvgDataProvider() throws Exception
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
            // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
            // otherwise HTML <p> tag is used. This is also the default value
            // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation
            // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible
            saveOptions.setExportListLabels(howExportListLabels);
        }

        doc.save(getArtifactsDir() + $"HtmlSaveOptions.ControlListLabelsExport.html", saveOptions);
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

        String[] dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.ALL_DIRECTORIES);

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
        //Assert that default value is true for HTML and false for MHTML and EPUB
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

        String[] imageFiles = Directory.getFiles(getArtifactsDir() + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.png", SearchOption.ALL_DIRECTORIES);
        Assert.assertEquals(8, imageFiles.length);

        String[] fontFiles = Directory.getFiles(getArtifactsDir() + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.ttf", SearchOption.ALL_DIRECTORIES);
        Assert.assertEquals(10, fontFiles.length);

        String[] cssFiles = Directory.getFiles(getArtifactsDir() + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.css", SearchOption.ALL_DIRECTORIES);
        Assert.assertEquals(1, cssFiles.length);

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExternalResourceSavingConfig.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
    }

    @Test
    public void convertFontsAsBase64() throws Exception
    {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(true);
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
	}

    @Test (dataProvider = "html5SupportDataProvider")
    public void html5Support(/*HtmlVersion*/int htmlVersion) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setHtmlVersion(htmlVersion);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.Html5Support.html", saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "html5SupportDataProvider")
	public static Object[][] html5SupportDataProvider() throws Exception
	{
		return new Object[][]
		{
			{com.aspose.words.HtmlVersion.HTML_5},
			{com.aspose.words.HtmlVersion.XHTML},
		};
	}

    @Test (dataProvider = "exportFontsDataProvider")
    public void exportFonts(boolean exportAsBase64) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setExportFontResources(true);
            saveOptions.setExportFontsAsBase64(exportAsBase64);
        }

        switch (exportAsBase64)
        {
            case false:

                doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFonts.False.html", saveOptions);
                Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportFonts.False.times.ttf",
                    SearchOption.ALL_DIRECTORIES));
                break;

            case true:

                doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFonts.True.html", saveOptions);
                msAssert.isEmpty(Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportFonts.True.times.ttf",
                    SearchOption.ALL_DIRECTORIES));
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
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

        String[] a = Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.001.jpeg",
            SearchOption.ALL_DIRECTORIES);
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
            "<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.SvgMetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.PNG); });
    }

    @Test
    public void pngMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an Png image: ");
        builder.insertHtml(
            "<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.PngMetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.PNG); });
    }

    @Test
    public void emfOrWmfMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an image as is: ");
        builder.insertHtml(
            "<img src=\"data:image/png;base64,\r\n                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\r\n                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\r\n                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\r\n                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\r\n                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\r\n                    vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.EmfOrWmfMetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF); });
    }

    @Test
    public void cssClassNamesPrefix() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.CssClassNamePrefix
        //ExSummary:Shows how to specifies a prefix which is added to all CSS class names.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
            saveOptions.setCssClassNamePrefix("myprefix-");
        }

        // The prefix will be found before CSS element names in the embedded stylesheet
        doc.save(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
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

    @Test (enabled = false, description = "Bug")
    public void resolveFontNames() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ResolveFontNames
        //ExSummary:Shows how to resolve all font names before writing them to HTML.
        Document document = new Document(getMyDir() + "Missing font.docx");

        FontSettings fontSettings = new FontSettings();
        {
            fontSettings.setSubstitutionSettings({
                fontSettings.getSubstitutionSettings().setDefaultFontSubstitution({
                    fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
                    fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setEnabled(true);
                });
            });
        }

        document.setFontSettings(fontSettings);
        
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            // By default this option is set to 'False' and Aspose.Words writes font names as specified in the source document
            saveOptions.setResolveFontNames(true); 
        }

        document.save(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);

        String outDocContents = File.readAllText(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html");

        Assert.assertTrue(Regex.match(outDocContents, "<span style=\"font-family:Arial\">").getSuccess());
        //ExEnd
    }

    @Test
    public void headingLevels() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.DocumentSplitHeadingLevel
        //ExSummary:Shows how to split a document into several html documents by heading levels.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings of levels 1 - 3
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

        // Create a HtmlSaveOptions object and set the split criteria to "HeadingParagraph", meaning that the document 
        // will be split into parts at the beginning of every paragraph of a "Heading" style, and each part will be saved as a separate document
        // Also, we will set the DocumentSplitHeadingLevel to 2, which will split the document only at headings that have levels from 1 to 2
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);
            options.setDocumentSplitHeadingLevel(2);
        }
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.HeadingLevels.html", options);
        //ExEnd
    }

    @Test
    public void negativeIndent() throws Exception
    {
        //ExStart
        //ExFor:HtmlElementSizeOutputMode
        //ExFor:HtmlSaveOptions.AllowNegativeIndent
        //ExFor:HtmlSaveOptions.TableWidthOutputMode
        //ExSummary:Shows how to preserve negative indents in the output .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table and give it a negative value for its indent, effectively pushing it out of the left page boundary
        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Cell 1");
        builder.insertCell();
        builder.write("Cell 2");
        builder.endTable();
        table.setLeftIndent(-36);
        table.setPreferredWidth(PreferredWidth.fromPoints(144.0));

        // When saving to .html, this indent will only be preserved if we set this flag
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setAllowNegativeIndent(true);
        options.setTableWidthOutputMode(HtmlElementSizeOutputMode.RELATIVE_ONLY);

        // The first cell with "Cell 1" will not be visible in the output 
        doc.save(getArtifactsDir() + "HtmlSaveOptions.NegativeIndent.html", options);
        //ExEnd
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
        //ExSummary:Shows how to set folders and folder aliases for externally saved resources when saving to html.
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

    @Test
    public void htmlVersion() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.#ctor(SaveFormat)
        //ExFor:HtmlSaveOptions.ExportXhtmlTransitional
        //ExFor:HtmlSaveOptions.HtmlVersion
        //ExFor:HtmlVersion
        //ExSummary:Shows how to set a saved .html document to a specific version.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Save the document to a .html file of the XHTML 1.0 Transitional standard
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        {
            options.setHtmlVersion(com.aspose.words.HtmlVersion.XHTML);
            options.setExportXhtmlTransitional(true);
            options.setPrettyFormat(true);
        }

        // The DOCTYPE declaration at the top of this document will indicate the html version we chose
        doc.save(getArtifactsDir() + "HtmlSaveOptions.HtmlVersion.html", options);
        //ExEnd
    }

    @Test
    public void epubHeadings() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.EpubNavigationMapLevel
        //ExSummary:Shows the relationship between heading levels and the Epub navigation panel.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert headings of levels 1 - 3
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

        // Epub readers normally treat paragraphs with "Heading" styles as anchors for a table of contents-style navigation pane
        // We set a maximum heading level above which headings won't be registered by the reader as navigation points with
        // a HtmlSaveOptions object and its EpubNavigationLevel attribute
        // Our document has headings of levels 1 to 3,
        // but our output epub will only place level 1 and 2 headings in the table of contents
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.EPUB);
        options.setEpubNavigationMapLevel(2);
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.EpubHeadings.epub", options);
        //ExEnd
    }

    @Test
    public void contentIdUrls() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportCidUrlsForMhtmlResources
        //ExSummary:Shows how to enable content IDs for output MHTML documents.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Setting this flag will replace "Content-Location" tags with "Content-ID" tags for each resource from the input document
        // The file names that were next to each "Content-Location" tag are re-purposed as content IDs
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.MHTML);
        {
            options.setExportCidUrlsForMhtmlResources(true);
            options.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            options.setExportFontResources(true);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ContentIdUrls.mht", options);
        //ExEnd
    }

    @Test
    public void dropDownFormField() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
        //ExSummary:Shows how to get drop down combo box form fields to blend in with paragraph text when saving to html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a combo box with the value "Two" selected
        builder.insertComboBox("MyComboBox", new String[] { "One", "Two", "Three" }, 1);
        
        // When converting to .html, drop down combo boxes will be converted to select/option tags to preserve their functionality
        // If we want to freeze a combo box at its current selected value and convert it into plain text, we can do so with this flag
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportDropDownFormFieldAsText(true);    

        doc.save(getArtifactsDir() + "HtmlSaveOptions.DropDownFormField.html", options);
        //ExEnd
    }

    @Test
    public void exportBase64() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportFontsAsBase64
        //ExFor:HtmlSaveOptions.ExportImagesAsBase64
        //ExSummary:Shows how to save a .html document with resources embedded inside it.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // By default, when converting a document with images to .html, resources such as images will be linked to in external files
        // We can set these flags to embed resources inside the output .html instead, cutting down on the amount of files created during the conversion
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportFontsAsBase64(true);
            options.setExportImagesAsBase64(true);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportBase64.html", options);
        //ExEnd
    }

    @Test
    public void exportLanguageInformation() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportLanguageInformation
        //ExSummary:Shows how to preserve language information when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use the builder to write text in more than one language
        builder.getFont().setLocaleId(2057); // en-GB
        builder.writeln("Hello world!");

        builder.getFont().setLocaleId(1049); // ru-RU
        builder.write("Привет, мир!");

        // Normally, when saving a document with more than one proofing language to .html,
        // only the text content is preserved with no traces of any other languages
        // Saving with a HtmlSaveOptions object with this flag set will add "lang" attributes to spans 
        // in places where other proofing languages were used 
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportLanguageInformation(true);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportLanguageInformation.html", options);
        //ExEnd
    }

    @Test
    public void list() throws Exception
    {
        //ExStart
        //ExFor:ExportListLabels
        //ExFor:HtmlSaveOptions.ExportListLabels
        //ExSummary:Shows how to export an indented list to .html as plain text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use the builder to insert a list
        List list = doc.getLists().add(ListTemplate.NUMBER_DEFAULT);
        builder.getListFormat().setList(list);
        
        builder.writeln("List item 1.");
        builder.getListFormat().listIndent();
        builder.writeln("List item 2.");
        builder.getListFormat().listIndent();
        builder.write("List item 3.");

        // When we save this to .html, normally our list will be represented by <li> tags
        // We can set this flag to have lists as plain text instead
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.List.html", options);
        //ExEnd
    }

    @Test
    public void exportPageMargins() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageMargins
        //ExSummary:Shows how to show out-of-bounds objects in output .html documents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a builder to insert a shape with no wrapping
        Shape shape = builder.insertShape(ShapeType.CUBE, 200.0, 200.0);

        shape.setRelativeHorizontalPosition(RelativeHorizontalPosition.PAGE);
        shape.setRelativeVerticalPosition(RelativeVerticalPosition.PAGE);
        shape.setWrapType(WrapType.NONE);

        // Negative values for shape position may cause the shape to go out of page bounds
        // If we export this to .html, the shape will be truncated
        shape.setLeft(-150);

        // We can avoid that and have the entire shape be visible by setting this flag
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportPageMargins(true);
    
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html", options);
        //ExEnd
    }

    @Test
    public void exportPageSetup() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportPageSetup
        //ExSummary:Shows how to preserve section structure/page setup information when saving to html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a DocumentBuilder to insert two sections with text
        builder.writeln("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.writeln("Section 2");

        // Change dimensions and paper size of first section
        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setTopMargin(36.0);
        pageSetup.setBottomMargin(36.0);
        pageSetup.setPaperSize(PaperSize.A5);

        // Section structure and pagination are normally lost when when converting to .html
        // We can create an HtmlSaveOptions object with the ExportPageSetup flag set to true
        // to preserve the section structure in <div> tags and page dimensions in the output document's CSS
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportPageSetup(true);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageSetup.html", options);
        //ExEnd
    }

    @Test
    public void relativeFontSize() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRelativeFontSize
        //ExSummary:Shows how to use relative font sizes when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a builder to write some text in various sizes
        builder.writeln("Default font size, ");
        builder.getFont().setSize(24.0);
        builder.writeln("2x default font size,");
        builder.getFont().setSize(96.0);
        builder.write("8x default font size");

        // We can save font sizes as ratios of the default size, which will be 12 in this case
        // If we use an input .html, this size can be set with the AbsSize {font-size:12pt} tag
        // The ExportRelativeFontSize will enable this feature
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportRelativeFontSize(true);
            options.setPrettyFormat(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RelativeFontSize.html", options);
        //ExEnd
    }

    @Test
    public void exportTextBox() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTextBoxAsSvg
        //ExSummary:Shows how to export text boxes as scalable vector graphics.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a DocumentBuilder to insert a text box and give it some text content
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 60.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("My text box");

        // Normally, all shapes such as the text box we placed are exported to .html as external images linked by the .html document
        // We can save with an HtmlSaveOptions object with the ExportTextBoxAsSvg set to true to save text boxes as <svg> tags,
        // which will cause no linked images to be saved and will make the inner text selectable
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportTextBoxAsSvg(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTextBox.html", options);
        //ExEnd
    }

    @Test
    public void roundTripInformation() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportRoundtripInformation
        //ExSummary:Shows how to preserve hidden elements when converting to .html.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // When converting a document to .html, some elements such as hidden bookmarks, original shape positions,
        // or footnotes will be either removed or converted to plain text and effectively be lost
        // Saving with a HtmlSaveOptions object with ExportRoundtripInformation set to true will preserve these elements
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportRoundtripInformation(true);
            options.setPrettyFormat(true);
        }

        // These elements will have tags that will start with "-aw", such as "-aw-import" or "-aw-left-pos"
        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html", options);
        //ExEnd
    }

    @Test
    public void exportTocPageNumbers() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTocPageNumbers
        //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents
        FieldToc fieldToc = (FieldToc)builder.insertField(FieldType.FIELD_TOC, true);

        // Populate the document with paragraphs of a "Heading" style that the table of contents will pick up
        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 1"));
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Entry 1");
        builder.writeln("Entry 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Entry 3");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("Entry 4");

        // Our headings span several pages, and those page numbers will be displayed by the TOC at the top of the document
        fieldToc.updatePageNumbers();
        doc.updateFields();

        // These page numbers are normally omitted since .html has no pagination, but we can still have them displayed
        // if we save with a HtmlSaveOptions object with the ExportTocPageNumbers set to true 
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportTocPageNumbers(true);
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTocPageNumbers.html", options);
        //ExEnd
    }

    @Test
    public void fontSubsetting() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.FontResourcesSubsettingSizeThreshold
        //ExSummary:Shows how to work with font subsetting.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a DocumentBuilder to insert text with several fonts
        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Times New Roman");
        builder.writeln("Hello world!");
        builder.getFont().setName("Courier New");
        builder.writeln("Hello world!");

        // When saving to .html, font subsetting fully applies by default, meaning that when we export fonts with our file,
        // the symbols not used by our document are not represented by the exported fonts, which cuts down file size dramatically
        // Font files of a file size larger than FontResourcesSubsettingSizeThreshold get subsetted, so a value of 0 will apply default full subsetting
        // Setting the value to something large will fully suppress subsetting, saving some very large font files that cover every glyph
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportFontResources(true);
            options.setFontResourcesSubsettingSizeThreshold(Integer.MAX_VALUE);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.FontSubsetting.html", options);
        //ExEnd
    }

    @Test
    public void metafileFormat() throws Exception
    {
        //ExStart
        //ExFor:HtmlMetafileFormat
        //ExFor:HtmlSaveOptions.MetafileFormat
        //ExSummary:Shows how to set a meta file in a different format.
        // Create a document from an html string
        String html = 
            "<html>\r\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\r\n                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\r\n                    </svg>\r\n                </html>";

        Document doc = new Document(new MemoryStream(Encoding.getUTF8().getBytes(html)));

        // This document contains a <svg> element in the form of text,
        // which by default will be saved as a linked external .png when we save the document as html
        // We can save with a HtmlSaveOptions object with this flag set to preserve the <svg> tag
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setMetafileFormat(HtmlMetafileFormat.SVG);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html", options);
        //ExEnd
    }

    @Test
    public void officeMathOutputMode() throws Exception
    {
        //ExStart
        //ExFor:HtmlOfficeMathOutputMode
        //ExFor:HtmlSaveOptions.OfficeMathOutputMode
        //ExSummary:Shows how to control the way how OfficeMath objects are exported to .html.
        // Open a document that contains OfficeMath objects
        Document doc = new Document(getMyDir() + "Office math.docx");

        // Create a HtmlSaveOptions object and configure it to export OfficeMath objects as images
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setOfficeMathOutputMode(HtmlOfficeMathOutputMode.IMAGE);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.OfficeMathOutputMode.html", options);
        //ExEnd
    }

    @Test
    public void scaleImageToShapeSize() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
        //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
        // Open a document which contains shapes with images
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // By default, images inside shapes get scaled to the size of their shapes while the document gets 
        // converted to .html, reducing image file size
        // We can save the document with a HtmlSaveOptions with ScaleImageToShapeSize set to false to prevent the scaling
        // and preserve the full quality and file size of the linked images
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setScaleImageToShapeSize(false);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ScaleImageToShapeSize.html", options);
        //ExEnd
    }

    //ExStart
    //ExFor:ImageSavingArgs.CurrentShape
    //ExFor:ImageSavingArgs.Document
    //ExFor:ImageSavingArgs.ImageStream
    //ExFor:ImageSavingArgs.IsImageAvailable
    //ExFor:ImageSavingArgs.KeepImageStreamOpen
    //ExSummary:Shows how to involve an image saving callback in an .html conversion process.
    @Test //ExSkip
    public void imageSavingCallback() throws Exception
    {
        // Open a document which contains shapes with images
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Create a HtmlSaveOptions object with a custom image saving callback that will print image information
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setImageSavingCallback(new ImageShapePrinter());
       
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ImageSavingCallback.html", options);
    }

    /// <summary>
    /// Prints information of all images that are about to be saved from within a document to image files
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
            System.out.println("\tDimensions:\t{args.CurrentShape.Bounds.ToString()}");
            System.out.println("\tAlignment:\t{args.CurrentShape.VerticalAlignment}");
            System.out.println("\tWrap type:\t{args.CurrentShape.WrapType}");
            System.out.println("Output filename:\t{args.ImageFileName}\n");
        }

        private int mImageCount;
    }
    //ExEnd
}
