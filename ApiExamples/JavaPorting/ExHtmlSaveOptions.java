// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.IO.SearchOption;
import org.testng.Assert;
import com.aspose.words.ExportListLabels;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.CssStyleSheetType;
import com.aspose.words.HtmlVersion;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.HtmlMetafileFormat;
import com.aspose.words.FontSettings;
import org.testng.annotations.DataProvider;


@Test
class ExHtmlSaveOptions !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test (dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(/*SaveFormat*/int saveFormat) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

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
        Document doc = new Document(getMyDir() + "OfficeMath.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setOfficeMathOutputMode(outputMode);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportToHtmlUsingImage" + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
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

        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportTextBoxAsSvg.docx");

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

    @Test (dataProvider = "controlListLabelsExportToHtmlDataProvider")
    public void controlListLabelsExportToHtml(/*ExportListLabels*/int howExportListLabels) throws Exception
    {
        Document doc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss, 
            // otherwise HTML <p> tag is used. This is also the default value.
            // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
            // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
            saveOptions.setExportListLabels(howExportListLabels);
        }

        doc.save(getArtifactsDir() + "Document.ExportListLabels.html", saveOptions);
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "controlListLabelsExportToHtmlDataProvider")
	public static Object[][] controlListLabelsExportToHtmlDataProvider() throws Exception
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
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setExportOriginalUrlForLinkedImages(export); }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        String[] dirFiles = Directory.getFiles(getArtifactsDir(), "HtmlSaveOptions.ExportUrlForLinkedImage.001.png", SearchOption.ALL_DIRECTORIES);

        if (dirFiles.length == 0)
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
        else
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
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
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setExportRoundtripInformation(true); }
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html", saveOptions);
    }

    @Test
    public void roundtripInformationDefaulValue()
    {
        //Assert that default value is true for HTML and false for MHTML and EPUB.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        msAssert.areEqual(true, saveOptions.getExportRoundtripInformation());

        saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        msAssert.areEqual(false, saveOptions.getExportRoundtripInformation());

        saveOptions = new HtmlSaveOptions(SaveFormat.EPUB);
        msAssert.areEqual(false, saveOptions.getExportRoundtripInformation());
    }

    @Test
    public void configForSavingExternalResources() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setExportFontResources(true);
            saveOptions.setResourceFolder("Resources");
            saveOptions.setResourceFolderAlias("https://www.aspose.com/");
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html", saveOptions);

        String[] imageFiles = Directory.getFiles(getArtifactsDir() + "Resources/", "*.png", SearchOption.ALL_DIRECTORIES);
        msAssert.areEqual(3, imageFiles.length);

        String[] fontFiles = Directory.getFiles(getArtifactsDir() + "Resources/", "*.ttf", SearchOption.ALL_DIRECTORIES);
        msAssert.areEqual(1, fontFiles.length);

        String[] cssFiles = Directory.getFiles(getArtifactsDir() + "Resources/", "*.css", SearchOption.ALL_DIRECTORIES);
        msAssert.areEqual(1, cssFiles.length);

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins.css\"");
    }

    @Test
    public void convertFontsAsBase64() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(true);
        
        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html", saveOptions);
	}

    @Test (dataProvider = "html5SupportDataProvider")
    public void html5Support(/*HtmlVersion*/int htmlVersion) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

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
			{HtmlVersion.HTML_5},
			{HtmlVersion.XHTML},
		};
	}

    @Test (dataProvider = "exportFontsDataProvider")
    public void exportFonts(boolean exportAsBase64) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setExportFontResources(true);
            saveOptions.setExportFontsAsBase64(exportAsBase64);
        }

        switch (exportAsBase64)
        {
            case false:

                doc.save(getArtifactsDir() + "DocumentExportFonts 1.html", saveOptions);
                Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir(), "DocumentExportFonts 1.times.ttf",
                    SearchOption.ALL_DIRECTORIES));
                break;

            case true:

                doc.save(getArtifactsDir() + "DocumentExportFonts 2.html", saveOptions);
                msAssert.isEmpty(Directory.getFiles(getArtifactsDir(), "DocumentExportFonts 2.times.ttf",
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
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolder.html", saveOptions);

        String[] a = Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder.001.jpeg",
            SearchOption.ALL_DIRECTORIES);
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder.001.jpeg", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder.002.png", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder.calibri.ttf", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder.css", SearchOption.ALL_DIRECTORIES));
    }

    @Test
    public void resourceFolderLowPriority() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setExportFontResources(true);
            saveOptions.setFontsFolder(getArtifactsDir() + "Fonts");
            saveOptions.setImagesFolder(getArtifactsDir() + "Images");
            saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
            saveOptions.setResourceFolderAlias("http://example.com/resources");
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolder.html", saveOptions);

        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Images",
            "HtmlSaveOptions.ResourceFolder.001.jpeg", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolder.002.png",
            SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Fonts",
            "HtmlSaveOptions.ResourceFolder.calibri.ttf", SearchOption.ALL_DIRECTORIES));
        Assert.IsNotEmpty(Directory.getFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder.css",
            SearchOption.ALL_DIRECTORIES));
    }

    @Test
    public void svgMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an SVG image: ");
        builder.insertHtml(
            "<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.PNG); });
    }

    @Test
    public void pngMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an Png image: ");
        builder.insertHtml(
            "<svg height='210' width='500'>\r\n                    <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                        style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n                  </svg> ");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.PNG); });
    }

    @Test
    public void emfOrWmfMetafileFormat() throws Exception
    {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an image as is: ");
        builder.insertHtml(
            "<img src=\"data:image/png;base64,\r\n                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\r\n                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\r\n                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\r\n                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\r\n                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\r\n                    vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html",
            new HtmlSaveOptions(); { .setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF); });
    }

    @Test
    public void cssClassNamesPrefix() throws Exception
    {
        //ExStart
        //ExFor:HtmlSaveOptions.CssClassNamePrefix
        //ExSummary: Shows how to specifies a prefix which is added to all CSS class names
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.CssClassNamePrefix.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
            saveOptions.setCssClassNamePrefix("aspose-");
        }

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
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.CssClassNamePrefix.docx");

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
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ContentIdScheme.docx");

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
        Document document = new Document(getMyDir() + "HtmlSaveOptions.ResolveFontNames.docx");

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
            // By default this option is set to 'False' and Aspose.Words writes font names as specified in the source document.
            saveOptions.setResolveFontNames(true); 
        }

        document.save(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);
        //ExEnd

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", "<span style=\"font-family:Arial\">");
    }
}
