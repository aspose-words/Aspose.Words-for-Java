package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.util.ArrayList;
import java.util.regex.Pattern;

public class ExHtmlSaveOptions extends ApiExampleBase {
    // For assert this test you need to open HTML docs and they shouldn't have negative left margins
    @Test(dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(final int saveFormat) throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setSaveFormat(saveFormat);
        saveOptions.setExportPageMargins(true);

        save(doc, "HtmlSaveOptions.ExportPageMargins." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportPageMarginsDataProvider")
    public static Object[][] exportPageMarginsDataProvider() {
        return new Object[][]{
                {SaveFormat.HTML},
                {SaveFormat.MHTML},
                {SaveFormat.EPUB}
        };
    }

    @Test(dataProvider = "exportOfficeMathDataProvider")
    public void exportOfficeMath(final int saveFormat, final int outputMode) throws Exception {
        Document doc = new Document(getMyDir() + "OfficeMath.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setOfficeMathOutputMode(outputMode);

        save(doc, "HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);

        switch (saveFormat) {
            case SaveFormat.HTML:
                DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(),
                        "<img src=\"HtmlSaveOptions.ExportToHtmlUsingImage.001.png\" width=\"49\" height=\"19\" alt=\"\" style=\"vertical-align:middle; -aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />");
                return;

            case SaveFormat.MHTML:
                DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(),
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>A</mi><mo>=</mo><mi>π</mi><msup><mrow><mi>r</mi></mrow><mrow><mn>2</mn></mrow></msup></math>");
                return;

            case SaveFormat.EPUB:
                DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportToHtmlUsingImage." + SaveFormat.toString(saveFormat).toLowerCase(),
                        "<span style=\"font-family:\'Cambria Math\'\">A=π</span><span style=\"font-family:\'Cambria Math\'\">r</span><span style=\"font-family:\'Cambria Math\'\">2</span>");
                return;
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportOfficeMathDataProvider")
    public static Object[][] exportOfficeMathDataProvider() {
        return new Object[][]{
                {SaveFormat.HTML, HtmlOfficeMathOutputMode.IMAGE},
                {SaveFormat.MHTML, HtmlOfficeMathOutputMode.MATH_ML},
                {SaveFormat.EPUB, HtmlOfficeMathOutputMode.TEXT}};
    }

    @Test(dataProvider = "exportTextBoxAsSvgDataProvider")
    public void exportTextBoxAsSvg(final int saveFormat, final boolean textBoxAsSvg) throws Exception {
        ArrayList<String> dirFiles;

        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportTextBoxAsSvg.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportTextBoxAsSvg(textBoxAsSvg);

        save(doc, "HtmlSaveOptions.ExportTextBoxAsSvg." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);

        switch (saveFormat) {
            case SaveFormat.HTML:

                dirFiles = directoryGetFiles(getArtifactsDir() + "", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png");
                Assert.assertTrue(dirFiles.isEmpty());
                return;

            case SaveFormat.EPUB:

                dirFiles = directoryGetFiles(getArtifactsDir() + "", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png");
                Assert.assertTrue(dirFiles.isEmpty());
                return;

            case SaveFormat.MHTML:

                dirFiles = directoryGetFiles(getArtifactsDir() + "", "HtmlSaveOptions.ExportTextBoxAsSvg.001.png");
                Assert.assertFalse(dirFiles.isEmpty());
                return;
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportTextBoxAsSvgDataProvider")
    public static Object[][] exportTextBoxAsSvgDataProvider() throws Exception {
        return new Object[][]{
                {SaveFormat.HTML, true},
                {SaveFormat.EPUB, true},
                {SaveFormat.MHTML, false}
        };
    }

    private ArrayList<String> directoryGetFiles(final String dirname, final String filenamePattern) {
        File dirFile = new File(dirname);
        Pattern re = Pattern.compile(filenamePattern.replace("*", ".*").replace("?", ".?"));
        ArrayList<String> dirFiles = new ArrayList<>();
        for (File file : dirFile.listFiles()) {
            if (file.isDirectory()) {
                dirFiles.addAll(directoryGetFiles(file.getPath(), filenamePattern));
            } else {
                if (re.matcher(file.getName()).matches()) {
                    dirFiles.add(file.getPath());
                }
            }
        }
        return dirFiles;
    }

    private static Document save(final Document inputDoc, final String outputDocPath, final int saveFormat, final SaveOptions saveOptions) throws Exception {
        switch (saveFormat) {
            case SaveFormat.HTML:
                inputDoc.save(getArtifactsDir() + outputDocPath, saveOptions);
                return inputDoc;
            case SaveFormat.MHTML:
                inputDoc.save(getArtifactsDir() + outputDocPath, saveOptions);
                return inputDoc;
            case SaveFormat.EPUB:
                inputDoc.save(getArtifactsDir() + outputDocPath, saveOptions);
                return inputDoc;
        }

        return inputDoc;
    }

    @Test(dataProvider = "controlListLabelsExportToHtmlDataProvider")
    public void controlListLabelsExportToHtml(final int howExportListLabels) throws Exception {
        Document doc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        // 'ExportListLabels.Auto' - this option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss,
        // otherwise HTML <p> tag is used. This is also the default value.
        // 'ExportListLabels.AsInlineText' - using this option the <p> tag is used for any list label representation.
        // 'ExportListLabels.ByHtmlTags' - The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
        saveOptions.setExportListLabels(howExportListLabels);

        doc.save(getArtifactsDir() + "Document.ExportListLabels.html", saveOptions);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "controlListLabelsExportToHtmlDataProvider")
    public static Object[][] controlListLabelsExportToHtmlDataProvider() {
        return new Object[][]{
                {ExportListLabels.AUTO},
                {ExportListLabels.AS_INLINE_TEXT},
                {ExportListLabels.BY_HTML_TAGS}
        };
    }

    @Test
    public void controlListLabelsExportToHtml() throws Exception {
        Document doc = new Document(getMyDir() + "Lists.PrintOutAllLists.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        // This option uses <ul> and <ol> tags are used for list label representation if it doesn't cause formatting loss,
        // otherwise HTML <p> tag is used. This is also the default value.
        saveOptions.setExportListLabels(ExportListLabels.AUTO);
        doc.save(getArtifactsDir() + "Document.ExportListLabels Auto.html", saveOptions);

        // Using this option the <p> tag is used for any list label representation.
        saveOptions.setExportListLabels(ExportListLabels.AS_INLINE_TEXT);
        doc.save(getArtifactsDir() + "Document.ExportListLabels InlineText.html", saveOptions);

        // The <ul> and <ol> tags are used for list label representation. Some formatting loss is possible.
        saveOptions.setExportListLabels(ExportListLabels.BY_HTML_TAGS);
        doc.save(getArtifactsDir() + "Document.ExportListLabels HtmlTags.html", saveOptions);
    }

    @Test(dataProvider = "exportUrlForLinkedImageDataProvider")
    public void exportUrlForLinkedImage(final boolean export) throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportOriginalUrlForLinkedImages(export);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        ArrayList<String> dirFiles = directoryGetFiles(getArtifactsDir() + "", "HtmlSaveOptions.ExportUrlForLinkedImage.001.png");

        if (dirFiles.size() == 0) {
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
        } else {
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportUrlForLinkedImageDataProvider")
    public static Object[][] exportUrlForLinkedImageDataProvider() {
        return new Object[][]{
                {true},
                {false}
        };
    }

    @Test(enabled = false, description = "Bug, css styles starting with -aw, even if ExportRoundtripInformation is false", dataProvider = "exportRoundtripInformationDataProvider")
    public void exportRoundtripInformation(final boolean valueHtml) throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportRoundtripInformation(valueHtml);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html");

        if (valueHtml) {
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html",
                    "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
        } else {
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html",
                    "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportRoundtripInformationDataProvider")
    public static Object[][] exportRoundtripInformationDataProvider() {
        return new Object[][]{
                {true},
                {false}
        };
    }

    @Test
    public void roundtripInformationDefaulValue() {
        // Assert that default value is true for HTML and false for MHTML and EPUB.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), true);

        saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);

        saveOptions = new HtmlSaveOptions(SaveFormat.EPUB);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);
    }

    @Test
    public void configForSavingExternalResources() throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setResourceFolderAlias("https://www.aspose.com/");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);

        ArrayList<String> imageFiles = directoryGetFiles(getArtifactsDir() + "Resources\\", "*.png");
        Assert.assertEquals(imageFiles.size(), 3);

        ArrayList<String> fontFiles = directoryGetFiles(getArtifactsDir() + "Resources\\", "*.ttf");
        Assert.assertEquals(fontFiles.size(), 1);

        ArrayList<String> cssFiles = directoryGetFiles(getArtifactsDir() + "Resources\\", "*.css");
        Assert.assertEquals(cssFiles.size(), 1);

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins Out.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins Out.css\"");
    }

    @Test
    public void convertFontsAsBase64() throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);
    }

    @Test(dataProvider = "html5SupportDataProvider")
    public void html5Support(final int htmlVersion) throws Exception {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setHtmlVersion(htmlVersion);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "html5SupportDataProvider")
    public static Object[][] html5SupportDataProvider() throws Exception {
        return new Object[][]{
                {HtmlVersion.HTML_5},
                {HtmlVersion.XHTML}
        };
    }

    @Test(dataProvider = "exportFontsDataProvider")
    public void exportFonts(final boolean exportAsBase64) throws Exception {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(exportAsBase64);

        if (!exportAsBase64) {
            doc.save(getArtifactsDir() + "DocumentExportFonts Out 1.html", saveOptions);
            // Verify that the font has been added to the folder
            Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "", "DocumentExportFonts Out 1.times.ttf").isEmpty());

        } else {
            doc.save(getArtifactsDir() + "DocumentExportFonts Out 2.html", saveOptions);
            // Verify that the font is not added to the folder
            Assert.assertTrue(directoryGetFiles(getArtifactsDir() + "", "DocumentExportFonts Out 2.times.ttf").isEmpty());

        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportFontsDataProvider")
    public static Object[][] exportFontsDataProvider() throws Exception {
        return new Object[][]{
                {false},
                {true}
        };
    }

    @Test
    public void resourceFolderPriority() throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.001.jpeg").isEmpty());
        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.002.png").isEmpty());
        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf").isEmpty());
        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.css").isEmpty());

    }

    @Test
    public void resourceFolderLowPriority() throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setFontsFolder(getArtifactsDir() + "Fonts");
        saveOptions.setImagesFolder(getArtifactsDir() + "Images");
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolder Out.001.jpeg").isEmpty());
        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolder Out.002.png").isEmpty());
        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Fonts", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf").isEmpty());
        Assert.assertFalse(directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.css").isEmpty());
    }

    @Test
    public void svgMetafileFormat() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an SVG image: ");
        builder.insertHtml("<svg height='210' width='500'>\r\n<polygon points='100,10 40,198 190,78 10,78 160,198'\r\n"
                + "style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n</svg> ");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setMetafileFormat(HtmlMetafileFormat.SVG);

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html", saveOptions);
    }

    @Test
    public void pngMetafileFormat() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an Png image: ");
        builder.insertHtml("<svg height='210' width='500'>\r\n<polygon points='100,10 40,198 190,78 10,78 160,198'\r\n"
                + "style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n</svg>");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setMetafileFormat(HtmlMetafileFormat.PNG);

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html", saveOptions);
    }

    @Test
    public void emfOrWmfMetafileFormat() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an image as is: ");
        builder.insertHtml("<img src=\"data:image/png;base64,\r\niVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP\r\n"
                + "C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA\r\nAAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J\r\n"
                + "REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq\r\nch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0\r\n"
                + "vr4MkhoXe0rZigAAAABJRU5ErkJggg==\" alt=\"Red dot\" />");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setMetafileFormat(HtmlMetafileFormat.EMF_OR_WMF);

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html", saveOptions);
    }

    @Test
    public void cssClassNamesPrefix() throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions.CssClassNamePrefix
        //ExSummary: Shows how to specifies a prefix which is added to all CSS class names
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.CssClassNamePrefix.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
        saveOptions.setCssClassNamePrefix("aspose-");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
        //ExEnd
    }

    @Test(expectedExceptions = IllegalArgumentException.class)
    public void cssClassNamesNotValidPrefix() {
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssClassNamePrefix("@%-");
    }

    @Test
    public void cssClassNamesNullPrefix() throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.CssClassNamePrefix.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
        saveOptions.setCssClassNamePrefix(null);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
    }

    @Test
    public void contentIdScheme() throws Exception {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ContentIdScheme.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        saveOptions.setPrettyFormat(true);
        saveOptions.setExportCidUrlsForMhtmlResources(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
    }

    @Test(enabled = false, description = "Bug")
    public void resolveFontNames() throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions.ResolveFontNames
        //ExSummary:Shows how to resolve all font names before writing them to HTML.
        Document document = new Document(getMyDir() + "HtmlSaveOptions.ResolveFontNames.docx");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontConfigSubstitution().setEnabled(false);

        document.setFontSettings(fontSettings);

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        // By default this option is set to 'False' and Aspose.Words writes font names as specified in the source document.
        saveOptions.setResolveFontNames(true);

        document.save(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);
        //ExEnd

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", "<span style=\"font-family:Arial\">");
    }
}
