package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.apache.commons.collections4.IterableUtils;
import org.apache.commons.io.FileUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.text.MessageFormat;
import java.util.ArrayList;

public class ExHtmlSaveOptions extends ApiExampleBase {
    @Test(dataProvider = "exportPageMarginsEpubDataProvider")
    public void exportPageMarginsEpub(int saveFormat) throws Exception {
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

    @DataProvider(name = "exportPageMarginsEpubDataProvider")
    public static Object[][] exportPageMarginsEpubDataProvider() {
        return new Object[][]
                {
                        {SaveFormat.HTML},
                        {SaveFormat.MHTML},
                        {SaveFormat.EPUB}
                };
    }

    @Test(dataProvider = "exportOfficeMathEpubDataProvider")
    public void exportOfficeMathEpub(int saveFormat, int outputMode) throws Exception {
        Document doc = new Document(getMyDir() + "Office math.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setOfficeMathOutputMode(outputMode);

        doc.save(
                getArtifactsDir() + "HtmlSaveOptions.ExportOfficeMathEpub" +
                        FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportOfficeMathEpubDataProvider")
    public static Object[][] exportOfficeMathEpubDataProvider() {
        return new Object[][]
                {
                        {SaveFormat.HTML, HtmlOfficeMathOutputMode.IMAGE},
                        {SaveFormat.MHTML, HtmlOfficeMathOutputMode.MATH_ML},
                        {SaveFormat.EPUB, HtmlOfficeMathOutputMode.TEXT}};
    }

    @Test(dataProvider = "exportTextBoxAsSvgEpubDataProvider")
    public void exportTextBoxAsSvgEpub(/*SaveFormat*/int saveFormat, boolean isTextBoxAsSvg) throws Exception {
        ArrayList<String> dirFiles;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Shape textbox = builder.insertShape(ShapeType.TEXT_BOX, 300.0, 100.0);
        builder.moveTo(textbox.getFirstParagraph());
        builder.write("Hello world!");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(saveFormat);
        saveOptions.setExportTextBoxAsSvg(isTextBoxAsSvg);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTextBoxAsSvgEpub" + FileFormatUtil.saveFormatToExtension(saveFormat), saveOptions);

        switch (saveFormat) {
            case SaveFormat.HTML:

                dirFiles = DocumentHelper.directoryGetFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png");
                Assert.assertTrue(dirFiles.isEmpty());
                return;

            case SaveFormat.EPUB:

                dirFiles = DocumentHelper.directoryGetFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png");
                Assert.assertTrue(dirFiles.isEmpty());
                return;

            case SaveFormat.MHTML:

                dirFiles = DocumentHelper.directoryGetFiles(getArtifactsDir(), "HtmlSaveOptions.ExportTextBoxAsSvgEpub.001.png");
                Assert.assertTrue(dirFiles.isEmpty());
                return;
        }
    }

    @DataProvider(name = "exportTextBoxAsSvgEpubDataProvider")
    public static Object[][] exportTextBoxAsSvgEpubDataProvider() {
        return new Object[][]
                {
                        {SaveFormat.HTML, true},
                        {SaveFormat.EPUB, true},
                        {SaveFormat.MHTML, false}
                };
    }

    @Test(dataProvider = "controlListLabelsExportDataProvider")
    public void controlListLabelsExport(final int howExportListLabels) throws Exception {
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

    @DataProvider(name = "controlListLabelsExportDataProvider")
    public static Object[][] controlListLabelsExportDataProvider() {
        return new Object[][]{
                {ExportListLabels.AUTO},
                {ExportListLabels.AS_INLINE_TEXT},
                {ExportListLabels.BY_HTML_TAGS}
        };
    }

    @Test(dataProvider = "exportUrlForLinkedImageDataProvider")
    public void exportUrlForLinkedImage(boolean export) throws Exception {
        Document doc = new Document(getMyDir() + "Linked image.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportOriginalUrlForLinkedImages(export);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        ArrayList<String> dirFiles = DocumentHelper.directoryGetFiles(getArtifactsDir() + "", "HtmlSaveOptions.ExportUrlForLinkedImage.001.png");

        if (dirFiles.size() == 0) {
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
        } else {
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
        }
    }

    @DataProvider(name = "exportUrlForLinkedImageDataProvider")
    public static Object[][] exportUrlForLinkedImageDataProvider() {
        return new Object[][]{
                {true},
                {false}
        };
    }

    @Test
    public void exportRoundtripInformation() throws Exception {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setExportRoundtripInformation(true);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html", saveOptions);
    }

    @Test
    public void roundtripInformationDefaulValue() {
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), true);

        saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);

        saveOptions = new HtmlSaveOptions(SaveFormat.EPUB);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);
    }

    @Test
    public void externalResourceSavingConfig() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setResourceFolderAlias("https://www.aspose.com/");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExternalResourceSavingConfig.html", saveOptions);

        ArrayList<String> imageFiles = DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.png");
        Assert.assertEquals(imageFiles.size(), 8);

        ArrayList<String> fontFiles = DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.ttf");
        Assert.assertEquals(fontFiles.size(), 10);

        ArrayList<String> cssFiles = DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources/", "HtmlSaveOptions.ExternalResourceSavingConfig*.css");
        Assert.assertEquals(cssFiles.size(), 1);

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExternalResourceSavingConfig.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExternalResourceSavingConfig.css\"");
    }

    @Test
    public void convertFontsAsBase64() throws Exception {
        Document doc = new Document(getMyDir() + "TextBoxes.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ConvertFontsAsBase64.html", saveOptions);
    }

    @Test(dataProvider = "html5SupportDataProvider")
    public void html5Support(final int htmlVersion) throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setHtmlVersion(htmlVersion);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.Html5Support.html", saveOptions);
    }

    @DataProvider(name = "html5SupportDataProvider")
    public static Object[][] html5SupportDataProvider() throws Exception {
        return new Object[][]{
                {HtmlVersion.HTML_5},
                {HtmlVersion.XHTML}
        };
    }

    @Test(dataProvider = "exportFontsDataProvider")
    public void exportFonts(boolean exportAsBase64) throws Exception {
        String fontsFolder = getArtifactsDir() + "HtmlSaveOptions.ExportFonts.Resources";

        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportFontResources(true);
        saveOptions.setFontsFolder(fontsFolder);
        saveOptions.setExportFontsAsBase64(exportAsBase64);

        if (exportAsBase64 == false) {
            doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFonts.False.html", saveOptions);
            Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir(), "HtmlSaveOptions.ExportFonts.False.times.ttf").isEmpty());

        } else if (exportAsBase64 == true) {
            doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportFonts.True.html", saveOptions);
            Assert.assertTrue(DocumentHelper.directoryGetFiles(getArtifactsDir(), "HtmlSaveOptions.ExportFonts.True.times.ttf").isEmpty());

        }
    }

    @DataProvider(name = "exportFontsDataProvider")
    public static Object[][] exportFontsDataProvider() {
        return new Object[][]{
                {false},
                {true}
        };
    }

    @Test
    public void resourceFolderPriority() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolderPriority.html", saveOptions);

        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.001.png").isEmpty());
        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.002.png").isEmpty());
        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.arial.ttf").isEmpty());
        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderPriority.css").isEmpty());
    }

    @Test
    public void resourceFolderLowPriority() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setFontsFolder(getArtifactsDir() + "Fonts");
        saveOptions.setImagesFolder(getArtifactsDir() + "Images");
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolderLowPriority.html", saveOptions);

        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolderLowPriority.001.png").isEmpty());
        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolderLowPriority.002.png").isEmpty());
        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Fonts", "HtmlSaveOptions.ResourceFolderLowPriority.arial.ttf").isEmpty());
        Assert.assertFalse(DocumentHelper.directoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolderLowPriority.css").isEmpty());
    }

    @Test
    public void svgMetafileFormat() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an SVG image: ");
        builder.insertHtml("<svg height='210' width='500'>\r\n<polygon points='100,10 40,198 190,78 10,78 160,198'\r\n"
                + "style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n</svg> ");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setMetafileFormat(HtmlMetafileFormat.SVG);

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.SvgMetafileFormat.html", saveOptions);
    }

    @Test
    public void pngMetafileFormat() throws Exception {
        DocumentBuilder builder = new DocumentBuilder();

        builder.write("Here is an Png image: ");
        builder.insertHtml("<svg height='210' width='500'>\r\n<polygon points='100,10 40,198 190,78 10,78 160,198'\r\n"
                + "style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n</svg>");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setMetafileFormat(HtmlMetafileFormat.PNG);

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.PngMetafileFormat.html", saveOptions);
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

        builder.getDocument().save(getArtifactsDir() + "HtmlSaveOptions.EmfOrWmfMetafileFormat.html", saveOptions);
    }

    @Test
    public void cssClassNamesPrefix() throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html"), StandardCharsets.UTF_8);

        Assert.assertTrue(outDocContents.contains("<p class=\"myprefix-Header\">"));
        Assert.assertTrue(outDocContents.contains("<p class=\"myprefix-Footer\">"));

        outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.css"), StandardCharsets.UTF_8);

        Assert.assertTrue(outDocContents.contains(".myprefix-Footer { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n" +
                ".myprefix-Header { margin-bottom:0pt; line-height:normal; font-family:Arial; font-size:11pt }\r\n"));
        //ExEnd
    }

    @Test
    public void cssClassNamesNotValidPrefix() {
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        Assert.assertThrows(IllegalArgumentException.class, () -> saveOptions.setCssClassNamePrefix("@%-"));
    }

    @Test
    public void cssClassNamesNullPrefix() throws Exception {
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EMBEDDED);
        saveOptions.setCssClassNamePrefix(null);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.CssClassNamePrefix.html", saveOptions);
    }

    @Test
    public void contentIdScheme() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        saveOptions.setPrettyFormat(true);
        saveOptions.setExportCidUrlsForMhtmlResources(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ContentIdScheme.mhtml", saveOptions);
    }

    @Test(enabled = false, description = "Bug", dataProvider = "resolveFontNamesDataProvider")
    public void resolveFontNames(boolean resolveFontNames) throws Exception {
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
            fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
            fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setEnabled(true);
        }

        doc.setFontSettings(fontSettings);

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            // By default, this option is set to 'False' and Aspose.Words writes font names as specified in the source document.
            saveOptions.setResolveFontNames(resolveFontNames);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html", saveOptions);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ResolveFontNames.html"), "utf-8");

        Assert.assertTrue(outDocContents.matches("<span style=\"font-family:Arial\">"));
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "resolveFontNamesDataProvider")
    public static Object[][] resolveFontNamesDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void headingLevels() throws Exception {
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

        Assert.assertEquals("Heading #5\rHeading #6", doc.getText().trim());
        //ExEnd
    }

    @Test(dataProvider = "negativeIndentDataProvider")
    public void negativeIndent(boolean allowNegativeIndent) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.NegativeIndent.html"), StandardCharsets.UTF_8);

        if (allowNegativeIndent) {
            Assert.assertTrue(outDocContents.contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:-41.65pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
            Assert.assertTrue(outDocContents.contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
        } else {
            Assert.assertTrue(outDocContents.contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
            Assert.assertTrue(outDocContents.contains(
                    "<table cellspacing=\"0\" cellpadding=\"0\" style=\"margin-left:30.35pt; border:0.75pt solid #000000; -aw-border:0.5pt single; -aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
        }
        //ExEnd
    }

    @DataProvider(name = "negativeIndentDataProvider")
    public static Object[][] negativeIndentDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void folderAlias() throws Exception {
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
    @Test(enabled = false) //ExSkip
    public void saveExportedFonts() throws Exception {
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

        File[] fontFileNames = new File(getArtifactsDir()).listFiles((d, name) -> name.endsWith(".ttf"));

        for (File fontFilename : fontFileNames) {
            System.out.println(fontFilename.getName());
        }

        Assert.assertEquals(10, fontFileNames.length); //ExSkip
    }

    /// <summary>
    /// Prints information about exported fonts and saves them in the same local system folder as their output .html.
    /// </summary>
    public static class HandleFontSaving implements IFontSavingCallback {
        public void fontSaving(FontSavingArgs args) throws Exception {
            System.out.println(MessageFormat.format("Font:\t{0}", args.getFontFamilyName()));
            if (args.getBold()) System.out.print(", bold");
            if (args.getItalic()) System.out.print(", italic");
            System.out.println(MessageFormat.format("\nSource:\t{0}, {1} bytes\n", args.getOriginalFileName(), args.getOriginalFileSize()));

            // We can also access the source document from here.
            Assert.assertTrue(args.getDocument().getOriginalFileName().endsWith("Rendering.docx"));

            Assert.assertTrue(args.isExportNeeded());
            Assert.assertTrue(args.isSubsettingNeeded());

            String[] splittedFileName = args.getOriginalFileName().split("\\\\");
            String fileName = splittedFileName[splittedFileName.length - 1];

            // There are two ways of saving an exported font.
            // 1 -  Save it to a local file system location:
            args.setFontFileName(fileName);

            // 2 -  Save it to a stream:
            args.setFontStream(new FileOutputStream(fileName));
            Assert.assertFalse(args.getKeepFontStreamOpen());
        }
    }
    //ExEnd

    @Test(dataProvider = "htmlVersionsDataProvider")
    public void htmlVersions(/*HtmlVersion*/int htmlVersion) throws Exception {
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
        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.HtmlVersions.html"), StandardCharsets.UTF_8);

        switch (htmlVersion) {
            case HtmlVersion.HTML_5:
                Assert.assertTrue(outDocContents.contains("<a id=\"_Toc76372689\"></a>"));
                Assert.assertTrue(outDocContents.contains("<a id=\"_Toc76372689\"></a>"));
                Assert.assertTrue(outDocContents.contains("<table style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
                break;
            case HtmlVersion.XHTML:
                Assert.assertTrue(outDocContents.contains("<a name=\"_Toc76372689\"></a>"));
                Assert.assertTrue(outDocContents.contains("<ul type=\"disc\" style=\"margin:0pt; padding-left:0pt\">"));
                Assert.assertTrue(outDocContents.contains("<table cellspacing=\"0\" cellpadding=\"0\" style=\"-aw-border-insideh:0.5pt single #000000; -aw-border-insidev:0.5pt single #000000; border-collapse:collapse\">"));
                break;
        }
        //ExEnd
    }

    @DataProvider(name = "htmlVersionsDataProvider")
    public static Object[][] htmlVersionsDataProvider() {
        return new Object[][]
                {
                        {HtmlVersion.HTML_5},
                        {HtmlVersion.XHTML},
                };
    }

    @Test(dataProvider = "exportXhtmlTransitionalDataProvider")
    public void exportXhtmlTransitional(boolean showDoctypeDeclaration) throws Exception {
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
        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportXhtmlTransitional.html"), StandardCharsets.UTF_8);

        if (showDoctypeDeclaration)
            Assert.assertTrue(outDocContents.contains(
                    "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"no\"?>\r\n" +
                            "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">\r\n" +
                            "<html xmlns=\"http://www.w3.org/1999/xhtml\">"));
        else
            Assert.assertTrue(outDocContents.contains("<html>"));
        //ExEnd
    }

    @DataProvider(name = "exportXhtmlTransitionalDataProvider")
    public static Object[][] exportXhtmlTransitionalDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void epubHeadings() throws Exception {
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
    }

    @Test
    public void doc2EpubSaveOptions() throws Exception {
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
        saveOptions.setEncoding(StandardCharsets.UTF_8);

        // By default, an output .epub document will have all of its contents in one HTML part.
        // A split criterion allows us to segment the document into several HTML parts.
        // We will set the criteria to split the document into heading paragraphs.
        // This is useful for readers who cannot read HTML files more significant than a specific size.
        saveOptions.setDocumentSplitCriteria(DocumentSplitCriteria.HEADING_PARAGRAPH);

        // Specify that we want to export document properties.
        saveOptions.setExportDocumentProperties(true);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.Doc2EpubSaveOptions.epub", saveOptions);
        //ExEnd
    }

    @Test(dataProvider = "contentIdUrlsDataProvider")
    public void contentIdUrls(boolean exportCidUrlsForMhtmlResources) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ContentIdUrls.mht"), StandardCharsets.UTF_8);

        if (exportCidUrlsForMhtmlResources) {
            Assert.assertTrue(outDocContents.contains("Content-ID: <document.html>"));
            Assert.assertTrue(outDocContents.contains("<link href=3D\"cid:styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
            Assert.assertTrue(outDocContents.contains("@font-face { font-family:'Arial Black'; src:url('cid:ariblk.ttf') }"));
            Assert.assertTrue(outDocContents.contains("<img src=3D\"cid:image.003.jpeg\" width=3D\"351\" height=3D\"180\" alt=3D\"\" />"));
        } else {
            Assert.assertTrue(outDocContents.contains("Content-Location: document.html"));
            Assert.assertTrue(outDocContents.contains("<link href=3D\"styles.css\" type=3D\"text/css\" rel=3D\"stylesheet\" />"));
            Assert.assertTrue(outDocContents.contains("@font-face { font-family:'Arial Black'; src:url('ariblk.ttf') }"));
            Assert.assertTrue(outDocContents.contains("<img src=3D\"image.003.jpeg\" width=3D\"351\" height=3D\"180\" alt=3D\"\" />"));
        }
        //ExEnd
    }

    @DataProvider(name = "contentIdUrlsDataProvider")
    public static Object[][] contentIdUrlsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "dropDownFormFieldDataProvider")
    public void dropDownFormField(boolean exportDropDownFormFieldAsText) throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportDropDownFormFieldAsText
        //ExSummary:Shows how to get drop-down combo box form fields to blend in with paragraph text when saving to html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a document builder to insert a combo box with the value "Two" selected.
        builder.insertComboBox("MyComboBox", new String[]{"One", "Two", "Three"}, 1);

        // The "ExportDropDownFormFieldAsText" flag of this SaveOptions object allows us to
        // control how saving the document to HTML treats drop-down combo boxes.
        // Setting it to "true" will convert each combo box into simple text
        // that displays the combo box's currently selected value, effectively freezing it.
        // Setting it to "false" will preserve the functionality of the combo box using <select> and <option> tags.
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportDropDownFormFieldAsText(exportDropDownFormFieldAsText);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.DropDownFormField.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.DropDownFormField.html"), StandardCharsets.UTF_8);

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
    public static Object[][] dropDownFormFieldDataProvider() throws Exception {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "exportImagesAsBase64DataProvider")
    public void exportImagesAsBase64(boolean exportItemsAsBase64) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportImagesAsBase64.html"), StandardCharsets.UTF_8);

        Assert.assertTrue(exportItemsAsBase64
                ? outDocContents.contains("<img src=\"data:image/png;base64")
                : outDocContents.contains("<img src=\"HtmlSaveOptions.ExportImagesAsBase64.001.png\""));
        //ExEnd
    }

    @DataProvider(name = "exportImagesAsBase64DataProvider")
    public static Object[][] exportImagesAsBase64DataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }


    @Test
    public void exportFontsAsBase64() throws Exception {
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

    @Test(dataProvider = "exportLanguageInformationDataProvider")
    public void exportLanguageInformation(boolean exportLanguageInformation) throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportLanguageInformation
        //ExSummary:Shows how to preserve language information when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use the builder to write text while formatting it in different locales.
        builder.getFont().setLocaleId(1033);
        builder.writeln("Hello world!");

        builder.getFont().setLocaleId(2057);
        builder.writeln("Hello again!");

        builder.getFont().setLocaleId(1049);
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportLanguageInformation.html"), StandardCharsets.UTF_8);

        if (exportLanguageInformation) {
            Assert.assertTrue(outDocContents.contains("<span>Hello world!</span>"));
            Assert.assertTrue(outDocContents.contains("<span lang=\"en-GB\">Hello again!</span>"));
            Assert.assertTrue(outDocContents.contains("<span lang=\"ru-RU\">Привет, мир!</span>"));
        } else {
            Assert.assertTrue(outDocContents.contains("<span>Hello world!</span>"));
            Assert.assertTrue(outDocContents.contains("<span>Hello again!</span>"));
            Assert.assertTrue(outDocContents.contains("<span>Привет, мир!</span>"));
        }
        //ExEnd
    }

    @DataProvider(name = "exportLanguageInformationDataProvider")
    public static Object[][] exportLanguageInformationDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "listDataProvider")
    public void list(int exportListLabels) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportListLabels(exportListLabels);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.List.html", options);
        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.List.html"), StandardCharsets.UTF_8);

        switch (exportListLabels) {
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

    @DataProvider(name = "listDataProvider")
    public static Object[][] listDataProvider() {
        return new Object[][]
                {
                        {ExportListLabels.AS_INLINE_TEXT},
                        {ExportListLabels.AUTO},
                        {ExportListLabels.BY_HTML_TAGS},
                };
    }

    @Test(dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(boolean exportPageMargins) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportPageMargins(exportPageMargins);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins.html"), StandardCharsets.UTF_8);

        if (exportPageMargins) {
            Assert.assertTrue(outDocContents.contains("<style type=\"text/css\">div.Section1 { margin:72pt }</style>"));
            Assert.assertTrue(outDocContents.contains("<div class=\"Section1\"><p style=\"margin-top:0pt; margin-left:151pt; margin-bottom:0pt\">"));
        } else {
            Assert.assertFalse(outDocContents.contains("style type=\"text/css\">"));
            Assert.assertTrue(outDocContents.contains("<div><p style=\"margin-top:0pt; margin-left:223pt; margin-bottom:0pt\">"));
        }
        //ExEnd
    }

    @DataProvider(name = "exportPageMarginsDataProvider")
    public static Object[][] exportPageMarginsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "exportPageSetupDataProvider")
    public void exportPageSetup(boolean exportPageSetup) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportPageSetup(exportPageSetup);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageSetup.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportPageSetup.html"), StandardCharsets.UTF_8);

        if (exportPageSetup) {
            Assert.assertTrue(outDocContents.contains(
                    "<style type=\"text/css\">" +
                            "@page Section1 { size:419.55pt 595.3pt; margin:36pt 72pt }" +
                            "@page Section2 { size:612pt 792pt; margin:72pt }" +
                            "div.Section1 { page:Section1 }div.Section2 { page:Section2 }" +
                            "</style>"));

            Assert.assertTrue(outDocContents.contains(
                    "<div class=\"Section1\">" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Section 1</span>" +
                            "</p>" +
                            "</div>"));
        } else {
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

    @DataProvider(name = "exportPageSetupDataProvider")
    public static Object[][] exportPageSetupDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "relativeFontSizeDataProvider")
    public void relativeFontSize(boolean exportRelativeFontSize) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportRelativeFontSize(exportRelativeFontSize);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RelativeFontSize.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.RelativeFontSize.html"), StandardCharsets.UTF_8);

        if (exportRelativeFontSize) {
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
        } else {
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

    @DataProvider(name = "relativeFontSizeDataProvider")
    public static Object[][] relativeFontSizeDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "exportTextBoxDataProvider")
    public void exportTextBox(boolean exportTextBoxAsSvg) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportTextBoxAsSvg(exportTextBoxAsSvg);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTextBox.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportTextBox.html"), StandardCharsets.UTF_8);

        if (exportTextBoxAsSvg) {
            Assert.assertTrue(outDocContents.contains(
                    "<span style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\">" +
                            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" version=\"1.1\" width=\"133\" height=\"80\">"));
        } else {
            Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<img src=\"HtmlSaveOptions.ExportTextBox.001.png\" width=\"136\" height=\"83\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />" +
                            "</p>"));
        }
        //ExEnd
    }

    @DataProvider(name = "exportTextBoxDataProvider")
    public static Object[][] exportTextBoxDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "roundTripInformationDataProvider")
    public void roundTripInformation(boolean exportRoundtripInformation) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportRoundtripInformation(exportRoundtripInformation);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html"), StandardCharsets.UTF_8);
        doc = new Document(getArtifactsDir() + "HtmlSaveOptions.RoundTripInformation.html");

        if (exportRoundtripInformation) {
            Assert.assertTrue(outDocContents.contains("<div style=\"-aw-headerfooter-type:header-primary; clear:both\">"));
            Assert.assertTrue(outDocContents.contains("<span style=\"-aw-import:ignore\">&#xa0;</span>"));

            Assert.assertTrue(outDocContents.contains(
                    "td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                            "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top; " +
                            "-aw-border-bottom:0.5pt single; -aw-border-left:0.5pt single; -aw-border-top:0.5pt single\">"));

            Assert.assertTrue(outDocContents.contains(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt; -aw-font-family:'Courier New'; -aw-font-weight:normal; -aw-number-format:'o'\">"));

            Assert.assertTrue(outDocContents.contains(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"351\" height=\"180\" alt=\"\" " +
                            "style=\"-aw-left-pos:0pt; -aw-rel-hpos:column; -aw-rel-vpos:paragraph; -aw-top-pos:0pt; -aw-wrap-type:inline\" />"));


            Assert.assertTrue(outDocContents.contains(
                    "<span>Page number </span>" +
                            "<span style=\"-aw-field-start:true\"></span>" +
                            "<span style=\"-aw-field-code:' PAGE   \\\\* MERGEFORMAT '\"></span>" +
                            "<span style=\"-aw-field-separator:true\"></span>" +
                            "<span>1</span>" +
                            "<span style=\"-aw-field-end:true\"></span>"));

            Assert.assertEquals(1, IterableUtils.countMatches(doc.getRange().getFields(), f -> f.getType() == FieldType.FIELD_PAGE));
        } else {
            Assert.assertTrue(outDocContents.contains("<div style=\"clear:both\">"));
            Assert.assertTrue(outDocContents.contains("<span>&#xa0;</span>"));

            Assert.assertTrue(outDocContents.contains(
                    "<td colspan=\"2\" style=\"width:210.6pt; border-style:solid; border-width:0.75pt 6pt 0.75pt 0.75pt; " +
                            "padding-right:2.4pt; padding-left:5.03pt; vertical-align:top\">"));

            Assert.assertTrue(outDocContents.contains(
                    "<li style=\"margin-left:30.2pt; padding-left:5.8pt\">"));

            Assert.assertTrue(outDocContents.contains(
                    "<img src=\"HtmlSaveOptions.RoundTripInformation.003.jpeg\" width=\"351\" height=\"180\" alt=\"\" />"));

            Assert.assertTrue(outDocContents.contains(
                    "<span>Page number 1</span>"));

            Assert.assertEquals(0, IterableUtils.countMatches(doc.getRange().getFields(), f -> f.getType() == FieldType.FIELD_PAGE));
        }
        //ExEnd
    }

    @DataProvider(name = "roundTripInformationDataProvider")
    public static Object[][] roundTripInformationDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "exportTocPageNumbersDataProvider")
    public void exportTocPageNumbers(boolean exportTocPageNumbers) throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions.ExportTocPageNumbers
        //ExSummary:Shows how to display page numbers when saving a document with a table of contents to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table of contents, and then populate the document with paragraphs formatted using a "Heading"
        // style that the table of contents will pick up as entries. Each entry will display the heading paragraph on the left,
        // and the page number that contains the heading on the right.
        FieldToc fieldToc = (FieldToc) builder.insertField(FieldType.FIELD_TOC, true);

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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setExportTocPageNumbers(exportTocPageNumbers);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportTocPageNumbers.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.ExportTocPageNumbers.html"), StandardCharsets.UTF_8);

        if (exportTocPageNumbers) {
            Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Entry 1</span>" +
                            "<span style=\"width:425.84pt; font-family:'Lucida Console'; font-size:10pt; display:inline-block; -aw-font-family:'Times New Roman'; " +
                            "-aw-tabstop-align:right; -aw-tabstop-leader:dots; -aw-tabstop-pos:467.5pt\">......................................................................</span>" +
                            "<span>2</span>" +
                            "</p>"));
        } else {
            Assert.assertTrue(outDocContents.contains(
                    "<p style=\"margin-top:0pt; margin-bottom:0pt\">" +
                            "<span>Entry 1</span>" +
                            "</p>"));
        }
        //ExEnd
    }

    @DataProvider(name = "exportTocPageNumbersDataProvider")
    public static Object[][] exportTocPageNumbersDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "fontSubsettingDataProvider")
    public void fontSubsetting(int fontResourcesSubsettingSizeThreshold) throws Exception {
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

        File[] fontFileNames = new File(fontsFolder).listFiles((d, name) -> name.endsWith(".ttf"));

        Assert.assertEquals(3, fontFileNames.length);
        //ExEnd
    }

    @DataProvider(name = "fontSubsettingDataProvider")
    public static Object[][] fontSubsettingDataProvider() {
        return new Object[][]
                {
                        {0},
                        {1000000},
                        {Integer.MAX_VALUE},
                };
    }

    @Test(dataProvider = "metafileFormatDataProvider")
    public void metafileFormat(int htmlMetafileFormat) throws Exception {
        //ExStart
        //ExFor:HtmlMetafileFormat
        //ExFor:HtmlSaveOptions.MetafileFormat
        //ExSummary:Shows how to convert SVG objects to a different format when saving HTML documents.
        String html =
                "<html>\r\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\r\n                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\r\n                    </svg>\r\n                </html>";

        Document doc = new Document(new ByteArrayInputStream(html.getBytes()));

        // This document contains a <svg> element in the form of text.
        // When we save the document to HTML, we can pass a SaveOptions object
        // to determine how the saving operation handles this object.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Png" to convert it to a PNG image.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.Svg" preserve it as a SVG object.
        // Setting the "MetafileFormat" property to "HtmlMetafileFormat.EmfOrWmf" to convert it to a metafile.
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setMetafileFormat(htmlMetafileFormat);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html", options);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.MetafileFormat.html"), StandardCharsets.UTF_8);

        switch (htmlMetafileFormat) {
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

    @DataProvider(name = "metafileFormatDataProvider")
    public static Object[][] metafileFormatDataProvider() {
        return new Object[][]
                {
                        {HtmlMetafileFormat.PNG},
                        {HtmlMetafileFormat.SVG},
                        {HtmlMetafileFormat.EMF_OR_WMF},
                };
    }

    @Test(dataProvider = "officeMathOutputModeDataProvider")
    public void officeMathOutputMode(/*HtmlOfficeMathOutputMode*/int htmlOfficeMathOutputMode) throws Exception {
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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setOfficeMathOutputMode(htmlOfficeMathOutputMode);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.OfficeMathOutputMode.html", options);
        //ExEnd
    }

    @DataProvider(name = "officeMathOutputModeDataProvider")
    public static Object[][] officeMathOutputModeDataProvider() {
        return new Object[][]
                {
                        {HtmlOfficeMathOutputMode.IMAGE},
                        {HtmlOfficeMathOutputMode.MATH_ML},
                        {HtmlOfficeMathOutputMode.TEXT},
                };
    }

    @Test(dataProvider = "scaleImageToShapeSizeDataProvider")
    public void scaleImageToShapeSize(boolean scaleImageToShapeSize) throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions.ScaleImageToShapeSize
        //ExSummary:Shows how to disable the scaling of images to their parent shape dimensions when saving to .html.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a shape which contains an image, and then make that shape considerably smaller than the image.
        BufferedImage image = ImageIO.read(new File(getImageDir() + "Transparent background logo.png"));

        Assert.assertEquals(400, image.getWidth());
        Assert.assertEquals(400, image.getHeight());

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
        HtmlSaveOptions options = new HtmlSaveOptions();
        {
            options.setScaleImageToShapeSize(scaleImageToShapeSize);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ScaleImageToShapeSize.html", options);
        //ExEnd
    }

    @DataProvider(name = "scaleImageToShapeSizeDataProvider")
    public static Object[][] scaleImageToShapeSizeDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void imageFolder() throws Exception {
        //ExStart
        //ExFor:HtmlSaveOptions
        //ExFor:HtmlSaveOptions.ExportTextInputFormFieldAsText
        //ExFor:HtmlSaveOptions.ImagesFolder
        //ExSummary:Shows how to specify the folder for storing linked images after saving to .html.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        File imagesDir = new File(getArtifactsDir() + "SaveHtmlWithOptions");

        if (imagesDir.exists())
            imagesDir.delete();

        imagesDir.mkdir();

        // Set an option to export form fields as plain text instead of HTML input elements.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.HTML);
        options.setExportTextInputFormFieldAsText(true);
        options.setImagesFolder(imagesDir.getPath());

        doc.save(getArtifactsDir() + "HtmlSaveOptions.SaveHtmlWithOptions.html", options);
        //ExEnd

        Assert.assertTrue(new File(getArtifactsDir() + "HtmlSaveOptions.SaveHtmlWithOptions.html").exists());
        Assert.assertEquals(9, imagesDir.list().length);

        imagesDir.delete();
    }

    //ExStart
    //ExFor:ImageSavingArgs.CurrentShape
    //ExFor:ImageSavingArgs.Document
    //ExFor:ImageSavingArgs.ImageStream
    //ExFor:ImageSavingArgs.IsImageAvailable
    //ExFor:ImageSavingArgs.KeepImageStreamOpen
    //ExSummary:Shows how to involve an image saving callback in an HTML conversion process.
    @Test //ExSkip
    public void imageSavingCallback() throws Exception {
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
    private static class ImageShapePrinter implements IImageSavingCallback {
        public void imageSaving(ImageSavingArgs args) throws Exception {
            args.setKeepImageStreamOpen(false);
            Assert.assertTrue(args.isImageAvailable());

            String[] splitOriginalFileName = args.getDocument().getOriginalFileName().split("\\\\");
            System.out.println(MessageFormat.format("{0} Image #{1}", splitOriginalFileName[splitOriginalFileName.length - 1], ++mImageCount));

            LayoutCollector layoutCollector = new LayoutCollector(args.getDocument());

            System.out.println(MessageFormat.format("\tOn page:\t{0}", layoutCollector.getStartPageIndex(args.getCurrentShape())));
            System.out.println(MessageFormat.format("\tDimensions:\t{0}", args.getCurrentShape().getBounds().toString()));
            System.out.println(MessageFormat.format("\tAlignment:\t{0}", args.getCurrentShape().getVerticalAlignment()));
            System.out.println(MessageFormat.format("\tWrap type:\t{0}", args.getCurrentShape().getWrapType()));
            System.out.println(MessageFormat.format("Output filename:\t{0}\n", args.getImageFileName()));
        }

        private int mImageCount;
    }
    //ExEnd

    @Test(dataProvider = "prettyFormatDataProvider")
    public void prettyFormat(boolean usePrettyFormat) throws Exception {
        //ExStart
        //ExFor:SaveOptions.PrettyFormat
        //ExSummary:Shows how to enhance the readability of the raw code of a saved .html document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            htmlOptions.setPrettyFormat(usePrettyFormat);
        }

        doc.save(getArtifactsDir() + "HtmlSaveOptions.PrettyFormat.html", htmlOptions);

        // Enabling pretty format makes the raw html code more readable by adding tab stop and new line characters.
        String html = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlSaveOptions.PrettyFormat.html"), StandardCharsets.UTF_8);

        if (usePrettyFormat)
            Assert.assertEquals(
                    "<html>\r\n" +
                            "\t<head>\r\n" +
                            "\t\t<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\r\n" +
                            "\t\t<meta http-equiv=\"Content-Style-Type\" content=\"text/css\" />\r\n" +
                            MessageFormat.format("\t\t<meta name=\"generator\" content=\"{0} {1}\" />\r\n", BuildVersionInfo.getProduct(), BuildVersionInfo.getVersion()) +
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
                            MessageFormat.format("<meta name=\"generator\" content=\"{0} {1}\" /><title></title></head>", BuildVersionInfo.getProduct(), BuildVersionInfo.getVersion()) +
                            "<body style=\"font-family:'Times New Roman'; font-size:12pt\">" +
                            "<div><p style=\"margin-top:0pt; margin-bottom:0pt\"><span>Hello world!</span></p>" +
                            "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"-aw-import:ignore\">&#xa0;</span></p></div></body></html>",
                    html);
        //ExEnd
    }

    @DataProvider(name = "prettyFormatDataProvider")
    public static Object[][] prettyFormatDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }
}
