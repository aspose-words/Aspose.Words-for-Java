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
import org.apache.commons.io.FilenameUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.text.MessageFormat;
import java.util.Arrays;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExHtmlFixedSaveOptions extends ApiExampleBase {
    @Test
    public void useEncoding() throws Exception {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.Encoding
        //ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello World!");

        // The default encoding is UTF-8. If we want to represent our document using a different encoding,
        // we can use a SaveOptions object to set a specific encoding.
        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        htmlFixedSaveOptions.setEncoding(StandardCharsets.US_ASCII);

        Assert.assertEquals("US-ASCII", htmlFixedSaveOptions.getEncoding().name());

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
        //ExEnd

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.UseEncoding.html"), "US-ASCII");

        Assert.assertTrue(outDocContents.contains("content=\"text/html; charset=US-ASCII\""));
    }

    @Test
    public void getEncoding() throws Exception {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();

        HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions();
        htmlFixedSaveOptions.setEncoding(StandardCharsets.UTF_16);

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.GetEncoding.html", htmlFixedSaveOptions);
    }

    @Test(dataProvider = "exportEmbeddedCssDataProvider")
    public void exportEmbeddedCss(boolean exportEmbeddedCss) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss.html"), "utf-8");

        if (exportEmbeddedCss) {
            Assert.assertTrue(outDocContents.contains("<style type=\"text/css\">"));
            Assert.assertFalse(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css").exists());
        } else {
            Assert.assertTrue(outDocContents.contains("<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css\" media=\"all\" />"));
            Assert.assertTrue(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css").exists());
        }
        //ExEnd
    }

    @DataProvider(name = "exportEmbeddedCssDataProvider")
    public static Object[][] exportEmbeddedCssDataProvider() throws Exception {
        return new Object[][]
                {
                        {true},
                        {false}
                };
    }

    @Test(dataProvider = "exportEmbeddedFontsDataProvider")
    public void exportEmbeddedFonts(boolean exportEmbeddedFonts) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css"), StandardCharsets.UTF_8);

        if (exportEmbeddedFonts)
        {
            Assert.assertTrue(Pattern.compile("@font-face [{] font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; [}]").matcher(outDocContents).find());
            Assert.assertEquals(0, DocumentHelper.directoryGetFiles(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts", "*.woff").size());
        }
        else
        {
            Assert.assertTrue(Pattern.compile("@font-face [{] font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; [}]").matcher(outDocContents).find());
            Assert.assertEquals(2, DocumentHelper.directoryGetFiles(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedFonts", "*.woff").size());
        }
        //ExEnd
    }

    @DataProvider(name = "exportEmbeddedFontsDataProvider")
    public static Object[][] exportEmbeddedFontsDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(dataProvider = "exportEmbeddedImagesDataProvider")
    public void exportEmbeddedImages(boolean exportImages) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages.html"), StandardCharsets.UTF_8);

        if (exportImages)
        {
            Assert.assertFalse(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg").exists());
            Assert.assertTrue(Pattern.compile("<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" src=\".+\" />").matcher(outDocContents).find());
        }
        else
        {
            Assert.assertTrue(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg").exists());
            Assert.assertTrue(Pattern.compile("<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" " +
                "src=\"HtmlFixedSaveOptions[.]ExportEmbeddedImages/image001[.]jpeg\" />").matcher(outDocContents).find());
        }
        //ExEnd
    }

    @DataProvider(name = "exportEmbeddedImagesDataProvider")
    public static Object[][] exportEmbeddedImagesDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(dataProvider = "exportEmbeddedSvgsDataProvider")
    public void exportEmbeddedSvgs(boolean exportSvgs) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html"), StandardCharsets.UTF_8);

        if (exportSvgs) {
            Assert.assertFalse(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg").exists());
            Assert.assertTrue(Pattern.compile("<image id=\"image004\" xlink:href=.+/>").matcher(outDocContents).find());
        } else {
            Assert.assertTrue(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg").exists());
            Assert.assertTrue(Pattern.compile("<object type=\"image/svg[+]xml\" data=\"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001[.]svg\"></object>").matcher(outDocContents).find());
        }
        //ExEnd
    }

    @DataProvider(name = "exportEmbeddedSvgsDataProvider")
    public static Object[][] exportEmbeddedSvgsDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(dataProvider = "exportFormFieldsDataProvider")
    public void exportFormFields(boolean exportFormFields) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.ExportFormFields.html"), StandardCharsets.UTF_8);

        if (exportFormFields)
        {
            Assert.assertTrue(Pattern.compile(
                "<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
                "<input style=\"position:absolute; left:0pt; top:0pt;\" type=\"checkbox\" name=\"CheckBox\" />").matcher(outDocContents).find());
        }
        else
        {
            Assert.assertTrue(Pattern.compile(
                "<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
                "<div class=\"awdiv\" style=\"left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;\"").matcher(outDocContents).find());
        }
        //ExEnd
    }

    @DataProvider(name = "exportFormFieldsDataProvider")
    public static Object[][] exportFormFieldsDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void addCssClassNamesPrefix() throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html"), StandardCharsets.UTF_8);

        Assert.assertTrue(Pattern.compile(
            "<div class=\"myprefixdiv myprefixpage\" style=\"width:595[.]3pt; height:841[.]9pt;\">" +
            "<div class=\"myprefixdiv\" style=\"left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];\">" +
            "<span class=\"myprefixspan myprefixtext001\" style=\"font-size:11pt; left:294[.]73pt; top:0[.]36pt;\">").matcher(outDocContents).find());

        outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css"), StandardCharsets.UTF_8);

        Assert.assertTrue(Pattern.compile(
            ".myprefixdiv [{] position:absolute; [}] " +
            ".myprefixspan [{] position:absolute; white-space:pre; color:#000000; font-size:12pt; [}]").matcher(outDocContents).find());
        //ExEnd
    }

    @Test(dataProvider = "horizontalAlignmentDataProvider")
    public void horizontalAlignment(int pageHorizontalAlignment) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.HorizontalAlignment/styles.css"), StandardCharsets.UTF_8);

        switch (pageHorizontalAlignment)
        {
            case HtmlFixedPageHorizontalAlignment.CENTER:
                Assert.assertTrue(Pattern.compile(
                    "[.]awpage [{] position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; [}]").matcher(outDocContents).find());
                break;
            case HtmlFixedPageHorizontalAlignment.LEFT:
                Assert.assertTrue(Pattern.compile(
                    "[.]awpage [{] position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; [}]").matcher(outDocContents).find());
                break;
            case HtmlFixedPageHorizontalAlignment.RIGHT:
                Assert.assertTrue(Pattern.compile(
                    "[.]awpage [{] position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; [}]").matcher(outDocContents).find());
                break;
        }
        //ExEnd
    }

    @DataProvider(name = "horizontalAlignmentDataProvider")
    public static Object[][] horizontalAlignmentDataProvider() {
        return new Object[][]
                {
                        {HtmlFixedPageHorizontalAlignment.CENTER},
                        {HtmlFixedPageHorizontalAlignment.LEFT},
                        {HtmlFixedPageHorizontalAlignment.RIGHT},
                };
    }

    @Test
    public void pageMargins() throws Exception {
        //ExStart
        //ExFor:HtmlFixedSaveOptions.PageMargins
        //ExSummary:Shows how to adjust page margins when saving a document to HTML.
        Document doc = new Document(getMyDir() + "Document.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setPageMargins(15.0);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.PageMargins.html", saveOptions);

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.PageMargins/styles.css"), StandardCharsets.UTF_8);

        Assert.assertTrue(Pattern.compile(
            "[.]awpage [{] position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; [}]").matcher(outDocContents).find());
        //ExEnd
    }

    @Test
    public void pageMarginsException() {
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        Assert.assertThrows(IllegalArgumentException.class, () -> saveOptions.setPageMargins(-1));
    }

    @Test(dataProvider = "optimizeGraphicsOutputDataProvider")
    public void optimizeGraphicsOutput(boolean optimizeOutput) throws Exception {
        //ExStart
        //ExFor:FixedPageSaveOptions.OptimizeOutput
        //ExFor:HtmlFixedSaveOptions.OptimizeOutput
        //ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setOptimizeOutput(optimizeOutput);
        }

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html", saveOptions);

        // The size of the optimized version of the document is almost a third of the size of the unoptimized document.
        if (optimizeOutput)
            Assert.assertEquals(57220.0,
                new File(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html").length(), 200.0);
        else
            Assert.assertEquals(159100.0,
                new File(getArtifactsDir() + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html").length(), 200.0);
        //ExEnd
    }

    @DataProvider(name = "optimizeGraphicsOutputDataProvider")
    public static Object[][] optimizeGraphicsOutputDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }


    @Test(dataProvider = "usingMachineFontsDataProvider")
    public void usingMachineFonts(boolean useTargetMachineFonts) throws Exception {
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

        String outDocContents = FileUtils.readFileToString(new File(getArtifactsDir() + "HtmlFixedSaveOptions.UsingMachineFonts.html"), StandardCharsets.UTF_8);

        if (useTargetMachineFonts)
            Assert.assertFalse(Pattern.compile("@font-face").matcher(outDocContents).find());
        else
            Assert.assertTrue(Pattern.compile(
                "@font-face [{] font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], " +
                "url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; [}]").matcher(outDocContents).find());
        //ExEnd
    }

    @DataProvider(name = "usingMachineFontsDataProvider")
    public static Object[][] usingMachineFontsDataProvider() {
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
    public void resourceSavingCallback() throws Exception {
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

    private static class FontSavingCallback implements IResourceSavingCallback {
        /// <summary>
        /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
        /// </summary>
        public void resourceSaving(ResourceSavingArgs args) {
            mText.append(MessageFormat.format("Original document URI:\t{0}", args.getDocument().getOriginalFileName()));
            mText.append(MessageFormat.format("Resource being saved:\t{0}", args.getResourceFileName()));
            mText.append(MessageFormat.format("Full uri after saving:\t{0}\n", args.getResourceFileUri()));
        }

        public String getText() {
            return mText.toString();
        }

        private final StringBuilder mText = new StringBuilder();
    }
    //ExEnd

    private void testResourceSavingCallback(FontSavingCallback callback) {
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
    public void htmlFixedResourceFolder() throws Exception {
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
        new File(options.getResourcesFolderAlias()).mkdir();

        doc.save(getArtifactsDir() + "HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

        System.out.println(callback.getText());

        String[] resourceFiles = new File(getArtifactsDir() + "HtmlFixedResourceFolderAlias").list();

        Assert.assertFalse(new File(getArtifactsDir() + "HtmlFixedResourceFolder").exists());
        Assert.assertEquals(6, IterableUtils.countMatches(Arrays.asList(resourceFiles),
                f -> f.endsWith(".jpeg") || f.endsWith(".png") || f.endsWith(".css")));
        testHtmlFixedResourceFolder(callback); //ExSkip
    }

    /// <summary>
    /// Counts and prints URIs of resources contained by as they are converted to fixed HTML.
    /// </summary>
    private static class ResourceUriPrinter implements IResourceSavingCallback {
        public void resourceSaving(ResourceSavingArgs args) throws Exception {
            // If we set a folder alias in the SaveOptions object, we will be able to print it from here.
            mText.append(MessageFormat.format("Resource #{0} \"{1}\"", ++mSavedResourceCount, args.getResourceFileName()));

            String extension = FilenameUtils.getExtension(args.getResourceFileName());
            switch (extension) {
                case "ttf":
                case "woff": {
                    // By default, 'ResourceFileUri' uses system folder for fonts.
                    // To avoid problems in other platforms you must explicitly specify the path for the fonts.
                    args.setResourceFileUri(getArtifactsDir() + File.separatorChar + args.getResourceFileName());
                    break;
                }
            }

            mText.append("\t" + args.getResourceFileUri());

            // If we have specified a folder in the "ResourcesFolderAlias" property,
            // we will also need to redirect each stream to put its resource in that folder.
            args.setResourceStream(new FileOutputStream(args.getResourceFileUri()));
            args.setKeepResourceStreamOpen(false);
        }

        public String getText() {
            return mText.toString();
        }

        private int mSavedResourceCount;
        private final /*final*/ StringBuilder mText = new StringBuilder();
    }
    //ExEnd

    private void testHtmlFixedResourceFolder(ResourceUriPrinter callback)
    {
        int count = 0;

        Matcher matcher = Pattern.compile("Resource #").matcher(callback.getText());
        while (matcher.find())
            count++;

        Assert.assertEquals(16, count);
    }
}