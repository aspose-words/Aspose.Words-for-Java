//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;
import org.testng.Assert;
import org.testng.annotations.DataProvider;

import java.io.File;
import java.util.ArrayList;
import java.util.regex.Pattern;

public class ExHtmlSaveOptions extends ApiExampleBase
{
    //For assert this test you need to open HTML docs and they shouldn't have negative left margins
    @Test(dataProvider = "exportPageMarginsDataProvider")
    public void exportPageMargins(/*SaveFormat*/int saveFormat) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setSaveFormat(saveFormat);
        saveOptions.setExportPageMargins(true);

        save(doc, getArtifactsDir()+ "HtmlSaveOptions.ExportPageMargins." + SaveFormat.toString(saveFormat).toLowerCase(), saveFormat, saveOptions);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportPageMarginsDataProvider")
    public static Object[][] exportPageMarginsDataProvider() {
        return new Object[][]{{SaveFormat.HTML}, {SaveFormat.MHTML}, {SaveFormat.EPUB},};
    }

    private ArrayList<String> DirectoryGetFiles(String dirname, String filenamePattern)
    {
        File dirFile = new File(dirname);
        Pattern re = Pattern.compile(filenamePattern.replace("*", ".*").replace("?", ".?"));
        ArrayList<String> dirFiles = new ArrayList<String>();
        for (File file : dirFile.listFiles())
        {
            if (file.isDirectory()) dirFiles.addAll(DirectoryGetFiles(file.getPath(), filenamePattern));
            else if (re.matcher(file.getName()).matches()) dirFiles.add(file.getPath());
        }
        return dirFiles;
    }

    private static Document save(Document inputDoc, String outputDocPath, /*SaveFormat*/int saveFormat, SaveOptions saveOptions) throws Exception
    {
        switch (saveFormat)
        {
            case SaveFormat.HTML:
                inputDoc.save(outputDocPath, saveOptions);
                return inputDoc;
            case SaveFormat.MHTML:
                inputDoc.save(outputDocPath, saveOptions);
                return inputDoc;
            case SaveFormat.EPUB:
                inputDoc.save(outputDocPath, saveOptions);
                return inputDoc;
        }

        return inputDoc;
    }

    @Test
    public void controlListLabelsExportToHtml() throws Exception
    {
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
    public void exportUrlForLinkedImage(boolean export) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportOriginalUrlForLinkedImages(export);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", saveOptions);

        ArrayList<String> dirFiles = DirectoryGetFiles(getArtifactsDir(), "HtmlSaveOptions.ExportUrlForLinkedImage.001.png");

        if (dirFiles.size() == 0)
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
        else
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportUrlForLinkedImage.html", "<img src=\"HtmlSaveOptions.ExportUrlForLinkedImage.001.png\"");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportUrlForLinkedImageDataProvider")
    public static Object[][] exportUrlForLinkedImageDataProvider() throws Exception
    {
        return new Object[][]{{true}, {false},};
    }

    @Test(enabled = false, description = "Bug, css styles starting with -aw, even if ExportRoundtripInformation is false", dataProvider = "exportRoundtripInformationDataProvider")
    public void exportRoundtripInformation(boolean valueHtml) throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportRoundtripInformation(valueHtml);

        doc.save(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html");

        if (valueHtml)
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
        else
            DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.RoundtripInformation.html", "<img src=\"HtmlSaveOptions.RoundtripInformation.003.png\" width=\"226\" height=\"132\" alt=\"\" style=\"margin-top:-53.74pt; margin-left:-26.75pt; -aw-left-pos:-26.25pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:41.25pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:1\"><img src=\"HtmlSaveOptions.RoundtripInformation.002.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:74.51pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:169.5pt; -aw-wrap-type:none; position:absolute\" /></span><span style=\"height:0pt; display:block; position:absolute; z-index:2\"><img src=\"HtmlSaveOptions.RoundtripInformation.001.png\" width=\"227\" height=\"132\" alt=\"\" style=\"margin-top:199.01pt; margin-left:-23pt; -aw-left-pos:-22.5pt; -aw-rel-hpos:column; -aw-rel-vpos:page; -aw-top-pos:294pt; -aw-wrap-type:none; position:absolute\" />");
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportRoundtripInformationDataProvider")
    public static Object[][] exportRoundtripInformationDataProvider() throws Exception
    {
        return new Object[][]{{true}, {false},};
    }

    @Test
    public void roundtripInformationDefaulValue()
    {
        //Assert that default value is true for HTML and false for MHTML and EPUB.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), true);

        saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);

        saveOptions = new HtmlSaveOptions(SaveFormat.EPUB);
        Assert.assertEquals(saveOptions.getExportRoundtripInformation(), false);
    }

    @Test
    public void configForSavingExternalResources() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ExportPageMargins.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setResourceFolder("Resources");
        saveOptions.setResourceFolderAlias("https://www.aspose.com/");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);

        ArrayList<String> imageFiles = DirectoryGetFiles(getArtifactsDir() + "Resources", "*.png");
        Assert.assertEquals(3, imageFiles.size());

        ArrayList<String> fontFiles = DirectoryGetFiles(getArtifactsDir() + "Resources", "*.ttf");
        Assert.assertEquals(1, fontFiles.size());

        ArrayList<String> cssFiles = DirectoryGetFiles(getArtifactsDir() + "Resources", "*.css");
        Assert.assertEquals(1, cssFiles.size());

        DocumentHelper.findTextInFile(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins Out.html", "<link href=\"https://www.aspose.com/HtmlSaveOptions.ExportPageMargins Out.css\"");
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

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ExportPageMargins Out.html", saveOptions);
    }

    @Test(dataProvider = "html5SupportDataProvider")
    public void html5Support(/*HtmlVersion*/int htmlVersion) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setHtmlVersion(htmlVersion);
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "html5SupportDataProvider")
    public static Object[][] html5SupportDataProvider() throws Exception
    {
        return new Object[][]{{HtmlVersion.HTML_5}, {HtmlVersion.XHTML},};
    }

    @Test(dataProvider = "exportFontsDataProvider")
    public void exportFonts(boolean exportAsBase64) throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.doc");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportFontResources(true);
        saveOptions.setExportFontsAsBase64(exportAsBase64);

        if (!exportAsBase64)
        {
            doc.save(getArtifactsDir() + "DocumentExportFonts Out 1.html", saveOptions);
            Assert.assertFalse(DirectoryGetFiles(getArtifactsDir(), "DocumentExportFonts Out 1.times.ttf").isEmpty()); //Verify that the font has been added to the folder

        } else
        {
            doc.save(getArtifactsDir() + "DocumentExportFonts Out 2.html", saveOptions);
            Assert.assertTrue(DirectoryGetFiles(getArtifactsDir(), "DocumentExportFonts Out 2.times.ttf").isEmpty()); //Verify that the font is not added to the folder

        }
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "exportFontsDataProvider")
    public static Object[][] exportFontsDataProvider() throws Exception
    {
        return new Object[][]{{false}, {true},};
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

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.001.jpeg").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.002.png").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.css").isEmpty());

    }

    @Test
    public void resourceFolderLowPriority() throws Exception
    {
        Document doc = new Document(getMyDir() + "HtmlSaveOptions.ResourceFolder.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
        saveOptions.setExportFontResources(true);
        saveOptions.setFontsFolder(getArtifactsDir() + "Fonts");
        saveOptions.setImagesFolder(getArtifactsDir() + "Images");
        saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
        saveOptions.setResourceFolderAlias("http://example.com/resources");

        doc.save(getArtifactsDir() + "HtmlSaveOptions.ResourceFolder Out.html", saveOptions);

        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolder Out.001.jpeg").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Images", "HtmlSaveOptions.ResourceFolder Out.002.png").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Fonts", "HtmlSaveOptions.ResourceFolder Out.calibri.ttf").isEmpty());
        Assert.assertFalse(DirectoryGetFiles(getArtifactsDir() + "Resources", "HtmlSaveOptions.ResourceFolder Out.css").isEmpty());
    }
}
