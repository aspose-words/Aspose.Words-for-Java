package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.pdf.TextAbsorber;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.concurrent.TimeUnit;

public class ExLoadOptions extends ApiExampleBase {
    //ExStart
    //ExFor:LoadOptions.ResourceLoadingCallback
    //ExSummary:Shows how to handle external resources when loading Html documents.
    @Test //ExSkip
    public void loadOptionsCallback() throws Exception {
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setResourceLoadingCallback(new HtmlLinkedResourceLoadingCallback());

        // When we load the document, our callback will handle linked resources such as CSS stylesheets and images.
        Document doc = new Document(getMyDir() + "Images.html", loadOptions);
        doc.save(getArtifactsDir() + "LoadOptions.LoadOptionsCallback.pdf");
    }

    /// <summary>
    /// Prints the filenames of all external stylesheets and substitutes all images of a loaded html document.
    /// </summary>
    private static class HtmlLinkedResourceLoadingCallback implements IResourceLoadingCallback {
        public int resourceLoading(ResourceLoadingArgs args) throws IOException {
            switch (args.getResourceType()) {
                case ResourceType.CSS_STYLE_SHEET:
                    System.out.println(MessageFormat.format("External CSS Stylesheet found upon loading: {0}", args.getOriginalUri()));
                    return ResourceLoadingAction.DEFAULT;
                case ResourceType.IMAGE:
                    System.out.println(MessageFormat.format("External Image found upon loading: {0}", args.getOriginalUri()));

                    final String newImageFilename = "Logo.jpg";
                    System.out.println(MessageFormat.format("\tImage will be substituted with: {0}", newImageFilename));

                    byte[] imageBytes = FileUtils.readFileToByteArray(new File(getImageDir() + newImageFilename));
                    args.setData(imageBytes);

                    return ResourceLoadingAction.USER_PROVIDED;
            }

            return ResourceLoadingAction.DEFAULT;
        }
    }
    //ExEnd

    @Test(dataProvider = "convertShapeToOfficeMathDataProvider")
    public void convertShapeToOfficeMath(boolean isConvertShapeToOfficeMath) throws Exception {
        //ExStart
        //ExFor:LoadOptions.ConvertShapeToOfficeMath
        //ExSummary:Shows how to convert EquationXML shapes to Office Math objects.
        LoadOptions loadOptions = new LoadOptions();

        // Use this flag to specify whether to convert the shapes with EquationXML attributes
        // to Office Math objects and then load the document.
        loadOptions.setConvertShapeToOfficeMath(isConvertShapeToOfficeMath);

        Document doc = new Document(getMyDir() + "Math shapes.docx", loadOptions);

        if (isConvertShapeToOfficeMath) {
            Assert.assertEquals(16, doc.getChildNodes(NodeType.SHAPE, true).getCount());
            Assert.assertEquals(34, doc.getChildNodes(NodeType.OFFICE_MATH, true).getCount());
        } else {
            Assert.assertEquals(24, doc.getChildNodes(NodeType.SHAPE, true).getCount());
            Assert.assertEquals(0, doc.getChildNodes(NodeType.OFFICE_MATH, true).getCount());
        }
        //ExEnd
    }

    @DataProvider(name = "convertShapeToOfficeMathDataProvider")
    public static Object[][] convertShapeToOfficeMathDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void setEncoding() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.Encoding
        //ExSummary:Shows how to set the encoding with which to open a document.
        LoadOptions loadOptions = new LoadOptions();
        {
            loadOptions.setEncoding(StandardCharsets.US_ASCII);
        }

        // Load the document while passing the LoadOptions object, then verify the document's contents.
        Document doc = new Document(getMyDir() + "English text.txt", loadOptions);

        Assert.assertTrue(doc.toString(SaveFormat.TEXT).contains("This is a sample text in English."));
        //ExEnd
    }

    @Test
    public void fontSettings() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to apply font substitution settings while loading a document. 
        // Create a FontSettings object that will substitute the "Times New Roman" font
        // with the font "Arvo" from our "MyFonts" folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getFontsDir(), false);
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Times New Roman", "Arvo");

        // Set that FontSettings object as a property of a newly created LoadOptions object.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);

        // Load the document, then render it as a PDF with the font substitution.
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        doc.save(getArtifactsDir() + "LoadOptions.FontSettings.pdf");
        //ExEnd
    }

    @Test
    public void loadOptionsMswVersion() throws Exception {
        //ExStart
        //ExFor:LoadOptions.MswVersion
        //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
        // By default, Aspose.Words load documents according to Microsoft Word 2019 specification.
        LoadOptions loadOptions = new LoadOptions();

        Assert.assertEquals(MsWordVersion.WORD_2019, loadOptions.getMswVersion());

        // This document is missing the default paragraph formatting style.
        // This default style will be regenerated when we load the document either with Microsoft Word or Aspose.Words.
        loadOptions.setMswVersion(MsWordVersion.WORD_2007);
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        // The style's line spacing will have this value when loaded by Microsoft Word 2007 specification.
        Assert.assertEquals(12.95d, doc.getStyles().getDefaultParagraphFormat().getLineSpacing(), 0.01d);
        //ExEnd
    }

    //ExStart
    //ExFor:LoadOptions.WarningCallback
    //ExSummary:Shows how to print and store warnings that occur during document loading.
    @Test //ExSkip
    public void loadOptionsWarningCallback() throws Exception {
        // Create a new LoadOptions object and set its WarningCallback attribute
        // as an instance of our IWarningCallback implementation.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setWarningCallback(new DocumentLoadingWarningCallback());

        // Our callback will print all warnings that come up during the load operation.
        Document doc = new Document(getMyDir() + "Document.docx", loadOptions);

        ArrayList<WarningInfo> warnings = ((DocumentLoadingWarningCallback) loadOptions.getWarningCallback()).getWarnings();
        Assert.assertEquals(3, warnings.size());
        testLoadOptionsWarningCallback(warnings); //ExSkip
    }

    /// <summary>
    /// IWarningCallback that prints warnings and their details as they arise during document loading.
    /// </summary>
    private static class DocumentLoadingWarningCallback implements IWarningCallback {
        public void warning(WarningInfo info) {
            System.out.println(MessageFormat.format("Warning: {0}", info.getWarningType()));
            System.out.println(MessageFormat.format("\tSource: {0}", info.getSource()));
            System.out.println(MessageFormat.format("\tDescription: {0}", info.getDescription()));
            mWarnings.add(info);
        }

        public ArrayList<WarningInfo> getWarnings() {
            return mWarnings;
        }

        private final /*final*/ ArrayList<WarningInfo> mWarnings = new ArrayList<WarningInfo>();
    }
    //ExEnd

    private static void testLoadOptionsWarningCallback(ArrayList<WarningInfo> warnings) {
        Assert.assertEquals(WarningType.UNEXPECTED_CONTENT, warnings.get(0).getWarningType());
        Assert.assertEquals(WarningSource.DOCX, warnings.get(0).getSource());
        Assert.assertEquals("3F01", warnings.get(0).getDescription());

        Assert.assertEquals(WarningType.MINOR_FORMATTING_LOSS, warnings.get(1).getWarningType());
        Assert.assertEquals(WarningSource.DOCX, warnings.get(1).getSource());
        Assert.assertEquals("Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings.get(1).getDescription());

        Assert.assertEquals(WarningType.MINOR_FORMATTING_LOSS, warnings.get(2).getWarningType());
        Assert.assertEquals(WarningSource.DOCX, warnings.get(2).getSource());
        Assert.assertEquals("Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings.get(2).getDescription());
    }

    @Test
    public void tempFolder() throws Exception {
        //ExStart
        //ExFor:LoadOptions.TempFolder
        //ExSummary:Shows how to use the hard drive instead of memory when loading a document.
        // When we load a document, various elements are temporarily stored in memory as the save operation occurs.
        // We can use this option to use a temporary folder in the local file system instead,
        // which will reduce our application's memory overhead.
        LoadOptions options = new LoadOptions();
        options.setTempFolder(getArtifactsDir() + "TempFiles");

        // The specified temporary folder must exist in the local file system before the load operation.
        Files.createDirectory(Paths.get(options.getTempFolder()));

        Document doc = new Document(getMyDir() + "Document.docx", options);

        // The folder will persist with no residual contents from the load operation.
        Assert.assertTrue(DocumentHelper.directoryGetFiles(options.getTempFolder(), "*.*").size() == 0);
        //ExEnd
    }

    @Test
    public void addEditingLanguage() throws Exception {
        //ExStart
        //ExFor:LanguagePreferences
        //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
        //ExFor:LoadOptions.LanguagePreferences
        //ExFor:EditingLanguage
        //ExSummary:Shows how to apply language preferences when loading a document.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().addEditingLanguage(EditingLanguage.JAPANESE);

        Document doc = new Document(getMyDir() + "No default editing language.docx", loadOptions);

        int localeIdFarEast = doc.getStyles().getDefaultFont().getLocaleIdFarEast();
        System.out.println(localeIdFarEast == EditingLanguage.JAPANESE
                ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
                : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
        //ExEnd

        Assert.assertEquals(EditingLanguage.JAPANESE, doc.getStyles().getDefaultFont().getLocaleIdFarEast());

        doc = new Document(getMyDir() + "No default editing language.docx");

        Assert.assertEquals(EditingLanguage.ENGLISH_US, doc.getStyles().getDefaultFont().getLocaleIdFarEast());
    }

    @Test
    public void setEditingLanguageAsDefault() throws Exception {
        //ExStart
        //ExFor:LanguagePreferences.DefaultEditingLanguage
        //ExSummary:Shows how set a default language when loading a document.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(getMyDir() + "No default editing language.docx", loadOptions);

        int localeId = doc.getStyles().getDefaultFont().getLocaleId();
        System.out.println(localeId == EditingLanguage.RUSSIAN
                ? "The document either has no any language set in defaults or it was set to Russian originally."
                : "The document default language was set to another than Russian language originally, so it is not overridden.");
        //ExEnd

        Assert.assertEquals(EditingLanguage.RUSSIAN, doc.getStyles().getDefaultFont().getLocaleId());

        doc = new Document(getMyDir() + "No default editing language.docx");

        Assert.assertEquals(EditingLanguage.ENGLISH_US, doc.getStyles().getDefaultFont().getLocaleId());
    }

    @Test
    public void convertMetafilesToPng() throws Exception {
        //ExStart
        //ExFor:LoadOptions.ConvertMetafilesToPng
        //ExSummary:Shows how to convert WMF/EMF to PNG during loading document.
        Document doc = new Document();

        Shape shape = new Shape(doc, ShapeType.IMAGE);
        shape.getImageData().setImage(getImageDir() + "Windows MetaFile.wmf");
        shape.setWidth(100.0);
        shape.setHeight(100.0);

        doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

        doc.save(getArtifactsDir() + "Image.CreateImageDirectly.docx");

        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.WMF, shape);

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setConvertMetafilesToPng(true);

        doc = new Document(getArtifactsDir() + "Image.CreateImageDirectly.docx", loadOptions);
        shape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);

        TestUtil.verifyImageInShape(1600, 1600, ImageType.PNG, shape);
        //ExEnd
    }

    @Test
    public void openChmFile() throws Exception {
        FileFormatInfo info = FileFormatUtil.detectFileFormat(getMyDir() + "HTML help.chm");
        Assert.assertEquals(info.getLoadFormat(), LoadFormat.CHM);

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setEncoding(Charset.forName("Windows-1251"));

        Document doc = new Document(getMyDir() + "HTML help.chm", loadOptions);
    }

    //ExStart
    //ExFor:LoadOptions.ProgressCallback
    //ExFor:IDocumentLoadingCallback
    //ExFor:IDocumentLoadingCallback.Notify
    //ExFor:DocumentLoadingArgs
    //ExFor:DocumentLoadingArgs.EstimatedProgress
    //ExSummary:Shows how to notify the user if document loading exceeded expected loading time.
    @Test//ExSkip
    public void progressCallback() throws Exception
    {
        LoadingProgressCallback progressCallback = new LoadingProgressCallback();

        LoadOptions loadOptions = new LoadOptions(); { loadOptions.setProgressCallback(progressCallback); }

        try
        {
            new Document(getMyDir() + "Big document.docx", loadOptions);
        }
        catch (IllegalStateException exception)
        {
            System.out.println(exception.getMessage());
            // Handle loading duration issue.
        }
    }

    /// <summary>
    /// Cancel a document loading after the "MaxDuration" seconds.
    /// </summary>
    public static class LoadingProgressCallback implements IDocumentLoadingCallback
    {
        /// <summary>
        /// Ctr.
        /// </summary>
        public LoadingProgressCallback()
        {
            mLoadingStartedAt = new Date();
        }

        /// <summary>
        /// Callback method which called during document loading.
        /// </summary>
        /// <param name="args">Loading arguments.</param>
        public void notify(DocumentLoadingArgs args)
        {
            Date canceledAt = new Date();
            long diff = canceledAt.getTime() - mLoadingStartedAt.getTime();
            long ellapsedSeconds = TimeUnit.MILLISECONDS.toSeconds(diff);

            if (ellapsedSeconds > MAX_DURATION)
                throw new IllegalStateException(MessageFormat.format("EstimatedProgress = {0}; CanceledAt = {1}", args.getEstimatedProgress(), canceledAt));
        }

        /// <summary>
        /// Date and time when document loading is started.
        /// </summary>
        private Date mLoadingStartedAt;

        /// <summary>
        /// Maximum allowed duration in sec.
        /// </summary>
        private static final double MAX_DURATION = 0.5;
    }
    //ExEnd

    @Test
    public void ignoreOleData() throws Exception
    {
        //ExStart
        //ExFor:LoadOptions.IgnoreOleData
        //ExSummary:Shows how to ingore OLE data while loading.
        // Ignoring OLE data may reduce memory consumption and increase performance
        // without data lost in a case when destination format does not support OLE objects.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setIgnoreOleData(true);
        Document doc = new Document(getMyDir() + "OLE objects.docx", loadOptions);

        doc.save(getArtifactsDir() + "LoadOptions.IgnoreOleData.docx");
        //ExEnd
    }
}

