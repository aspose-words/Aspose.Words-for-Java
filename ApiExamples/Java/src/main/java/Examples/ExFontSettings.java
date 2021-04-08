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
import org.apache.commons.lang.SystemUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URL;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExFontSettings extends ApiExampleBase {
    @Test
    public void defaultFontInstance() throws Exception {
        //ExStart
        //ExFor:Fonts.FontSettings.DefaultInstance
        //ExSummary:Shows how to configure the default font settings instance.
        // Configure the default font settings instance to use the "Courier New" font
        // as a backup substitute when we attempt to use an unknown font.
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Courier New");

        Assert.assertTrue(FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().getEnabled());

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Non-existent font");
        builder.write("Hello world!");

        // This document does not have a FontSettings configuration. When we render the document,
        // the default FontSettings instance will resolve the missing font.
        // Aspose.Words will use "Courier New" to render text that uses the unknown font.
        Assert.assertNull(doc.getFontSettings());

        doc.save(getArtifactsDir() + "FontSettings.DefaultFontInstance.pdf");
        //ExEnd
    }

    @Test
    public void defaultFontName() throws Exception {
        //ExStart
        //ExFor:DefaultFontSubstitutionRule.DefaultFontName
        //ExSummary:Shows how to specify a default font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Arvo");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();

        // The font sources that the document uses contain the font "Arial", but not "Arvo".
        Assert.assertEquals(1, fontSources.length);
        Assert.assertTrue(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));
        Assert.assertFalse(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arvo")));

        // Set the "DefaultFontName" property to "Courier New" to,
        // while rendering the document, apply that font in all cases when another font is not available. 
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Courier New");

        Assert.assertTrue(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Courier New")));

        // Aspose.Words will now use the default font in place of any missing fonts during any rendering calls.
        doc.save(getArtifactsDir() + "FontSettings.DefaultFontName.pdf");
        //ExEnd
    }

    @Test
    public void updatePageLayoutWarnings() throws Exception {
        // Store the font sources currently used so we can restore them later.
        FontSourceBase[] originalFontSources = FontSettings.getDefaultInstance().getFontsSources();

        // Load the document to render.
        Document doc = new Document(getMyDir() + "Document.docx");

        // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);

        // We can choose the default font to use in the case of any missing fonts.
        FontSettings.getDefaultInstance().getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");

        // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
        // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback.
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback.
        doc.updatePageLayout();

        // Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc.save(getArtifactsDir() + "FontSettings.UpdatePageLayoutWarnings.pdf");

        Assert.assertTrue(callback.FontWarnings.getCount() > 0);
        Assert.assertTrue(callback.FontWarnings.get(0).getWarningType() == WarningType.FONT_SUBSTITUTION);
        Assert.assertTrue(callback.FontWarnings.get(0).getDescription().contains("has not been found"));

        // Restore default fonts.
        FontSettings.getDefaultInstance().setFontsSources(originalFontSources);
    }

    public static class HandleDocumentWarnings implements IWarningCallback {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// potential issue during document processing. The callback can be set to listen for warnings generated during document
        /// load and/or document save.
        /// </summary>
        public void warning(WarningInfo info) {
            // We are only interested in fonts being substituted.
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION) {
                System.out.println("Font substitution: " + info.getDescription());
                FontWarnings.warning(info);
            }
        }

        public WarningInfoCollection FontWarnings = new WarningInfoCollection();
    }

    //ExStart
    //ExFor:IWarningCallback
    //ExFor:DocumentBase.WarningCallback
    //ExFor:Fonts.FontSettings.DefaultInstance
    //ExSummary:Shows how to use the IWarningCallback interface to monitor font substitution warnings.
    @Test //ExSkip
    public void substitutionWarning() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Times New Roman");
        builder.writeln("Hello world!");

        FontSubstitutionWarningCollector callback = new FontSubstitutionWarningCollector();
        doc.setWarningCallback(callback);

        // Store the current collection of font sources, which will be the default font source for every document
        // for which we do not specify a different font source.
        FontSourceBase[] originalFontSources = FontSettings.getDefaultInstance().getFontsSources();

        // For testing purposes, we will set Aspose.Words to look for fonts only in a folder that does not exist.
        FontSettings.getDefaultInstance().setFontsFolder("", false);

        // When rendering the document, there will be no place to find the "Times New Roman" font.
        // This will cause a font substitution warning, which our callback will detect.
        doc.save(getArtifactsDir() + "FontSettings.SubstitutionWarning.pdf");

        FontSettings.getDefaultInstance().setFontsSources(originalFontSources);

        Assert.assertEquals(1, callback.FontSubstitutionWarnings.getCount()); //ExSkip
        Assert.assertTrue(callback.FontSubstitutionWarnings.get(0).getWarningType() == WarningType.FONT_SUBSTITUTION);
        Assert.assertTrue(callback.FontSubstitutionWarnings.get(0).getDescription()
                .equals("Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));
    }

    private static class FontSubstitutionWarningCollector implements IWarningCallback {
        /// <summary>
        /// Called every time a warning occurs during loading/saving.
        /// </summary>
        public void warning(WarningInfo info) {
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
                FontSubstitutionWarnings.warning(info);
        }

        public WarningInfoCollection FontSubstitutionWarnings = new WarningInfoCollection();
    }
    //ExEnd

    //ExStart
    //ExFor:FontSourceBase.WarningCallback
    //ExSummary:Shows how to call warning callback when the font sources working with.
    @Test
    public void fontSourceWarning()
    {
        FontSettings settings = new FontSettings();
        settings.setFontsFolder("bad folder?", false);

        FontSourceBase source = settings.getFontsSources()[0];
        FontSourceWarningCollector callback = new FontSourceWarningCollector();
        source.setWarningCallback(callback);

        // Get the list of fonts to call warning callback.
        ArrayList<PhysicalFontInfo> fontInfos = source.getAvailableFonts();

        Assert.assertEquals("Error loading font from the folder \"bad folder?\": ",
            callback.FontSubstitutionWarnings.get(0).getDescription());
    }

    private static class FontSourceWarningCollector implements IWarningCallback
    {
        /// <summary>
        /// Called every time a warning occurs during processing of font source.
        /// </summary>
        public void warning(WarningInfo info)
        {
            FontSubstitutionWarnings.warning(info);
        }

        public WarningInfoCollection FontSubstitutionWarnings = new WarningInfoCollection();
    }
    //ExEnd

    //ExStart
    //ExFor:Fonts.FontInfoSubstitutionRule
    //ExFor:Fonts.FontSubstitutionSettings.FontInfoSubstitution
    //ExFor:IWarningCallback
    //ExFor:IWarningCallback.Warning(WarningInfo)
    //ExFor:WarningInfo
    //ExFor:WarningInfo.Description
    //ExFor:WarningInfo.WarningType
    //ExFor:WarningInfoCollection
    //ExFor:WarningInfoCollection.Warning(WarningInfo)
    //ExFor:WarningInfoCollection.GetEnumerator
    //ExFor:WarningInfoCollection.Clear
    //ExFor:WarningType
    //ExFor:DocumentBase.WarningCallback
    //ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
    @Test
    public void enableFontSubstitution() throws Exception {
        // Open a document that contains text formatted with a font that does not exist in any of our font sources.
        Document doc = new Document(getMyDir() + "Missing font.docx");

        // Assign a callback for handling font substitution warnings.
        HandleDocumentSubstitutionWarnings substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(substitutionWarningHandler);

        // Set a default font name and enable font substitution.
        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(true);

        // We will get a font substitution warning if we save a document with a missing font.
        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "FontSettings.EnableFontSubstitution.pdf");

        Iterator<WarningInfo> warnings = substitutionWarningHandler.FontWarnings.iterator();

        while (warnings.hasNext())
            System.out.println(warnings.next().getDescription());

        // We can also verify warnings in the collection and clear them.
        Assert.assertEquals(WarningSource.LAYOUT, substitutionWarningHandler.FontWarnings.get(0).getSource());
        Assert.assertEquals("Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.",
                substitutionWarningHandler.FontWarnings.get(0).getDescription());

        substitutionWarningHandler.FontWarnings.clear();

        Assert.assertTrue(substitutionWarningHandler.FontWarnings.getCount() == 0);
    }

    public static class HandleDocumentSubstitutionWarnings implements IWarningCallback {
        /// <summary>
        /// Called every time a warning occurs during loading/saving.
        /// </summary>
        public void warning(WarningInfo info) {
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
                FontWarnings.warning(info);
        }

        public WarningInfoCollection FontWarnings = new WarningInfoCollection();
    }
    //ExEnd

    @Test
    public void substitutionWarningsClosestMatch() throws Exception {
        Document doc = new Document(getMyDir() + "Bullet points with alternative font.docx");

        HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(callback);

        doc.save(getArtifactsDir() + "FontSettings.SubstitutionWarningsClosestMatch.pdf");

        Assert.assertTrue(callback.FontWarnings.get(0).getDescription()
                .equals("Font 'SymbolPS' has not been found. Using 'Wingdings' font instead. Reason: font info substitution."));
    }

    @Test
    public void disableFontSubstitution() throws Exception {
        Document doc = new Document(getMyDir() + "Missing font.docx");

        HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(callback);

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(false);

        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "FontSettings.DisableFontSubstitution.pdf");

        Pattern pattern = Pattern.compile("Font '28 Days Later' has not been found. Using (.*) font instead. Reason: alternative name from document.");

        for (WarningInfo fontWarning : callback.FontWarnings) {
            Matcher match = pattern.matcher(fontWarning.getDescription());
            if (match.find() == false) {
                Assert.fail();
                break;
            }
        }
    }

    @Test
    public void substitutionWarnings() throws Exception {
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(callback);

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.setFontsFolder(getFontsDir(), false);
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Arial", "Arvo", "Slab");

        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "FontSettings.SubstitutionWarnings.pdf");

        Assert.assertEquals("Font 'Arial' has not been found. Using 'Arvo' font instead. Reason: table substitution.",
                callback.FontWarnings.get(0).getDescription());
        Assert.assertEquals("Font 'Times New Roman' has not been found. Using 'M+ 2m' font instead. Reason: font info substitution.",
                callback.FontWarnings.get(1).getDescription());
    }

    @Test
    public void fontSourceFile() throws Exception {
        //ExStart
        //ExFor:Fonts.FileFontSource
        //ExFor:Fonts.FileFontSource.#ctor(String)
        //ExFor:Fonts.FileFontSource.#ctor(String, Int32)
        //ExFor:Fonts.FileFontSource.FilePath
        //ExFor:Fonts.FileFontSource.Type
        //ExFor:Fonts.FontSourceBase
        //ExFor:Fonts.FontSourceBase.Priority
        //ExFor:Fonts.FontSourceBase.Type
        //ExFor:Fonts.FontSourceType
        //ExSummary:Shows how to use a font file in the local file system as a font source.
        FileFontSource fileFontSource = new FileFontSource(getMyDir() + "Alte DIN 1451 Mittelschrift.ttf", 0);

        Document doc = new Document();
        doc.setFontSettings(new FontSettings());
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{fileFontSource});

        Assert.assertEquals(getMyDir() + "Alte DIN 1451 Mittelschrift.ttf", fileFontSource.getFilePath());
        Assert.assertEquals(FontSourceType.FONT_FILE, fileFontSource.getType());
        Assert.assertEquals(0, fileFontSource.getPriority());
        //ExEnd
    }

    @Test
    public void fontSourceFolder() throws Exception {
        //ExStart
        //ExFor:Fonts.FolderFontSource
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean)
        //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean, Int32)
        //ExFor:Fonts.FolderFontSource.FolderPath
        //ExFor:Fonts.FolderFontSource.ScanSubfolders
        //ExFor:Fonts.FolderFontSource.Type
        //ExSummary:Shows how to use a local system folder which contains fonts as a font source.

        // Create a font source from a folder that contains font files.
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false, 1);

        Document doc = new Document();
        doc.setFontSettings(new FontSettings());
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{folderFontSource});

        Assert.assertEquals(getFontsDir(), folderFontSource.getFolderPath());
        Assert.assertEquals(false, folderFontSource.getScanSubfolders());
        Assert.assertEquals(FontSourceType.FONTS_FOLDER, folderFontSource.getType());
        Assert.assertEquals(1, folderFontSource.getPriority());
        //ExEnd
    }

    @Test(dataProvider = "setFontsFolderDataProvider")
    public void setFontsFolder(boolean recursive) throws Exception {
        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolder(String, Boolean)
        //ExSummary:Shows how to set a font source directory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arvo");
        builder.writeln("Hello world!");
        builder.getFont().setName("Amethysta");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        // Our font sources do not contain the font that we have used for text in this document.
        // If we use these font settings while rendering this document,
        // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
        FontSourceBase[] originalFontSources = FontSettings.getDefaultInstance().getFontsSources();

        Assert.assertEquals(1, originalFontSources.length);
        Assert.assertTrue(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));

        // The default font sources are missing the two fonts that we are using in this document.
        Assert.assertFalse(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arvo")));
        Assert.assertFalse(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));

        // Use the "SetFontsFolder" method to set a directory which will act as a new font source.
        // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
        // that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
        // Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
        // in the first argument, as well as all the fonts in its subdirectories.
        FontSettings.getDefaultInstance().setFontsFolder(getFontsDir(), recursive);

        FontSourceBase[] newFontSources = FontSettings.getDefaultInstance().getFontsSources();

        Assert.assertEquals(1, newFontSources.length);
        Assert.assertFalse(IterableUtils.matchesAny(newFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));
        Assert.assertTrue(IterableUtils.matchesAny(newFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arvo")));

        // The "Amethysta" font is in a subfolder of the font directory.
        if (recursive) {
            Assert.assertEquals(25, newFontSources[0].getAvailableFonts().size());
            Assert.assertTrue(IterableUtils.matchesAny(newFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));
        } else {
            Assert.assertEquals(18, newFontSources[0].getAvailableFonts().size());
            Assert.assertFalse(IterableUtils.matchesAny(newFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));
        }

        doc.save(getArtifactsDir() + "FontSettings.SetFontsFolder.pdf");

        // Restore the original font sources.
        FontSettings.getDefaultInstance().setFontsSources(originalFontSources);
        //ExEnd
    }

    @DataProvider(name = "setFontsFolderDataProvider")
    public static Object[][] setFontsFolderDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "setFontsFoldersDataProvider")
    public void setFontsFolders(boolean recursive) throws Exception {
        //ExStart
        //ExFor:FontSettings
        //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
        //ExSummary:Shows how to set multiple font source directories.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Amethysta");
        builder.writeln("The quick brown fox jumps over the lazy dog.");
        builder.getFont().setName("Junction Light");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        // Our font sources do not contain the font that we have used for text in this document.
        // If we use these font settings while rendering this document,
        // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
        FontSourceBase[] originalFontSources = FontSettings.getDefaultInstance().getFontsSources();

        Assert.assertEquals(1, originalFontSources.length);
        Assert.assertTrue(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));

        // The default font sources are missing the two fonts that we are using in this document.
        Assert.assertFalse(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));
        Assert.assertFalse(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Junction Light")));

        // Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
        // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
        // that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
        // Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
        // in the first argument, as well as all the fonts in their subdirectories.
        FontSettings.getDefaultInstance().setFontsFolders(new String[]{getFontsDir() + "/Amethysta", getFontsDir() + "/Junction"}, recursive);

        FontSourceBase[] newFontSources = FontSettings.getDefaultInstance().getFontsSources();

        Assert.assertEquals(2, newFontSources.length);
        Assert.assertFalse(IterableUtils.matchesAny(newFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));
        Assert.assertEquals(1, newFontSources[0].getAvailableFonts().size());
        Assert.assertTrue(IterableUtils.matchesAny(newFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));

        // The "Junction" folder itself contains no font files, but has subfolders that do.
        if (recursive) {
            Assert.assertEquals(6, newFontSources[1].getAvailableFonts().size());
            Assert.assertTrue(IterableUtils.matchesAny(newFontSources[1].getAvailableFonts(), f -> f.getFullFontName().contains("Junction Light")));
        } else {
            Assert.assertEquals(0, newFontSources[1].getAvailableFonts().size());
        }

        doc.save(getArtifactsDir() + "FontSettings.SetFontsFolders.pdf");

        // Restore the original font sources.
        FontSettings.getDefaultInstance().setFontsSources(originalFontSources);
        //ExEnd
    }

    @DataProvider(name = "setFontsFoldersDataProvider")
    public static Object[][] setFontsFoldersDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void addFontSource() throws Exception {
        //ExStart
        //ExFor:FontSettings            
        //ExFor:FontSettings.GetFontsSources()
        //ExFor:FontSettings.SetFontsSources()
        //ExSummary:Shows how to add a font source to our existing font sources.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Amethysta");
        builder.writeln("The quick brown fox jumps over the lazy dog.");
        builder.getFont().setName("Junction Light");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        FontSourceBase[] originalFontSources = FontSettings.getDefaultInstance().getFontsSources();

        Assert.assertEquals(1, originalFontSources.length);

        Assert.assertTrue(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));

        // The default font source is missing two of the fonts that we are using in our document.
        // When we save this document, Aspose.Words will apply fallback fonts to all text formatted with inaccessible fonts.
        Assert.assertFalse(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));
        Assert.assertFalse(IterableUtils.matchesAny(originalFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Junction Light")));

        // Create a font source from a folder that contains fonts.
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), true);

        // Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
        FontSourceBase[] updatedFontSources = {originalFontSources[0], folderFontSource};
        FontSettings.getDefaultInstance().setFontsSources(updatedFontSources);

        // Verify that Aspose.Words has access to all required fonts before we render the document to PDF.
        updatedFontSources = FontSettings.getDefaultInstance().getFontsSources();

        Assert.assertTrue(IterableUtils.matchesAny(updatedFontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));
        Assert.assertTrue(IterableUtils.matchesAny(updatedFontSources[1].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));
        Assert.assertTrue(IterableUtils.matchesAny(updatedFontSources[1].getAvailableFonts(), f -> f.getFullFontName().contains("Junction Light")));

        doc.save(getArtifactsDir() + "FontSettings.AddFontSource.pdf");

        // Restore the original font sources.
        FontSettings.getDefaultInstance().setFontsSources(originalFontSources);
        //ExEnd
    }

    @Test
    public void setSpecifyFontFolder() throws Exception {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getFontsDir(), false);

        // Using load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);

        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);

        FolderFontSource folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[0]);

        Assert.assertEquals(getFontsDir(), folderSource.getFolderPath());
        Assert.assertFalse(folderSource.getScanSubfolders());
    }

    @Test
    public void tableSubstitution() throws Exception {
        //ExStart
        //ExFor:Document.FontSettings
        //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
        //ExSummary:Shows how set font substitution rules.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Hello world!");
        builder.getFont().setName("Amethysta");
        builder.writeln("The quick brown fox jumps over the lazy dog.");

        FontSourceBase[] fontSources = FontSettings.getDefaultInstance().getFontsSources();

        // The default font sources contain the first font that the document uses.
        Assert.assertEquals(1, fontSources.length);
        Assert.assertTrue(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arial")));

        // The second font, "Amethysta", is unavailable.
        Assert.assertFalse(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Amethysta")));

        // We can configure a font substitution table which determines
        // which fonts Aspose.Words will use as substitutes for unavailable fonts.
        // Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
        // If the first substitute is unavailable, Aspose.Words attempts to use the second substitute, and so on.
        doc.setFontSettings(new FontSettings());
        doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().setSubstitutes(
                "Amethysta", "Arvo", "Courier New");

        // "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo". 
        Assert.assertFalse(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Arvo")));

        // "Arvo" is also unavailable, but "Courier New" is. 
        Assert.assertTrue(IterableUtils.matchesAny(fontSources[0].getAvailableFonts(), f -> f.getFullFontName().contains("Courier New")));

        // The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
        doc.save(getArtifactsDir() + "FontSettings.TableSubstitution.pdf");
        //ExEnd
    }

    @Test
    public void setSpecifyFontFolders() throws Exception {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolders(new String[]{getFontsDir(), "C:\\Windows\\Fonts\\"}, true);

        // Using load options
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);
        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);

        FolderFontSource folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[0]);
        Assert.assertEquals(getFontsDir(), folderSource.getFolderPath());
        Assert.assertTrue(folderSource.getScanSubfolders());

        folderSource = ((FolderFontSource) doc.getFontSettings().getFontsSources()[1]);
        Assert.assertEquals("C:\\Windows\\Fonts\\", folderSource.getFolderPath());
        Assert.assertTrue(folderSource.getScanSubfolders());
    }

    @Test
    public void addFontSubstitutes() throws Exception {
        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getTableSubstitution().setSubstitutes("Slab", "Times New Roman", "Arial");
        fontSettings.getSubstitutionSettings().getTableSubstitution().addSubstitutes("Arvo", "Open Sans", "Arial");

        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.setFontSettings(fontSettings);

        Iterable<String> alternativeFonts = doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Slab");
        Assert.assertEquals(new String[]{"Times New Roman", "Arial"}, IterableUtils.toList(alternativeFonts).toArray());

        alternativeFonts = doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Arvo");
        Assert.assertEquals(new String[]{"Open Sans", "Arial"}, IterableUtils.toList(alternativeFonts).toArray());
    }

    @Test
    public void fontSourceMemory() throws Exception {
        //ExStart
        //ExFor:Fonts.MemoryFontSource
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[])
        //ExFor:Fonts.MemoryFontSource.#ctor(Byte[], Int32)
        //ExFor:Fonts.MemoryFontSource.FontData
        //ExFor:Fonts.MemoryFontSource.Type
        //ExSummary:Shows how to use a byte array with data from a font file as a font source.

        byte[] fontBytes = DocumentHelper.getBytesFromStream(new FileInputStream(getMyDir() + "Alte DIN 1451 Mittelschrift.ttf"));
        MemoryFontSource memoryFontSource = new MemoryFontSource(fontBytes, 0);

        Document doc = new Document();
        doc.setFontSettings(new FontSettings());
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{memoryFontSource});

        Assert.assertEquals(FontSourceType.MEMORY_FONT, memoryFontSource.getType());
        Assert.assertEquals(0, memoryFontSource.getPriority());
        //ExEnd
    }

    @Test
    public void fontSourceSystem() throws Exception {
        //ExStart
        //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
        //ExFor:FontSubstitutionRule.Enabled
        //ExFor:TableSubstitutionRule.GetSubstitutes(String)
        //ExFor:Fonts.FontSettings.ResetFontSources
        //ExFor:Fonts.FontSettings.SubstitutionSettings
        //ExFor:Fonts.FontSubstitutionSettings
        //ExFor:Fonts.SystemFontSource
        //ExFor:Fonts.SystemFontSource.#ctor
        //ExFor:Fonts.SystemFontSource.#ctor(Int32)
        //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
        //ExFor:Fonts.SystemFontSource.Type
        //ExSummary:Shows how to access a document's system font source and set font substitutes.
        Document doc = new Document();
        doc.setFontSettings(new FontSettings());

        // By default, a blank document always contains a system font source.
        Assert.assertEquals(1, doc.getFontSettings().getFontsSources().length);

        SystemFontSource systemFontSource = (SystemFontSource) doc.getFontSettings().getFontsSources()[0];
        Assert.assertEquals(FontSourceType.SYSTEM_FONTS, systemFontSource.getType());
        Assert.assertEquals(0, systemFontSource.getPriority());

        if (SystemUtils.IS_OS_WINDOWS) {
            final String FONTS_PATH = "C:\\WINDOWS\\Fonts";
            Assert.assertEquals(FONTS_PATH.toLowerCase(), SystemFontSource.getSystemFontFolders()[0].toLowerCase());
        }

        for (String systemFontFolder : SystemFontSource.getSystemFontFolders()) {
            System.out.println(systemFontFolder);
        }

        // Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
        doc.getFontSettings().getSubstitutionSettings().getFontInfoSubstitution().setEnabled(true);
        doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().addSubstitutes("Kreon-Regular", "Calibri");

        Assert.assertEquals(1, IterableUtils.size(doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Kreon-Regular")));
        Assert.assertTrue(IterableUtils.toString(doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Kreon-Regular")).contains("Calibri"));

        // Alternatively, we could add a folder font source in which the corresponding folder contains the font.
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false);
        doc.getFontSettings().setFontsSources(new FontSourceBase[]{systemFontSource, folderFontSource});
        Assert.assertEquals(2, doc.getFontSettings().getFontsSources().length);

        // Resetting the font sources still leaves us with the system font source as well as our substitutes.
        doc.getFontSettings().resetFontSources();

        Assert.assertEquals(1, doc.getFontSettings().getFontsSources().length);
        Assert.assertEquals(FontSourceType.SYSTEM_FONTS, doc.getFontSettings().getFontsSources()[0].getType());
        Assert.assertEquals(1, IterableUtils.size(doc.getFontSettings().getSubstitutionSettings().getTableSubstitution().getSubstitutes("Kreon-Regular")));
        //ExEnd
    }

    @Test
    public void loadFontFallbackSettingsFromFile() throws Exception {
        //ExStart
        //ExFor:FontFallbackSettings.Load(String)
        //ExFor:FontFallbackSettings.Save(String)
        //ExSummary:Shows how to load and save font fallback settings to/from an XML document in the local file system.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Load an XML document that defines a set of font fallback settings.
        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().load(getMyDir() + "Font fallback rules.xml");

        doc.setFontSettings(fontSettings);
        doc.save(getArtifactsDir() + "FontSettings.LoadFontFallbackSettingsFromFile.pdf");

        // Save our document's current font fallback settings as an XML document.
        doc.getFontSettings().getFallbackSettings().save(getArtifactsDir() + "FallbackSettings.xml");
        //ExEnd
    }

    @Test
    public void loadFontFallbackSettingsFromStream() throws Exception {
        //ExStart
        //ExFor:FontFallbackSettings.Load(Stream)
        //ExFor:FontFallbackSettings.Save(Stream)
        //ExSummary:Shows how to load and save font fallback settings to/from a stream.
        Document doc = new Document(getMyDir() + "Rendering.docx");

        // Load an XML document that defines a set of font fallback settings.
        try (FileInputStream fontFallbackRulesStream = new FileInputStream(getMyDir() + "Font fallback rules.xml")) {
            FontSettings fontSettings = new FontSettings();
            fontSettings.getFallbackSettings().load(fontFallbackRulesStream);

            doc.setFontSettings(fontSettings);
        }

        doc.save(getArtifactsDir() + "FontSettings.LoadFontFallbackSettingsFromStream.pdf");

        // Use a stream to save our document's current font fallback settings as an XML document.
        try (FileOutputStream fontFallbackSettingsStream = new FileOutputStream(getArtifactsDir() + "FallbackSettings.xml")) {
            doc.getFontSettings().getFallbackSettings().save(fontFallbackSettingsStream);
        }
        //ExEnd
    }

    @Test
    public void loadNotoFontsFallbackSettings() throws Exception {
        //ExStart
        //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
        FontSettings fontSettings = new FontSettings();

        // These are free fonts licensed under the SIL Open Font License.
        // We can download the fonts here:
        // https://www.google.com/get/noto/#sans-lgc
        fontSettings.setFontsFolder(getFontsDir() + "Noto", false);

        // Note that the predefined settings only use Sans-style Noto fonts with regular weight. 
        // Some of the Noto fonts use advanced typography features.
        // Fonts featuring advanced typography may not be rendered correctly as Aspose.Words currently do not support them.
        fontSettings.getFallbackSettings().loadNotoFallbackSettings();
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(false);
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Noto Sans");

        Document doc = new Document();
        doc.setFontSettings(fontSettings);
        //ExEnd

        TestUtil.verifyWebResponseStatusCode(200, new URL("https://www.google.com/get/noto/#sans-lgc"));
    }

    @Test
    public void defaultFontSubstitutionRule() throws Exception {
        //ExStart
        //ExFor:Fonts.DefaultFontSubstitutionRule
        //ExFor:Fonts.DefaultFontSubstitutionRule.DefaultFontName
        //ExFor:Fonts.FontSubstitutionSettings.DefaultFontSubstitution
        //ExSummary:Shows how to set the default font substitution rule.
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);

        // Get the default substitution rule within FontSettings.
        // This rule will substitute all missing fonts with "Times New Roman".
        DefaultFontSubstitutionRule defaultFontSubstitutionRule = fontSettings.getSubstitutionSettings().getDefaultFontSubstitution();
        Assert.assertTrue(defaultFontSubstitutionRule.getEnabled());
        Assert.assertEquals("Times New Roman", defaultFontSubstitutionRule.getDefaultFontName());

        // Set the default font substitute to "Courier New".
        defaultFontSubstitutionRule.setDefaultFontName("Courier New");

        // Using a document builder, add some text in a font that we do not have to see the substitution take place,
        // and then render the result in a PDF.
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Missing Font");
        builder.writeln("Line written in a missing font, which will be substituted with Courier New.");

        doc.save(getArtifactsDir() + "FontSettings.DefaultFontSubstitutionRule.pdf");
        //ExEnd

        Assert.assertEquals("Courier New", defaultFontSubstitutionRule.getDefaultFontName());
    }

    @Test
    public void fontConfigSubstitution() {
        //ExStart
        //ExFor:Fonts.FontConfigSubstitutionRule
        //ExFor:Fonts.FontConfigSubstitutionRule.Enabled
        //ExFor:Fonts.FontConfigSubstitutionRule.IsFontConfigAvailable
        //ExFor:Fonts.FontConfigSubstitutionRule.ResetCache
        //ExFor:Fonts.FontSubstitutionRule
        //ExFor:Fonts.FontSubstitutionRule.Enabled
        //ExFor:Fonts.FontSubstitutionSettings.FontConfigSubstitution
        //ExSummary:Shows operating system-dependent font config substitution.
        FontSettings fontSettings = new FontSettings();
        FontConfigSubstitutionRule fontConfigSubstitution = fontSettings.getSubstitutionSettings().getFontConfigSubstitution();

        // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
        // On Windows, it is unavailable.
        if (SystemUtils.IS_OS_WINDOWS) {
            Assert.assertFalse(fontConfigSubstitution.getEnabled());
            Assert.assertFalse(fontConfigSubstitution.isFontConfigAvailable());
        }

        // On Linux/Mac, we will have access to it, and will be able to perform operations.
        if (SystemUtils.IS_OS_LINUX) {
            Assert.assertTrue(fontConfigSubstitution.getEnabled());
            Assert.assertTrue(fontConfigSubstitution.isFontConfigAvailable());

            fontConfigSubstitution.resetCache();
        }
        //ExEnd
    }

    @Test
    public void fallbackSettings() throws Exception {
        //ExStart
        //ExFor:Fonts.FontFallbackSettings.LoadMsOfficeFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.LoadNotoFallbackSettings
        //ExSummary:Shows how to load pre-defined fallback font settings.
        Document doc = new Document();

        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);
        FontFallbackSettings fontFallbackSettings = fontSettings.getFallbackSettings();

        // Save the default fallback font scheme to an XML document.
        // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts.
        // This means that if the font some text is using does not have symbols for the 0x0C00-0x0C7F Unicode block,
        // the fallback scheme will use symbols from the "Vani" font substitute.
        fontFallbackSettings.save(getArtifactsDir() + "FontSettings.FallbackSettings.Default.xml");

        // Below are two pre-defined font fallback schemes we can choose from.
        // 1 -  Use the default Microsoft Office scheme, which is the same one as the default:
        fontFallbackSettings.loadMsOfficeFallbackSettings();
        fontFallbackSettings.save(getArtifactsDir() + "FontSettings.FallbackSettings.LoadMsOfficeFallbackSettings.xml");

        // 2 -  Use the scheme built from Google Noto fonts:
        fontFallbackSettings.loadNotoFallbackSettings();
        fontFallbackSettings.save(getArtifactsDir() + "FontSettings.FallbackSettings.LoadNotoFallbackSettings.xml");
        //ExEnd
    }

    @Test
    public void fallbackSettingsCustom() throws Exception {
        //ExStart
        //ExFor:Fonts.FontSettings.FallbackSettings
        //ExFor:Fonts.FontFallbackSettings
        //ExFor:Fonts.FontFallbackSettings.BuildAutomatic
        //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
        Document doc = new Document();

        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);
        FontFallbackSettings fontFallbackSettings = fontSettings.getFallbackSettings();

        // Configure our font settings to source fonts only from the "MyFonts" folder.
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false);
        fontSettings.setFontsSources(new FontSourceBase[]{folderFontSource});

        // Calling the "BuildAutomatic" method will generate a fallback scheme that
        // distributes accessible fonts across as many Unicode character codes as possible.
        // In our case, it only has access to the handful of fonts inside the "MyFonts" folder.
        fontFallbackSettings.buildAutomatic();
        fontFallbackSettings.save(getArtifactsDir() + "FontSettings.FallbackSettingsCustom.BuildAutomatic.xml");

        // We can also load a custom substitution scheme from a file like this.
        // This scheme applies the "AllegroOpen" font across the "0000-00ff" Unicode blocks, the "AllegroOpen" font across "0100-024f",
        // and the "M+ 2m" font in all other ranges that other fonts in the scheme do not cover.
        fontFallbackSettings.load(getMyDir() + "Custom font fallback settings.xml");

        // Create a document builder and set its font to one that does not exist in any of our sources.
        // Our font settings will invoke the fallback scheme for characters that we type using the unavailable font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Missing Font");

        // Use the builder to print every Unicode character from 0x0021 to 0x052F,
        // with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme.
        for (int i = 0x0021; i < 0x0530; i++) {
            switch (i) {
                case 0x0021:
                    builder.writeln("\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"AllegroOpen\" font:");
                    break;
                case 0x0100:
                    builder.writeln("\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"AllegroOpen\" font:");
                    break;
                case 0x0250:
                    builder.writeln("\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
                    break;
            }

            builder.write(MessageFormat.format("{0}", (char) i));
        }

        doc.save(getArtifactsDir() + "FontSettings.FallbackSettingsCustom.pdf");
        //ExEnd
    }

    @Test
    public void tableSubstitutionRule() throws Exception {
        //ExStart
        //ExFor:Fonts.TableSubstitutionRule
        //ExFor:Fonts.TableSubstitutionRule.LoadLinuxSettings
        //ExFor:Fonts.TableSubstitutionRule.LoadWindowsSettings
        //ExFor:Fonts.TableSubstitutionRule.Save(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Save(System.String)
        //ExSummary:Shows how to access font substitution tables for Windows and Linux.
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);

        // Create a new table substitution rule and load the default Microsoft Windows font substitution table.
        TableSubstitutionRule tableSubstitutionRule = fontSettings.getSubstitutionSettings().getTableSubstitution();
        tableSubstitutionRule.loadWindowsSettings();

        // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
        Assert.assertEquals(new String[]{"Times New Roman"},
                IterableUtils.toList(tableSubstitutionRule.getSubstitutes("Times New Roman CE")).toArray());

        // We can save the table in the form of an XML document.
        tableSubstitutionRule.save(getArtifactsDir() + "FontSettings.TableSubstitutionRule.Windows.xml");

        // Linux has its own substitution table.
        // There are multiple substitute fonts for "Times New Roman CE".
        // If the first substitute, "FreeSerif" is also unavailable,
        // this rule will cycle through the others in the array until it finds an available one.
        tableSubstitutionRule.loadLinuxSettings();
        Assert.assertEquals(new String[]{"FreeSerif", "Liberation Serif", "DejaVu Serif"},
                IterableUtils.toList(tableSubstitutionRule.getSubstitutes("Times New Roman CE")).toArray());

        // Save the Linux substitution table in the form of an XML document using a stream.
        try (FileOutputStream fileStream = new FileOutputStream(getArtifactsDir() + "FontSettings.TableSubstitutionRule.Linux.xml")) {
            tableSubstitutionRule.save(fileStream);
        }
        //ExEnd
    }

    @Test
    public void tableSubstitutionRuleCustom() throws Exception {
        //ExStart
        //ExFor:Fonts.FontSubstitutionSettings.TableSubstitution
        //ExFor:Fonts.TableSubstitutionRule.AddSubstitutes(System.String,System.String[])
        //ExFor:Fonts.TableSubstitutionRule.GetSubstitutes(System.String)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.IO.Stream)
        //ExFor:Fonts.TableSubstitutionRule.Load(System.String)
        //ExFor:Fonts.TableSubstitutionRule.SetSubstitutes(System.String,System.String[])
        //ExSummary:Shows how to work with custom font substitution tables.
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.setFontSettings(fontSettings);

        // Create a new table substitution rule and load the default Windows font substitution table.
        TableSubstitutionRule tableSubstitutionRule = fontSettings.getSubstitutionSettings().getTableSubstitution();

        // If we select fonts exclusively from our folder, we will need a custom substitution table.
        // We will no longer have access to the Microsoft Windows fonts,
        // such as "Arial" or "Times New Roman" since they do not exist in our new font folder.
        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), false);
        fontSettings.setFontsSources(new FontSourceBase[]{folderFontSource});

        // Below are two ways of loading a substitution table from a file in the local file system.
        // 1 -  From a stream:
        try (FileInputStream fileStream = new FileInputStream(getMyDir() + "Font substitution rules.xml")) {
            tableSubstitutionRule.load(fileStream);
        }

        // 2 -  Directly from a file:
        tableSubstitutionRule.load(getMyDir() + "Font substitution rules.xml");

        // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font".
        // We do not have this font so that it will move onto the next substitute, "Kreon", found in the "MyFonts" folder.
        Assert.assertEquals(new String[]{"Missing Font", "Kreon"}, IterableUtils.toList(tableSubstitutionRule.getSubstitutes("Arial")).toArray());

        // We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
        Assert.assertNull(tableSubstitutionRule.getSubstitutes("Times New Roman"));
        tableSubstitutionRule.addSubstitutes("Times New Roman", "Arvo");
        Assert.assertEquals(new String[]{"Arvo"}, IterableUtils.toList(tableSubstitutionRule.getSubstitutes("Times New Roman")).toArray());

        // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
        // In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
        tableSubstitutionRule.addSubstitutes("Times New Roman", "M+ 2m");
        Assert.assertEquals(new String[]{"Arvo", "M+ 2m"}, IterableUtils.toList(tableSubstitutionRule.getSubstitutes("Times New Roman")).toArray());

        // SetSubstitutes() can set a new list of substitute fonts for a font.
        tableSubstitutionRule.setSubstitutes("Times New Roman", "Squarish Sans CT", "M+ 2m");
        Assert.assertEquals(new String[]{"Squarish Sans CT", "M+ 2m"}, IterableUtils.toList(tableSubstitutionRule.getSubstitutes("Times New Roman")).toArray());

        // Writing text in fonts that we do not have access to will invoke our substitution rules.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setName("Arial");
        builder.writeln("Text written in Arial, to be substituted by Kreon.");

        builder.getFont().setName("Times New Roman");
        builder.writeln("Text written in Times New Roman, to be substituted by Squarish Sans CT.");

        doc.save(getArtifactsDir() + "FontSettings.TableSubstitutionRule.Custom.pdf");
        //ExEnd
    }

    @Test
    public void resolveFontsBeforeLoadingDocument() throws Exception {
        //ExStart
        //ExFor:LoadOptions.FontSettings
        //ExSummary:Shows how to designate font substitutes during loading.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(new FontSettings());

        // Set a font substitution rule for a LoadOptions object.
        // If the document we are loading uses a font which we do not have,
        // this rule will substitute the unavailable font with one that does exist.
        // In this case, all uses of the "MissingFont" will convert to "Comic Sans MS".
        TableSubstitutionRule substitutionRule = loadOptions.getFontSettings().getSubstitutionSettings().getTableSubstitution();
        substitutionRule.addSubstitutes("MissingFont", "Comic Sans MS");

        Document doc = new Document(getMyDir() + "Missing font.html", loadOptions);

        // At this point such text will still be in "MissingFont".
        // Font substitution will take place when we render the document.
        Assert.assertEquals("MissingFont", doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getFont().getName());

        doc.save(getArtifactsDir() + "FontSettings.ResolveFontsBeforeLoadingDocument.pdf");
        //ExEnd
    }

    //ExStart
    //ExFor:StreamFontSource
    //ExFor:StreamFontSource.OpenFontDataStream
    //ExSummary:Shows how to load fonts from stream.
    @Test //ExSkip
    public void streamFontSourceFileRendering() throws Exception {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsSources(new FontSourceBase[]{new StreamFontSourceFile()});

        DocumentBuilder builder = new DocumentBuilder();
        builder.getDocument().setFontSettings(fontSettings);
        builder.getFont().setName("Kreon-Regular");
        builder.writeln("Test aspose text when saving to PDF.");

        builder.getDocument().save(getArtifactsDir() + "FontSettings.StreamFontSourceFileRendering.pdf");
    }

    /// <summary>
    /// Load the font data only when required instead of storing it in the memory for the entire lifetime of the "FontSettings" object.
    /// </summary>
    private static class StreamFontSourceFile extends StreamFontSource  {
        public FileInputStream openFontDataStream() throws Exception {
            return new FileInputStream(getFontsDir() + "Kreon-Regular.ttf");
        }
    }
    //ExEnd
}