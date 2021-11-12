package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;

public class SpecifyTrueTypeFontsLocation {

    private static final String dataDir = Utils.getSharedDataDir(SpecifyTrueTypeFontsLocation.class) + "RenderingAndPrinting/";

    public static void main(String[] args) throws Exception {
        // Specifying a Font Folder
        specifyAFontFolder();

        // Specifying Multiple Font Folders
        specifyMultipleFontFolders();

        // Specifying Fonts to be Read from both the System Fonts Folder and a Custom Folder
        specifyFontsFromBothSystemFontsFolderAndCustomFolder();
    }

    public static void specifyAFontFolder() throws Exception {
        //ExStart:SpecifyAFontFolder
        Document doc = new Document(dataDir + "Rendering.doc");

        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        FontSettings.getDefaultInstance().setFontsFolder("/Users/username/MyFonts/", false);

        doc.save(dataDir + "Rendering.SpecifyingAFontFolder_Out.pdf");
        //ExEnd:SpecifyAFontFolder
    }

    public static void specifyMultipleFontFolders() throws Exception {
        //ExStart:SpecifyMultipleFontFolders
        Document doc = new Document(dataDir + "Rendering.doc");

        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        FontSettings.getDefaultInstance().setFontsFolders(new String[]{"/Users/username/MyFonts/", "/Users/username/Documents/Fonts/"}, true);

        doc.save(dataDir + "Rendering.SpecifyMultipleFontFolders_Out.pdf");
        //ExEnd:SpecifyMultipleFontFolders
    }

    public static void specifyFontsFromBothSystemFontsFolderAndCustomFolder() throws Exception {
        //ExStart:SpecifyFontsFromBothSystemFontsFolderAndCustomFolder
        Document doc = new Document(dataDir + "Rendering.doc");

        // Retrieve the array of environment-dependent font sources that are searched by default.
        // For example this will contain a "Windows\Fonts\" source on a Windows machines and /Library/Fonts/ on Mac OS X.
        // We add this array to a new ArrayList to make adding or removing font entries much easier.
        ArrayList<FolderFontSource> fontSources = new ArrayList(Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource("/Users/username/MyFonts/", true);

        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        // Convert the ArrayList of source back into a primitive array of FontSource objects.
        FontSourceBase[] updatedFontSources = fontSources.toArray(new FontSourceBase[fontSources.size()]);

        // Apply the new set of font sources to use.
        FontSettings.getDefaultInstance().setFontsSources(updatedFontSources);

        doc.save(dataDir + "Rendering.SpecifyFontsToBeReadFromBothSystemFontsFolderAndCustomFolder_Out.pdf");
        //ExEnd:SpecifyFontsFromBothSystemFontsFolderAndCustomFolder
    }

    //ExStart:loadingFontsStream
    public void streamFontSourceFileRendering() throws Exception {
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsSources(new FontSourceBase[]{new StreamFontSourceFile()});

        DocumentBuilder builder = new DocumentBuilder();
        builder.getDocument().setFontSettings(fontSettings);
        builder.getFont().setName("Kreon-Regular");
        builder.writeln("Test aspose text when saving to PDF.");

        builder.getDocument().save(dataDir + "FontSettings.StreamFontSourceFileRendering.pdf");
    }

    /// <summary>
    /// Load the font data only when required instead of storing it in the memory for the entire lifetime of the "FontSettings" object.
    /// </summary>
    private static class StreamFontSourceFile extends StreamFontSource  {
        public FileInputStream openFontDataStream() throws Exception {
            return new FileInputStream(dataDir + "Kreon-Regular.ttf");
        }
    }
    //ExEnd:loadingFontsStream
}
