package com.aspose.words.examples.rendering_printing;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.ArrayList;
import java.util.Arrays;

public class WorkingWithFontSources {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SetPdfSaveOptions.class);

        getListOfAvailableFonts(dataDir);
    }

    public static void getListOfAvailableFonts(String dataDir) throws Exception {
        // ExStart:GetListOfAvailableFonts
        // The path to the documents directory.
        Document doc = new Document(dataDir + "TestFile.docx");

        FontSettings fontSettings = new FontSettings();
        ArrayList<FolderFontSource> fontSources = new ArrayList(Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource(dataDir, true);

        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        // Convert the Arraylist of source back into a primitive array of FontSource objects.
        FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources.toArray(new FontSourceBase[fontSources.size()]);

        for (PhysicalFontInfo fontInfo : (Iterable<PhysicalFontInfo>) updatedFontSources[0].getAvailableFonts()) {
            System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
            System.out.println("FullFontName  : " + fontInfo.getFullFontName());
            System.out.println("Version  : " + fontInfo.getVersion());
            System.out.println("FilePath : " + fontInfo.getFilePath());
        }
        // ExEnd:GetListOfAvailableFonts
    }
}
