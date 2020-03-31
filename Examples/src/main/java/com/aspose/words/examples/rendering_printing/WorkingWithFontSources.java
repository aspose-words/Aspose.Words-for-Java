package com.aspose.words.examples.rendering_printing;

import java.util.ArrayList;
import java.util.Arrays;

import com.aspose.words.Document;
import com.aspose.words.FolderFontSource;
import com.aspose.words.FontSettings;
import com.aspose.words.FontSourceBase;
import com.aspose.words.LoadOptions;
import com.aspose.words.PhysicalFontInfo;
import com.aspose.words.SystemFontSource;
import com.aspose.words.examples.Utils;

public class WorkingWithFontSources {

	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(WorkingWithFontSources.class);

		SetFontsFolder(dataDir);
		SetTrueTypeFontsFolder(dataDir);
		SetMultipleFontsFolder(dataDir);
		SetFontsFolderWithPriority(dataDir);
		GetAllAvailableFonts(dataDir);
	}

	public static void SetFontsFolder(String dataDir) throws Exception {
		// ExStart: SetFontsFolder
		FontSettings.getDefaultInstance().setFontsSources(
				new FontSourceBase[] { new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true) });

		Document doc = new Document(dataDir + "Rendering.doc");
		dataDir = dataDir + "Rendering.SetFontsFolders_out.pdf";
		doc.save(dataDir);
		// ExEnd: SetFontsFolder
	}

	public static void SetTrueTypeFontsFolder(String dataDir) throws Exception {
		// ExStart: SetTrueTypeFontsFolder
		// For complete examples and data files, please go to
		// https://github.com/aspose-words/Aspose.Words-for-Java
		Document doc = new Document(dataDir + "Rendering.doc");

		FontSettings FontSettings = new FontSettings();

		// Note that this setting will override any default font sources that are being
		// searched by default. Now only these folders will be searched for
		// Fonts when rendering or embedding fonts. To add an extra font source while
		// keeping system font sources then use both FontSettings.GetFontSources and
		// FontSettings.SetFontSources instead.
		FontSettings.setFontsFolder("C:\\MyFonts\\", false);

		// Set font settings
		doc.setFontSettings(FontSettings);
		doc.save(dataDir + "Rendering.SetFontsFolder_out.pdf");
		// ExEnd: SetTrueTypeFontsFolder
	}

	public static void SetMultipleFontsFolder(String dataDir) throws Exception {
		// ExStart: SetMultipleFontsFolder
		// For complete examples and data files, please go to
		// https://github.com/aspose-words/Aspose.Words-for-.NET
		// Open a Document
		Document doc = new Document(dataDir + "Rendering.doc");
		FontSettings FontSettings = new FontSettings();

		// Note that this setting will override any default font sources that are being
		// searched by default. Now only these folders will be searched for
		// Fonts when rendering or embedding fonts. To add an extra font source while
		// keeping system font sources then use both FontSettings.GetFontSources and
		// FontSettings.SetFontSources instead.
		FontSettings.setFontsFolders(new String[] { "C:\\MyFonts\\", "D:\\Misc\\Fonts\\" }, true);

		// Set font settings
		doc.setFontSettings(FontSettings);
		doc.save(dataDir + "Rendering.SetFontsFolders_out.pdf");
		// ExEnd: SetMultipleFontsFolder
	}

	public static void SetFontsFolderWithPriority(String dataDir) throws Exception {
		// ExStart: SetFontsFolderWithPriority
		FontSettings.getDefaultInstance().setFontsSources(
				new FontSourceBase[] { new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true, 1) });
		// ExEnd: SetFontsFolderWithPriority
		Document doc = new Document(dataDir + "Rendering.doc");
		doc.save(dataDir + "Rendering.SetFontsFolder_out.pdf");
	}

	public static void GetAllAvailableFonts(String dataDir) throws Exception {
		// ExStart: GetAllAvailableFonts
		// Get available system fonts
		for (PhysicalFontInfo fontInfo : (Iterable<PhysicalFontInfo>) new SystemFontSource().getAvailableFonts()) {
			System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
			System.out.println("FullFontName  : " + fontInfo.getFullFontName());
			System.out.println("Version  : " + fontInfo.getVersion());
			System.out.println("FilePath : " + fontInfo.getFilePath());
		}

		// Get available fonts in folder
		for (PhysicalFontInfo fontInfo : (Iterable<PhysicalFontInfo>) new FolderFontSource(dataDir, true)
				.getAvailableFonts()) {
			System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
			System.out.println("FullFontName  : " + fontInfo.getFullFontName());
			System.out.println("Version  : " + fontInfo.getVersion());
			System.out.println("FilePath : " + fontInfo.getFilePath());
		}

		// Get available fonts from FontSettings
		ArrayList<FolderFontSource> fontSources = new ArrayList(
				Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

		// Convert the Arraylist of source back into a primitive array of FontSource
		// objects.
		FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources
				.toArray(new FontSourceBase[fontSources.size()]);

		for (PhysicalFontInfo fontInfo : (Iterable<PhysicalFontInfo>) updatedFontSources[0].getAvailableFonts()) {
			System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
			System.out.println("FullFontName  : " + fontInfo.getFullFontName());
			System.out.println("Version  : " + fontInfo.getVersion());
			System.out.println("FilePath : " + fontInfo.getFilePath());
		}
		// ExEnd: GetAllAvailableFonts
	}
}
