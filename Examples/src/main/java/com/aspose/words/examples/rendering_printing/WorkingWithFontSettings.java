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
import com.aspose.words.TableSubstitutionRule;
import com.aspose.words.examples.Utils;

public class WorkingWithFontSettings {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String dataDir = Utils.getDataDir(WorkingWithFontSettings.class);
		FontSettingsWithLoadOptions(dataDir);
		FontSettingsDefaultInstance(dataDir);
		SetFontsFolder(dataDir);
		GetListOfAvailableFonts(dataDir);
	}

	public static void FontSettingsWithLoadOptions(String dataDir) throws Exception {
		// ExStart:FontSettingsWithLoadOptions
		FontSettings fontSettings = new FontSettings();
		TableSubstitutionRule substitutionRule = fontSettings.getSubstitutionSettings().getTableSubstitution();
		// If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS".
		substitutionRule.addSubstitutes("UnknownFont1", new String[] { "Comic Sans MS" });
		LoadOptions lo = new LoadOptions();
		lo.setFontSettings(fontSettings);
		Document doc = new Document(dataDir + "myfile.html", lo);
		// ExEnd:FontSettingsWithLoadOptions
		System.out.println("\nFile created successfully.\nFile saved at " + dataDir);
	}

	public static void FontSettingsDefaultInstance(String dataDir) throws Exception {
		// ExStart:FontSettingsFontSource
		FontSettings fontSettings = FontSettings.getDefaultInstance();
		fontSettings.setFontsSources(
				new FontSourceBase[] { new SystemFontSource(), new FolderFontSource("/home/user/MyFonts", true) });
		// ExEnd:FontSettingsFontSource

		// init font settings
		LoadOptions loadOptions = new LoadOptions();
		loadOptions.setFontSettings(fontSettings);
		Document doc1 = new Document(dataDir + "MyDocument.docx", loadOptions);

		LoadOptions loadOptions2 = new LoadOptions();
		loadOptions2.setFontSettings(fontSettings);
		Document doc2 = new Document(dataDir + "MyDocument.docx", loadOptions2);
	}

	public static void SetFontsFolder(String dataDir) throws Exception {
		// ExStart:SetFontsFolder
		FontSettings fontSettings = new FontSettings();
		fontSettings.setFontsFolder(dataDir + "Fonts", false);
		LoadOptions lo = new LoadOptions();
		lo.setFontSettings(fontSettings);
		Document doc = new Document(dataDir + "myfile.html", lo);
		// ExEnd:SetFontsFolder
	}

	public static void GetListOfAvailableFonts(String dataDir) throws Exception {
		// ExStart:GetListOfAvailableFonts
		// The path to the documents directory.
		Document doc = new Document(dataDir + "TestFile.docx");

		FontSettings fontSettings = new FontSettings();
		ArrayList<FolderFontSource> fontSources = new ArrayList(
				Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

		// Add a new folder source which will instruct Aspose.Words to search the
		// following folder for fonts.
		FolderFontSource folderFontSource = new FolderFontSource(dataDir, true);

		// Add the custom folder which contains our fonts to the list of existing font sources.
		fontSources.add(folderFontSource);

		// Convert the Arraylist of source back into a primitive array of FontSource objects.
		FontSourceBase[] updatedFontSources = (FontSourceBase[]) fontSources
				.toArray(new FontSourceBase[fontSources.size()]);

		for (PhysicalFontInfo fontInfo : (Iterable<PhysicalFontInfo>) updatedFontSources[0].getAvailableFonts()) {
			System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
			System.out.println("FullFontName  : " + fontInfo.getFullFontName());
			System.out.println("Version  : " + fontInfo.getVersion());
			System.out.println("FilePath : " + fontInfo.getFilePath());
		}
		// ExEnd:GetListOfAvailableFonts
	}

}
