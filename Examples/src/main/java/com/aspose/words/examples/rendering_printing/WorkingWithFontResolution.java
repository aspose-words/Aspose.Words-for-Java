package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.LoadOptions;
import com.aspose.words.TableSubstitutionRule;
import com.aspose.words.examples.Utils;

public class WorkingWithFontResolution {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String dataDir = Utils.getDataDir(WorkingWithFontResolution.class);
		FontSettingsWithLoadOptions(dataDir);
        SetFontsFolder(dataDir);
	}
	
	public static void FontSettingsWithLoadOptions(String dataDir) throws Exception
    {
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

    public static void SetFontsFolder(String dataDir) throws Exception
    {
        // ExStart:SetFontsFolder
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(dataDir + "Fonts", false);
        LoadOptions lo = new LoadOptions();
        lo.setFontSettings(fontSettings);
        Document doc = new Document(dataDir + "myfile.html", lo);
        // ExEnd:SetFontsFolder
    }

}
