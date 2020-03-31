package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.FolderFontSource;
import com.aspose.words.FontSettings;
import com.aspose.words.FontSourceBase;
import com.aspose.words.LoadOptions;
import com.aspose.words.SystemFontSource;
import com.aspose.words.examples.Utils;

public class SetFontSettings {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(SetFontSettings.class) + "RenderingAndPrinting/";

        SetFontsFolder(dataDir);
        enableDisableFontSubstitution(dataDir);
        setFontFallbackSettings(dataDir);
        setPredefinedFontFallbackSettings(dataDir);
    }

    public static void SetFontsFolder(String dataDir) throws Exception {
		// ExStart:SetFontsFolder
    	FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[] {
    			    new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true)
    	});
    	
    	Document doc = new Document(dataDir + "Rendering.doc");
    	dataDir = dataDir + "Rendering.SetFontsFolders_out.pdf";
    	doc.save(dataDir);
		// ExEnd:SetFontsFolder
	}
    
    public static void enableDisableFontSubstitution(String dataDir) throws Exception {
        // ExStart:EnableDisableFontSubstitution
        // The path to the documents directory.
        Document doc = new Document(dataDir + "Rendering.doc");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(false);

        // Set font settings
        doc.setFontSettings(fontSettings);
        dataDir = dataDir + "Rendering.DisableFontSubstitution_out.pdf";
        doc.save(dataDir);
        // ExEnd:EnableDisableFontSubstitution
        System.out.println("\nDocument is rendered to PDF with disabled font substitution.\nFile saved at " + dataDir);
    }

    public static void setFontFallbackSettings(String dataDir) throws Exception {
        // ExStart:SetFontFallbackSettings
        Document doc = new Document(dataDir + "Rendering.doc");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().load(dataDir + "Fallback.xml");

        // Set font settings
        doc.setFontSettings(fontSettings);
        dataDir = dataDir + "Rendering.FontFallback_out.pdf";
        doc.save(dataDir);
        // ExEnd:SetFontFallbackSettings
        System.out.println("\nDocument is rendered to PDF with font fallback.\nFile saved at " + dataDir);
    }

    public static void setPredefinedFontFallbackSettings(String dataDir) throws Exception {
        // ExStart: setPredefinedFontFallbackSettings
        Document doc = new Document(dataDir + "Rendering.doc");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().loadNotoFallbackSettings();

        // Set font settings
        doc.setFontSettings(fontSettings);
        dataDir = dataDir + "Rendering.FontFallbackGoogleNoto_out.pdf";
        doc.save(dataDir);
        // ExEnd: setPredefinedFontFallbackSettings
        System.out.println("\nDocument is rendered to PDF with font fallback.\nFile saved at " + dataDir);
    }
}
