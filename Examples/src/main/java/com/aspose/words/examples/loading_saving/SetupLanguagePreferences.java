package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.EditingLanguage;
import com.aspose.words.LoadOptions;
import com.aspose.words.examples.Utils;

public class SetupLanguagePreferences {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(SetupLanguagePreferences.class);

        addJapaneseAsEditinglanguages(dataDir);
        setRussianAsDefaultEditingLanguage(dataDir);
    }

    private static void addJapaneseAsEditinglanguages(String dataDir) throws Exception {
        // ExStart:AddJapaneseAsEditinglanguages
        // Specify LoadOptions to add Editing Language
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().addEditingLanguage(EditingLanguage.JAPANESE);

        Document doc = new Document(dataDir + "languagepreferences.docx", loadOptions);

        int localeIdFarEast = doc.getStyles().getDefaultFont().getLocaleIdFarEast();
        if (localeIdFarEast == (int) EditingLanguage.JAPANESE)
            System.out.println("The document either has no any FarEast language set in defaults or it was set to Japanese originally.");
        else
            System.out.println("The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
        // ExEnd:AddJapaneseAsEditinglanguages
    }

    private static void setRussianAsDefaultEditingLanguage(String dataDir) throws Exception {
        // ExStart:SetRussianAsDefaultEditingLanguage
    	// Specify LoadOptions to set Default Editing Language
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(dataDir + "languagepreferences.docx", loadOptions);

        int localeId = doc.getStyles().getDefaultFont().getLocaleId();
        if (localeId == (int) EditingLanguage.RUSSIAN)
            System.out.println("The document either has no any language set in defaults or it was set to Russian originally.");
        else
            System.out.println("The document default language was set to another than Russian language originally, so it is not overridden.");
        // ExEnd:SetRussianAsDefaultEditingLanguage
    }
}