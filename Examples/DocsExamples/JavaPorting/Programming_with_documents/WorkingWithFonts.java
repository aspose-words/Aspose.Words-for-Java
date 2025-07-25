package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import java.awt.Color;
import com.aspose.words.Underline;
import com.aspose.ms.System.msConsole;
import com.aspose.words.RunCollection;
import com.aspose.words.TextDmlEffect;
import com.aspose.ms.System.Drawing.msColor;
import com.aspose.words.EmphasisMark;
import com.aspose.words.FontSettings;
import java.util.ArrayList;
import com.aspose.words.FontSourceBase;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.words.FolderFontSource;
import com.aspose.words.SystemFontSource;
import com.aspose.words.TableSubstitutionRule;
import com.aspose.words.LoadOptions;
import com.aspose.words.PhysicalFontInfo;
import com.aspose.words.IWarningCallback;
import com.aspose.words.WarningInfo;
import com.aspose.words.WarningType;
import com.aspose.words.StreamFontSource;
import com.aspose.ms.System.IO.Stream;
import org.testng.Assert;
import com.aspose.words.WarningInfoCollection;


class WorkingWithFonts extends DocsExamplesBase
{
    @Test
    public void fontFormatting() throws Exception
    {
        //ExStart:WriteAndFont
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Font font = builder.getFont();
        font.setSize(16.0);
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setName("Arial");
        font.setUnderline(Underline.DASH);

        builder.write("Sample text.");
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.FontFormatting.docx");
        //ExEnd:WriteAndFont
    }

    @Test
    public void getFontLineSpacing() throws Exception
    {
        //ExStart:GetFontLineSpacing
        //GistId:7cb86f131b74afcbebc153f0039e3947
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.getFont().setName("Calibri");
        builder.writeln("qText");

        Font font = builder.getDocument().getFirstSection().getBody().getFirstParagraph().getRuns().get(0).getFont();
        System.out.println("lineSpacing = {font.LineSpacing}");
        //ExEnd:GetFontLineSpacing
    }

    @Test
    public void checkDMLTextEffect() throws Exception
    {
        //ExStart:CheckDMLTextEffect
        Document doc = new Document(getMyDir() + "DrawingML text effects.docx");
        
        RunCollection runs = doc.getFirstSection().getBody().getFirstParagraph().getRuns();
        Font runFont = runs.get(0).getFont();

        // One run might have several Dml text effects applied.
        msConsole.writeLine(runFont.hasDmlEffect(TextDmlEffect.SHADOW));
        msConsole.writeLine(runFont.hasDmlEffect(TextDmlEffect.EFFECT_3_D));
        msConsole.writeLine(runFont.hasDmlEffect(TextDmlEffect.REFLECTION));
        msConsole.writeLine(runFont.hasDmlEffect(TextDmlEffect.OUTLINE));
        msConsole.writeLine(runFont.hasDmlEffect(TextDmlEffect.FILL));
        //ExEnd:CheckDMLTextEffect
    }

    @Test
    public void setFontFormatting() throws Exception
    {
        //ExStart:SetFontFormatting
        //GistId:7cb86f131b74afcbebc153f0039e3947
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Font font = builder.getFont();
        font.setBold(true);
        font.setColor(msColor.getDarkBlue());
        font.setItalic(true);
        font.setName("Arial");
        font.setSize(24.0);
        font.setSpacing(5.0);
        font.setUnderline(Underline.DOUBLE);

        builder.writeln("I'm a very nice formatted string.");
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontFormatting.docx");
        //ExEnd:SetFontFormatting
    }

    @Test
    public void setFontEmphasisMark() throws Exception
    {
        //ExStart:SetFontEmphasisMark
        //GistId:7cb86f131b74afcbebc153f0039e3947
        Document document = new Document();
        DocumentBuilder builder = new DocumentBuilder(document);

        builder.getFont().setEmphasisMark(EmphasisMark.UNDER_SOLID_CIRCLE);

        builder.write("Emphasis text");
        builder.writeln();
        builder.getFont().clearFormatting();
        builder.write("Simple text");

        document.save(getArtifactsDir() + "WorkingWithFonts.SetFontEmphasisMark.docx");
        //ExEnd:SetFontEmphasisMark
    }

    @Test
    public void enableDisableFontSubstitution() throws Exception
    {
        //ExStart:EnableDisableFontSubstitution
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        fontSettings.getSubstitutionSettings().getFontInfoSubstitution().setEnabled(false);
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.EnableDisableFontSubstitution.pdf");
        //ExEnd:EnableDisableFontSubstitution
    }

    @Test
    public void fontFallbackSettings() throws Exception
    {
        //ExStart:FontFallbackSettings
        //GistId:a08698f540d47082b4e2dbb1cb67fc1b
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().load(getMyDir() + "Font fallback rules.xml");
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.FontFallbackSettings.pdf");
        //ExEnd:FontFallbackSettings
    }

    @Test
    public void notoFallbackSettings() throws Exception
    {
        //ExStart:NotoFallbackSettings
        //GistId:a08698f540d47082b4e2dbb1cb67fc1b
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().loadNotoFallbackSettings();
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.NotoFallbackSettings.pdf");
        //ExEnd:NotoFallbackSettings
    }

    @Test
    public void defaultInstance() throws Exception
    {
        //ExStart:DefaultInstance
        //GistId:7e64f6d40825be58a8c12f1307c12964
        FontSettings.getDefaultInstance().setFontsFolder("C:\\MyFonts\\", true);
        //ExEnd:DefaultInstance

        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.save(getArtifactsDir() + "WorkingWithFonts.DefaultInstance.pdf");
    }

    @Test
    public void multipleFolders() throws Exception
    {
        //ExStart:MultipleFolders
        //GistId:7e64f6d40825be58a8c12f1307c12964
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        FontSettings fontSettings = new FontSettings();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        fontSettings.setFontsFolders(new String[] { "C:\\MyFonts\\", "D:\\Misc\\Fonts\\" }, true);
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.MultipleFolders.pdf");
        //ExEnd:MultipleFolders
    }

    @Test
    public void setFontsFoldersSystemAndCustomFolder() throws Exception
    {
        //ExStart:SetFontsFoldersSystemAndCustomFolder
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        // Retrieve the array of environment-dependent font sources that are searched by default.
        // For example this will contain a "Windows\Fonts\" source on a Windows machines.
        // We add this array to a new List to make adding or removing font entries much easier.
        ArrayList<FontSourceBase> fontSources = msArrayList.ctor(fontSettings.getFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);
        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        FontSourceBase[] updatedFontSources = msArrayList.toArray(fontSources, new FontSourceBase[0]);
        fontSettings.setFontsSources(updatedFontSources);

        doc.setFontSettings(fontSettings);

        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFoldersSystemAndCustomFolder.pdf");
        //ExEnd:SetFontsFoldersSystemAndCustomFolder
    }

    @Test
    public void fontsFoldersWithPriority() throws Exception
    {
        //ExStart:FontsFoldersWithPriority
        //GistId:7e64f6d40825be58a8c12f1307c12964
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[]
        {
            new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true, 1)
        });
        //ExEnd:FontsFoldersWithPriority

        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.save(getArtifactsDir() + "WorkingWithFonts.FontsFoldersWithPriority.pdf");
    }

    @Test
    public void trueTypeFontsFolder() throws Exception
    {
        //ExStart:TrueTypeFontsFolder
        //GistId:7e64f6d40825be58a8c12f1307c12964
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead
        fontSettings.setFontsFolder("C:\\MyFonts\\", false);
        // Set font settings
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.TrueTypeFontsFolder.pdf");
        //ExEnd:TrueTypeFontsFolder
    }

    @Test
    public void specifyDefaultFontWhenRendering() throws Exception
    {
        //ExStart:SpecifyDefaultFontWhenRendering
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        // If the default font defined here cannot be found during rendering then
        // the closest font on the machine is used instead.
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial Unicode MS");
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SpecifyDefaultFontWhenRendering.pdf");
        //ExEnd:SpecifyDefaultFontWhenRendering
    }

    @Test
    public void fontSettingsWithLoadOptions() throws Exception
    {
        //ExStart:FontSettingsWithLoadOptions
        FontSettings fontSettings = new FontSettings();

        TableSubstitutionRule substitutionRule = fontSettings.getSubstitutionSettings().getTableSubstitution();
        // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS"
        substitutionRule.addSubstitutes("UnknownFont1", new String[] { "Comic Sans MS" });
        
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);
        
        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);
        //ExEnd:FontSettingsWithLoadOptions
    }

    @Test
    public void setFontsFolder() throws Exception
    {
        //ExStart:SetFontsFolder
        FontSettings fontSettings = new FontSettings();
        fontSettings.setFontsFolder(getMyDir() + "Fonts", false);
        
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);
        
        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);
        //ExEnd:SetFontsFolder
    }

    @Test
    public void loadOptionFontSettings() throws Exception
    {
        //ExStart:LoadOptionFontSettings
        //GistId:a08698f540d47082b4e2dbb1cb67fc1b
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(new FontSettings());

        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);
        //ExEnd:LoadOptionFontSettings
    }

    @Test
    public void fontSettingsDefaultInstance() throws Exception
    {
        //ExStart:FontsFolders
        //GistId:7e64f6d40825be58a8c12f1307c12964
        //ExStart:FontSettingsFontSource
        //GistId:a08698f540d47082b4e2dbb1cb67fc1b
        //ExStart:FontSettingsDefaultInstance
        //GistId:a08698f540d47082b4e2dbb1cb67fc1b
        FontSettings fontSettings = FontSettings.getDefaultInstance();
        //ExEnd:FontSettingsDefaultInstance
        fontSettings.setFontsSources(new FontSourceBase[]
        {
            new SystemFontSource(),
            new FolderFontSource("C:\\MyFonts\\", true)
        });
        //ExEnd:FontSettingsFontSource

        Document doc = new Document(getMyDir() + "Rendering.docx");
        //ExEnd:FontsFolders
    }

    @Test
    public void availableFonts()
    {
        //ExStart:AvailableFonts
        //GistId:7e64f6d40825be58a8c12f1307c12964
        FontSettings fontSettings = new FontSettings();
        ArrayList<FontSourceBase> fontSources = msArrayList.ctor(fontSettings.getFontsSources());

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource(getMyDir(), true);
        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        FontSourceBase[] updatedFontSources = msArrayList.toArray(fontSources, new FontSourceBase[0]);

        for (PhysicalFontInfo fontInfo : updatedFontSources[0].getAvailableFonts())
        {
            System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
            System.out.println("FullFontName  : " + fontInfo.getFullFontName());
            System.out.println("Version  : " + fontInfo.getVersion());
            System.out.println("FilePath : " + fontInfo.getFilePath());
        }
        //ExEnd:AvailableFonts
    }

    @Test
    public void receiveNotificationsOfFonts() throws Exception
    {
        //ExStart:ReceiveNotificationsOfFonts
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();

        // We can choose the default font to use in the case of any missing fonts.
        fontSettings.getSubstitutionSettings().getDefaultFontSubstitution().setDefaultFontName("Arial");
        // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
        // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
        fontSettings.setFontsFolder("", false);

        // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
        HandleDocumentWarnings callback = new HandleDocumentWarnings();

        doc.setWarningCallback(callback);
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.ReceiveNotificationsOfFonts.pdf");
        //ExEnd:ReceiveNotificationsOfFonts
    }

    @Test
    public void receiveWarningNotification() throws Exception
    {
        //ExStart:ReceiveWarningNotification
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
        // are stored until the document save and then sent to the appropriate WarningCallback.
        doc.updatePageLayout();

        HandleDocumentWarnings callback = new HandleDocumentWarnings();
        doc.setWarningCallback(callback);
        
        // Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc.save(getArtifactsDir() + "WorkingWithFonts.ReceiveWarningNotification.pdf");
        //ExEnd:ReceiveWarningNotification  
    }

    //ExStart:HandleDocumentWarnings
    public static class HandleDocumentWarnings implements IWarningCallback
    {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        /// Potential issue during document procssing. The callback can be set to listen for warnings generated
        /// during document load and/or document save.
        /// </summary>
        public void warning(WarningInfo info)
        {
            // We are only interested in fonts being substituted.
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
            {
                System.out.println("Font substitution: " + info.getDescription());
            }
        }
    }
    //ExEnd:HandleDocumentWarnings

    @Test
    //ExStart:ResourceSteam
    //GistId:7e64f6d40825be58a8c12f1307c12964
    public void resourceSteam() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[]
            { new SystemFontSource(), new ResourceSteamFontSource() });

        doc.save(getArtifactsDir() + "WorkingWithFonts.ResourceSteam.pdf");
    }

    static class ResourceSteamFontSource extends StreamFontSource
    {
        public /*override*/ Stream openFontDataStream()
        {
            return Assembly.GetExecutingAssembly().GetManifestResourceStream("resourceName");
        }
    }
    //ExEnd:ResourceSteam

    @Test
    //ExStart:GetSubstitutionWithoutSuffixes
    //GistId:a08698f540d47082b4e2dbb1cb67fc1b
    public void getSubstitutionWithoutSuffixes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Get substitution without suffixes.docx");

        DocumentSubstitutionWarnings substitutionWarningHandler = new DocumentSubstitutionWarnings();
        doc.setWarningCallback(substitutionWarningHandler);

        ArrayList<FontSourceBase> fontSources = msArrayList.ctor(FontSettings.getDefaultInstance().getFontsSources());

        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), true);
        fontSources.add(folderFontSource);

        FontSourceBase[] updatedFontSources = msArrayList.toArray(fontSources, new FontSourceBase[0]);
        FontSettings.getDefaultInstance().setFontsSources(updatedFontSources);

        doc.save(getArtifactsDir() + "WorkingWithFonts.GetSubstitutionWithoutSuffixes.pdf");

        Assert.assertEquals(
            "Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
            substitutionWarningHandler.FontWarnings.get(0).getDescription());
    }

    public static class DocumentSubstitutionWarnings implements IWarningCallback
    {
        /// <summary>
        /// Our callback only needs to implement the "Warning" method.
        /// This method is called whenever there is a potential issue during document processing.
        /// The callback can be set to listen for warnings generated during document load and/or document save.
        /// </summary>
        public void warning(WarningInfo info)
        {
            // We are only interested in fonts being substituted.
            if (info.getWarningType() == WarningType.FONT_SUBSTITUTION)
                FontWarnings.warning(info);
        }

        public WarningInfoCollection FontWarnings = new WarningInfoCollection();
    }
    //ExEnd:GetSubstitutionWithoutSuffixes
}
