package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Font;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@Test
public class WorkingWithFonts extends DocsExamplesBase
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
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.SHADOW));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.EFFECT_3_D));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.REFLECTION));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.OUTLINE));
        System.out.println(runFont.hasDmlEffect(TextDmlEffect.FILL));
        //ExEnd:CheckDMLTextEffect
    }

    @Test
    public void setFontFormatting() throws Exception
    {
        //ExStart:DocumentBuilderSetFontFormatting
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Font font = builder.getFont();
        font.setBold(true);
        font.setColor(Color.BLUE);
        font.setItalic(true);
        font.setName("Arial");
        font.setSize(24.0);
        font.setSpacing(5.0);
        font.setUnderline(Underline.DOUBLE);

        builder.writeln("I'm a very nice formatted string.");
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontFormatting.docx");
        //ExEnd:DocumentBuilderSetFontFormatting
    }

    @Test
    public void setFontEmphasisMark() throws Exception
    {
        //ExStart:SetFontEmphasisMark
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
    public void setFontsFolders() throws Exception
    {
        //ExStart:SetFontsFolders
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[]
        {
            new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true)
        });

        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFolders.pdf");
        //ExEnd:SetFontsFolders           
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
    public void setFontFallbackSettings() throws Exception
    {
        //ExStart:SetFontFallbackSettings
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().load(getMyDir() + "Font fallback rules.xml");
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontFallbackSettings.pdf");
        //ExEnd:SetFontFallbackSettings
    }

    @Test
    public void notoFallbackSettings() throws Exception
    {
        //ExStart:SetPredefinedFontFallbackSettings
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        fontSettings.getFallbackSettings().loadNotoFallbackSettings();
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.NotoFallbackSettings.pdf");
        //ExEnd:SetPredefinedFontFallbackSettings
    }

    @Test
    public void setFontsFoldersDefaultInstance() throws Exception
    {
        //ExStart:SetFontsFoldersDefaultInstance
        FontSettings.getDefaultInstance().setFontsFolder("C:\\MyFonts\\", true);
        //ExEnd:SetFontsFoldersDefaultInstance           

        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFoldersDefaultInstance.pdf");
    }

    @Test
    public void setFontsFoldersMultipleFolders() throws Exception
    {
        //ExStart:SetFontsFoldersMultipleFolders
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        FontSettings fontSettings = new FontSettings();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead.
        fontSettings.setFontsFolders(new String[] { "C:\\MyFonts\\", "D:\\Misc\\Fonts\\" }, true);
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFoldersMultipleFolders.pdf");
        //ExEnd:SetFontsFoldersMultipleFolders           
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
        List<FontSourceBase> fontSources = new ArrayList<>(Arrays.asList(fontSettings.getFontsSources()));

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);

        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        FontSourceBase[] updatedFontSources = fontSources.toArray(new FontSourceBase[0]);
        fontSettings.setFontsSources(updatedFontSources);
        
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFoldersSystemAndCustomFolder.pdf");
        //ExEnd:SetFontsFoldersSystemAndCustomFolder
    }

    @Test
    public void setFontsFoldersWithPriority() throws Exception
    {
        //ExStart:SetFontsFoldersWithPriority
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[]
        {
            new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true,1)
        });
        //ExEnd:SetFontsFoldersWithPriority           

        Document doc = new Document(getMyDir() + "Rendering.docx");
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFoldersWithPriority.pdf");
    }

    @Test
    public void setTrueTypeFontsFolder() throws Exception
    {
        //ExStart:SetTrueTypeFontsFolder
        Document doc = new Document(getMyDir() + "Rendering.docx");

        FontSettings fontSettings = new FontSettings();
        // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
        // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
        // FontSettings.SetFontSources instead
        fontSettings.setFontsFolder("C:\\MyFonts\\", false);
        // Set font settings
        doc.setFontSettings(fontSettings);
        
        doc.save(getArtifactsDir() + "WorkingWithFonts.SetTrueTypeFontsFolder.pdf");
        //ExEnd:SetTrueTypeFontsFolder
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
    public void fontSettingsWithLoadOption() throws Exception
    {
        //ExStart:FontSettingsWithLoadOption
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(new FontSettings());

        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);
        //ExEnd:FontSettingsWithLoadOption   
    }

    @Test
    public void fontSettingsDefaultInstance() throws Exception
    {
        //ExStart:FontSettingsFontSource
        //ExStart:FontSettingsDefaultInstance
        FontSettings fontSettings = FontSettings.getDefaultInstance();
        //ExEnd:FontSettingsDefaultInstance   
        fontSettings.setFontsSources(new FontSourceBase[]
        {
            new SystemFontSource(),
            new FolderFontSource("C:\\MyFonts\\", true)
        });
        //ExEnd:FontSettingsFontSource

        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setFontSettings(fontSettings);
        Document doc = new Document(getMyDir() + "Rendering.docx", loadOptions);
    }

    @Test
    public void getListOfAvailableFonts()
    {
        //ExStart:GetListOfAvailableFonts
        List<FontSourceBase> fontSources = new ArrayList<>(Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

        // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        FolderFontSource folderFontSource = new FolderFontSource(getMyDir(), true);
        // Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.add(folderFontSource);

        FontSourceBase[] updatedFontSources = fontSources.toArray(new FontSourceBase[0]);

        for (PhysicalFontInfo fontInfo : updatedFontSources[0].getAvailableFonts())
        {
            System.out.println("FontFamilyName : " + fontInfo.getFontFamilyName());
            System.out.println("FullFontName  : " + fontInfo.getFullFontName());
            System.out.println("Version  : " + fontInfo.getVersion());
            System.out.println("FilePath : " + fontInfo.getFilePath());
        }
        //ExEnd:GetListOfAvailableFonts
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

    //ExStart:ResourceSteamFontSourceExample
    @Test
    public void resourceSteamFontSourceExample() throws Exception
    {
        Document doc = new Document(getMyDir() + "Rendering.docx");
        
        FontSettings.getDefaultInstance().setFontsSources(new FontSourceBase[]
            { new SystemFontSource(), new ResourceSteamFontSource() });

        doc.save(getArtifactsDir() + "WorkingWithFonts.SetFontsFolders.pdf");
    }

    static class ResourceSteamFontSource extends StreamFontSource
    {
        public InputStream openFontDataStream() throws IOException {
            return getClass().getClassLoader().getResource("resourceName").openStream();
        }
    }
    //ExEnd:ResourceSteamFontSourceExample

    //ExStart:GetSubstitutionWithoutSuffixes
    @Test
    public void getSubstitutionWithoutSuffixes() throws Exception
    {
        Document doc = new Document(getMyDir() + "Get substitution without suffixes.docx");

        DocumentSubstitutionWarnings substitutionWarningHandler = new DocumentSubstitutionWarnings();
        doc.setWarningCallback(substitutionWarningHandler);

        List<FontSourceBase> fontSources = new ArrayList<>(Arrays.asList(FontSettings.getDefaultInstance().getFontsSources()));

        FolderFontSource folderFontSource = new FolderFontSource(getFontsDir(), true);
        fontSources.add(folderFontSource);

        FontSourceBase[] updatedFontSources = fontSources.toArray(new FontSourceBase[0]);
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
