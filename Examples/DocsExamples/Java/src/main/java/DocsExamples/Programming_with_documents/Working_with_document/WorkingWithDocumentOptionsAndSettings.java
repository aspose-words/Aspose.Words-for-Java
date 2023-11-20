package DocsExamples.Programming_with_documents.Working_with_document;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;

@Test
public class WorkingWithDocumentOptionsAndSettings extends DocsExamplesBase
{
    @Test
    public void optimizeForMsWord() throws Exception
    {
        //ExStart:OptimizeForMsWord
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2016);

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.OptimizeForMsWord.docx");
        //ExEnd:OptimizeForMsWord
    }

    @Test
    public void showGrammaticalAndSpellingErrors() throws Exception
    {
        //ExStart:ShowGrammaticalAndSpellingErrors
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.setShowGrammaticalErrors(true);
        doc.setShowSpellingErrors(true);

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.ShowGrammaticalAndSpellingErrors.docx");
        //ExEnd:ShowGrammaticalAndSpellingErrors
    }

    @Test
    public void cleanupUnusedStylesAndLists() throws Exception
    {
        //ExStart:CleanupUnusedStylesandLists
        Document doc = new Document(getMyDir() + "Unused styles.docx");

        // Combined with the built-in styles, the document now has eight styles.
        // A custom style is marked as "used" while there is any text within the document
        // formatted in that style. This means that the 4 styles we added are currently unused.
        System.out.println(MessageFormat.format("Count of styles before Cleanup: {0}\n", doc.getStyles().getCount()) +
                              MessageFormat.format("Count of lists before Cleanup: {0}", doc.getLists().getCount()));

        // Cleans unused styles and lists from the document depending on given CleanupOptions. 
        CleanupOptions cleanupOptions = new CleanupOptions(); { cleanupOptions.setUnusedLists(false); cleanupOptions.setUnusedStyles(true); }
        doc.cleanup(cleanupOptions);

        System.out.println(MessageFormat.format("Count of styles after Cleanup was decreased: {0}\n", doc.getStyles().getCount()) +
                              MessageFormat.format("Count of lists after Cleanup is the same: {0}", doc.getLists().getCount()));

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.CleanupUnusedStylesAndLists.docx");
        //ExEnd:CleanupUnusedStylesandLists
    }

    @Test
    public void cleanupDuplicateStyle() throws Exception
    {
        //ExStart:CleanupDuplicateStyle
        Document doc = new Document(getMyDir() + "Document.docx");

        // Count of styles before Cleanup.
        System.out.println(doc.getStyles().getCount());

        // Cleans duplicate styles from the document.
        CleanupOptions options = new CleanupOptions(); { options.setDuplicateStyle(true); }
        doc.cleanup(options);

        // Count of styles after Cleanup was decreased.
        System.out.println(doc.getStyles().getCount());

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.CleanupDuplicateStyle.docx");
        //ExEnd:CleanupDuplicateStyle
    }

    @Test
    public void viewOptions() throws Exception
    {
        //ExStart:SetViewOption
        Document doc = new Document(getMyDir() + "Document.docx");
        
        doc.getViewOptions().setViewType(ViewType.PAGE_LAYOUT);
        doc.getViewOptions().setZoomPercent(50);

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.ViewOptions.docx");
        //ExEnd:SetViewOption
    }

    @Test
    public void documentPageSetup() throws Exception
    {
        //ExStart:DocumentPageSetup
        Document doc = new Document(getMyDir() + "Document.docx");

        // Set the layout mode for a section allowing to define the document grid behavior.
        // Note that the Document Grid tab becomes visible in the Page Setup dialog of MS Word
        // if any Asian language is defined as editing language.
        doc.getFirstSection().getPageSetup().setLayoutMode(SectionLayoutMode.GRID);
        doc.getFirstSection().getPageSetup().setCharactersPerLine(30);
        doc.getFirstSection().getPageSetup().setLinesPerPage(10);

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.DocumentPageSetup.docx");
        //ExEnd:DocumentPageSetup
    }

    @Test
    public void addJapaneseAsEditingLanguages() throws Exception
    {
        //ExStart:AddJapaneseAsEditinglanguages
        LoadOptions loadOptions = new LoadOptions();
        
        // Set language preferences that will be used when document is loading.
        loadOptions.getLanguagePreferences().addEditingLanguage(EditingLanguage.JAPANESE);
        //ExEnd:AddJapaneseAsEditinglanguages

        Document doc = new Document(getMyDir() + "No default editing language.docx", loadOptions);

        int localeIdFarEast = doc.getStyles().getDefaultFont().getLocaleIdFarEast();
        System.out.println(localeIdFarEast == (int) EditingLanguage.JAPANESE
                    ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
                    : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
    }

    @Test
    public void setRussianAsDefaultEditingLanguage() throws Exception
    {
        //ExStart:SetRussianAsDefaultEditingLanguage
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.getLanguagePreferences().setDefaultEditingLanguage(EditingLanguage.RUSSIAN);

        Document doc = new Document(getMyDir() + "No default editing language.docx", loadOptions);

        int localeId = doc.getStyles().getDefaultFont().getLocaleId();
        System.out.println(localeId == (int) EditingLanguage.RUSSIAN
                    ? "The document either has no any language set in defaults or it was set to Russian originally."
                    : "The document default language was set to another than Russian language originally, so it is not overridden.");
        //ExEnd:SetRussianAsDefaultEditingLanguage
    }

    @Test
    public void pageSetupAndSectionFormatting() throws Exception
    {
        //ExStart:PageSetupAndSectionFormatting
        //GistId:7c0668453e53ed7a57d3ea3a05520f21
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getPageSetup().setOrientation(Orientation.LANDSCAPE);
        builder.getPageSetup().setLeftMargin(50.0);
        builder.getPageSetup().setPaperSize(PaperSize.PAPER_10_X_14);

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.SetPageSetupAndSectionFormatting.docx");
        //ExEnd:PageSetupAndSectionFormatting
    }

    @Test
    public void pageBorderProperties() throws Exception
    {
        //ExStart:PageBorderProperties
        Document doc = new Document();

        PageSetup pageSetup = doc.getSections().get(0).getPageSetup();
        pageSetup.setBorderAlwaysInFront(false);
        pageSetup.setBorderDistanceFrom(PageBorderDistanceFrom.PAGE_EDGE);
        pageSetup.setBorderAppliesTo(PageBorderAppliesTo.FIRST_PAGE);

        Border border = pageSetup.getBorders().getByBorderType(BorderType.TOP);
        border.setLineStyle(LineStyle.SINGLE);
        border.setLineWidth(30.0);
        border.setColor(Color.BLUE);
        border.setDistanceFromText(0.0);

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.PageBorderProperties.docx");
        //ExEnd:PageBorderProperties
    }

    @Test
    public void lineGridSectionLayoutMode() throws Exception
    {
        //ExStart:LineGridSectionLayoutMode
        //GistId:7c0668453e53ed7a57d3ea3a05520f21
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable pitching, and then use it to set the number of lines per page in this section.
        // A large enough font size will push some lines down onto the next page to avoid overlapping characters.
        builder.getPageSetup().setLayoutMode(SectionLayoutMode.LINE_GRID);
        builder.getPageSetup().setLinesPerPage(15);

        builder.getParagraphFormat().setSnapToGrid(true);

        for (int i = 0; i < 30; i++)
            builder.write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ");

        doc.save(getArtifactsDir() + "WorkingWithDocumentOptionsAndSettings.LinesPerPage.docx");
        //ExEnd:LineGridSectionLayoutMode
    }
}
