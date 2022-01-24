package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.HtmlFixedSaveOptions;

@Test
public class WorkingWithHtmlFixedSaveOptions extends DocsExamplesBase {
    @Test
    public void useFontFromTargetMachine() throws Exception {
        //ExStart:UseFontFromTargetMachine
        Document doc = new Document(getMyDir() + "Bullet points with alternative font.docx");

        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setUseTargetMachineFonts(true);
        }

        doc.save(getArtifactsDir() + "WorkingWithHtmlFixedSaveOptions.UseFontFromTargetMachine.html", saveOptions);
        //ExEnd:UseFontFromTargetMachine
    }

    @Test
    public void writeAllCssRulesInSingleFile() throws Exception {
        //ExStart:WriteAllCssRulesInSingleFile
        Document doc = new Document(getMyDir() + "Document.docx");

        // Setting this property to true restores the old behavior (separate files) for compatibility with legacy code.
        // All CSS rules are written into single file "styles.css.
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        {
            saveOptions.setSaveFontFaceCssSeparately(false);
        }

        doc.save(getArtifactsDir() + "WorkingWithHtmlFixedSaveOptions.WriteAllCssRulesInSingleFile.html", saveOptions);
        //ExEnd:WriteAllCssRulesInSingleFile
    }
}
