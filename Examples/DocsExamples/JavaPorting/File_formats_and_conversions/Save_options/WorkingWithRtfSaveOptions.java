package DocsExamples.File_Formats_and_Conversions.Save_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.RtfSaveOptions;


public class WorkingWithRtfSaveOptions extends DocsExamplesBase
{
    @Test
    public void savingImagesAsWmf() throws Exception
    {
        //ExStart:SavingImagesAsWmf
        //GistId:6f849e51240635a6322ab0460938c922
        Document doc = new Document(getMyDir() + "Document.docx");

        RtfSaveOptions saveOptions = new RtfSaveOptions(); { saveOptions.setSaveImagesAsWmf(true); }

        doc.save(getArtifactsDir() + "WorkingWithRtfSaveOptions.SavingImagesAsWmf.rtf", saveOptions);
        //ExEnd:SavingImagesAsWmf
    }
}
