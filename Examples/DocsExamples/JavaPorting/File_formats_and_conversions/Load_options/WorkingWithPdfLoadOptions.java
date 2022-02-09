package DocsExamples.File_Formats_and_Conversions.Load_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.PdfSaveOptions;
import com.aspose.words.PdfEncryptionDetails;
import com.aspose.words.PdfEncryptionAlgorithm;
import com.aspose.words.PdfLoadOptions;
import com.aspose.words.LoadFormat;


public class WorkingWithPdfLoadOptions extends DocsExamplesBase
{
    @Test
    public void loadEncryptedPdf() throws Exception
    {
        //ExStart:LoadEncryptedPdf  
        Document doc = new Document(getMyDir() + "Pdf Document.pdf");

        PdfSaveOptions saveOptions = new PdfSaveOptions();
        {
            saveOptions.setEncryptionDetails(new PdfEncryptionDetails("Aspose", null, PdfEncryptionAlgorithm.RC_4_40));
        }

        doc.save(getArtifactsDir() + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", saveOptions);

        PdfLoadOptions loadOptions = new PdfLoadOptions(); { loadOptions.setPassword("Aspose"); loadOptions.setLoadFormat(LoadFormat.PDF); }

        doc = new Document(getArtifactsDir() + "WorkingWithPdfLoadOptions.LoadEncryptedPdf.pdf", loadOptions);
        //ExEnd:LoadEncryptedPdf
    }

    @Test
    public void loadPageRangeOfPdf() throws Exception
    {
        //ExStart:LoadPageRangeOfPdf  
        PdfLoadOptions loadOptions = new PdfLoadOptions(); { loadOptions.setPageIndex(0); loadOptions.setPageCount(1); }

        //ExStart:LoadPDF
        Document doc = new Document(getMyDir() + "Pdf Document.pdf", loadOptions);

        doc.save(getArtifactsDir() + "WorkingWithPdfLoadOptions.LoadPageRangeOfPdf.pdf");
        //ExEnd:LoadPDF
        //ExEnd:LoadPageRangeOfPdf
    }
}

