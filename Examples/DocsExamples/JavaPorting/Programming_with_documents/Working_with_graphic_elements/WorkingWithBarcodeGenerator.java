package DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;


class WorkingWithBarcodeGenerator extends DocsExamplesBase
{
    @Test
    public void barcodeGenerator() throws Exception
    {
        //ExStart:BarcodeGenerator
        //GistId:00d34dba66626dbc0175b60bb3b71c8a
        Document doc = new Document(getMyDir() + "Field sample - BARCODE.docx");

        doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());
        
        doc.save(getArtifactsDir() + "WorkingWithBarcodeGenerator.BarcodeGenerator.pdf");
        //ExEnd:BarcodeGenerator
    }
}
