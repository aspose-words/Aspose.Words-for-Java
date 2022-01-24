package DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;


class WorkingWithBarcodeGenerator extends DocsExamplesBase
{
    @Test
    public void generateACustomBarCodeImage() throws Exception
    {
        //ExStart:GenerateACustomBarCodeImage
        Document doc = new Document(getMyDir() + "Field sample - BARCODE.docx");

        doc.getFieldOptions().setBarcodeGenerator(new CustomBarcodeGenerator());
        
        doc.save(getArtifactsDir() + "WorkingWithBarcodeGenerator.GenerateACustomBarCodeImage.pdf");
        //ExEnd:GenerateACustomBarCodeImage
    }
}
