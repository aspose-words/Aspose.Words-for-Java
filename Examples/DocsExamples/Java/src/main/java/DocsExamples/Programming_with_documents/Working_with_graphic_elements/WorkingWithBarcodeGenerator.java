package DocsExamples.Programming_with_documents.Working_with_graphic_elements;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;

@Test
public class WorkingWithBarcodeGenerator extends DocsExamplesBase
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
