package DocsExamples.Programming_with_documents.Contents_managment;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;

@Test
public class WorkingWithRanges extends DocsExamplesBase
{
    @Test
    public void rangesDeleteText() throws Exception
    {
        //ExStart:RangesDeleteText
        Document doc = new Document(getMyDir() + "Document.docx");
        doc.getSections().get(0).getRange().delete();
        //ExEnd:RangesDeleteText
    }

    @Test
    public void rangesGetText() throws Exception
    {
        //ExStart:RangesGetText
        Document doc = new Document(getMyDir() + "Document.docx");
        String text = doc.getRange().getText();
        //ExEnd:RangesGetText
    }
}
