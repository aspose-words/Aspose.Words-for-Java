package DocsExamples.Getting_started;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.util.Date;

@Test
public class HelloWorld extends DocsExamplesBase
{
    @Test
    public void helloWorld() throws Exception
    {
        //ExStart:HelloWorld
        //GistId:4e111aa3d11a41428c8a0cadfc23b972
        Document docA = new Document();
        DocumentBuilder builder = new DocumentBuilder(docA);

        // Insert text to the document start.
        builder.moveToDocumentStart();
        builder.write("First Hello World paragraph");

        Document docB = new Document(getMyDir() + "Document.docx");
        // Add document B to the and of document A, preserving document B formatting.
        docA.appendDocument(docB, ImportFormatMode.KEEP_SOURCE_FORMATTING);

        docA.save(getArtifactsDir() + "HelloWorld.SimpleHelloWorld.pdf");
        //ExEnd:HelloWorld
    }
}
