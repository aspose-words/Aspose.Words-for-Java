package DocsExamples.Getting_started;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ImportFormatMode;


public class HelloWorld extends DocsExamplesBase
{
    @Test
    public void simpleHelloWorld() throws Exception
    {
        //ExStart:HelloWorld
        //GistId:542a463e1857480986d18ec296ed43d5
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

