package DocsExamples.File_Formats_and_Conversions.Save_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.MarkdownSaveOptions;
import com.aspose.words.TableContentAlignment;
import com.aspose.ms.System.IO.MemoryStream;


public class WorkingWithMarkdownSaveOptions extends DocsExamplesBase
{
    @Test
    public void exportIntoMarkdownWithTableContentAlignment() throws Exception
    {
        //ExStart:ExportIntoMarkdownWithTableContentAlignment
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);
        builder.write("Cell1");
        builder.insertCell();
        builder.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);
        builder.write("Cell2");

        // Makes all paragraphs inside the table to be aligned.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        {
            saveOptions.setTableContentAlignment(TableContentAlignment.LEFT);
        }
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.LeftTableContentAlignment.md", saveOptions);

        saveOptions.setTableContentAlignment(TableContentAlignment.RIGHT);
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.RightTableContentAlignment.md", saveOptions);

        saveOptions.setTableContentAlignment(TableContentAlignment.CENTER);
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.CenterTableContentAlignment.md", saveOptions);

        // The alignment in this case will be taken from the first paragraph in corresponding table column.
        saveOptions.setTableContentAlignment(TableContentAlignment.AUTO);
        doc.save(getArtifactsDir() + "WorkingWithMarkdownSaveOptions.AutoTableContentAlignment.md", saveOptions);
        //ExEnd:ExportIntoMarkdownWithTableContentAlignment
    }

    @Test
    public void setImagesFolder() throws Exception
    {
        //ExStart:SetImagesFolder
        Document doc = new Document(getMyDir() + "Image bullet points.docx");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions(); { saveOptions.setImagesFolder(getArtifactsDir() + "Images"); }

        MemoryStream stream = new MemoryStream();
        try /*JAVA: was using*/
    	{
            doc.save(stream, saveOptions);
    	}
        finally { if (stream != null) stream.close(); }
        //ExEnd:SetImagesFolder
    }
}

