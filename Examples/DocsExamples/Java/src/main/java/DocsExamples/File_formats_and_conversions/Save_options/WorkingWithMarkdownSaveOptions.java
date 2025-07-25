package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.MarkdownSaveOptions;
import com.aspose.words.TableContentAlignment;

import java.io.ByteArrayOutputStream;

@Test
public class WorkingWithMarkdownSaveOptions extends DocsExamplesBase {
    @Test
    public void markdownTableContentAlignment() throws Exception {
        //ExStart:MarkdownTableContentAlignment
        //GistId:19de942ef8827201c1dca99f76c59133
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
        //ExEnd:MarkdownTableContentAlignment
    }

    @Test
    public void imagesFolder() throws Exception {
        //ExStart:ImagesFolder
        //GistId:51b4cb9c451832f23527892e19c7bca6
        Document doc = new Document(getMyDir() + "Image bullet points.docx");

        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        {
            saveOptions.setImagesFolder(getArtifactsDir() + "Images");
        }

        try (ByteArrayOutputStream stream = new ByteArrayOutputStream()) {
            doc.save(stream, saveOptions);
        }
        //ExEnd:ImagesFolder
    }
}

