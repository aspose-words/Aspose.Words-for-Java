package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.TxtSaveOptions;

@Test
public class WorkingWithTxtSaveOptions extends DocsExamplesBase
{
    @Test
    public void addBidiMarks() throws Exception
    {
        //ExStart:AddBidiMarks
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.getParagraphFormat().setBidi(true);
        builder.writeln("שלום עולם!");
        builder.writeln("مرحبا بالعالم!");

        TxtSaveOptions saveOptions = new TxtSaveOptions(); { saveOptions.setAddBidiMarks(true); }

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
        //ExEnd:AddBidiMarks
    }

    @Test
    public void useTabForListIndentation() throws Exception
    {
        //ExStart:UseTabForListIndentation
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list with three levels of indentation.
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent(); 
        builder.write("Item 3");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.getListIndentation().setCount(1);
        saveOptions.getListIndentation().setCharacter('\t');

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.UseTabForListIndentation.txt", saveOptions);
        //ExEnd:UseTabForListIndentation
    }

    @Test
    public void useSpaceForListIndentation() throws Exception
    {
        //ExStart:UseSpaceForListIndentation
        //GistId:ddafc3430967fb4f4f70085fa577d01a
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list with three levels of indentation.
        builder.getListFormat().applyNumberDefault();
        builder.writeln("Item 1");
        builder.getListFormat().listIndent();
        builder.writeln("Item 2");
        builder.getListFormat().listIndent(); 
        builder.write("Item 3");

        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.getListIndentation().setCount(3);
        saveOptions.getListIndentation().setCharacter(' ');

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.UseSpaceForListIndentation.txt", saveOptions);
        //ExEnd:UseSpaceForListIndentation
    }
}
