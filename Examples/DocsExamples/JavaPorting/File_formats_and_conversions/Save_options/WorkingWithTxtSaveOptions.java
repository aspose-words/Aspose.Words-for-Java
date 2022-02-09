package DocsExamples.File_Formats_and_Conversions.Save_Options;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.TxtSaveOptions;


public class WorkingWithTxtSaveOptions extends DocsExamplesBase
{
    @Test
    public void addBidiMarks() throws Exception
    {
        //ExStart:AddBidiMarks
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
    public void useTabCharacterPerLevelForListIndentation() throws Exception
    {
        //ExStart:UseTabCharacterPerLevelForListIndentation
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

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.UseTabCharacterPerLevelForListIndentation.txt", saveOptions);
        //ExEnd:UseTabCharacterPerLevelForListIndentation
    }

    @Test
    public void useSpaceCharacterPerLevelForListIndentation() throws Exception
    {
        //ExStart:UseSpaceCharacterPerLevelForListIndentation
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

        doc.save(getArtifactsDir() + "WorkingWithTxtSaveOptions.UseSpaceCharacterPerLevelForListIndentation.txt", saveOptions);
        //ExEnd:UseSpaceCharacterPerLevelForListIndentation
    }
}
