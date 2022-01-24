package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.ListTemplate;
import com.aspose.words.List;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.OoxmlCompliance;
import java.awt.Color;
import java.text.MessageFormat;

import com.aspose.words.ListLevelAlignment;

@Test
public class WorkingWithList extends DocsExamplesBase
{
    @Test
    public void restartListAtEachSection() throws Exception
    {
        //ExStart:RestartListAtEachSection
        Document doc = new Document();
        
        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        List list = doc.getLists().get(0);
        list.isRestartAtEachSection(true);

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getListFormat().setList(list);

        for (int i = 1; i < 45; i++)
        {
            builder.writeln(MessageFormat.format("List Item {0}", i));

            if (i == 15)
                builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        }

        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376.
        OoxmlSaveOptions options = new OoxmlSaveOptions(); { options.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL); }

        doc.save(getArtifactsDir() + "WorkingWithList.RestartListAtEachSection.docx", options);
        //ExEnd:RestartListAtEachSection
    }

    @Test
    public void specifyListLevel() throws Exception
    {
        //ExStart:SpecifyListLevel
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a numbered list based on one of the Microsoft Word list templates
        // and apply it to the document builder's current paragraph.
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.NUMBER_ARABIC_DOT));

        // There are nine levels in this list, let's try them all.
        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        // Create a bulleted list based on one of the Microsoft Word list templates
        // and apply it to the document builder's current paragraph.
        builder.getListFormat().setList(doc.getLists().add(ListTemplate.BULLET_DIAMONDS));

        for (int i = 0; i < 9; i++)
        {
            builder.getListFormat().setListLevelNumber(i);
            builder.writeln("Level " + i);
        }

        // This is a way to stop list formatting.
        builder.getListFormat().setList(null);

        builder.getDocument().save(getArtifactsDir() + "WorkingWithList.SpecifyListLevel.docx");
        //ExEnd:SpecifyListLevel
    }

    @Test
    public void restartListNumber() throws Exception
    {
        //ExStart:RestartListNumber
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a list based on a template.
        List list1 = doc.getLists().add(ListTemplate.NUMBER_ARABIC_PARENTHESIS);
        list1.getListLevels().get(0).getFont().setColor(Color.RED);
        list1.getListLevels().get(0).setAlignment(ListLevelAlignment.RIGHT);

        builder.writeln("List 1 starts below:");
        builder.getListFormat().setList(list1);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        // To reuse the first list, we need to restart numbering by creating a copy of the original list formatting.
        List list2 = doc.getLists().addCopy(list1);

        // We can modify the new list in any way, including setting a new start number.
        list2.getListLevels().get(0).setStartAt(10);

        builder.writeln("List 2 starts below:");
        builder.getListFormat().setList(list2);
        builder.writeln("Item 1");
        builder.writeln("Item 2");
        builder.getListFormat().removeNumbers();

        builder.getDocument().save(getArtifactsDir() + "WorkingWithList.RestartListNumber.docx");
        //ExEnd:RestartListNumber
    }
}
