package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.FootnotePosition;
import com.aspose.words.EndnotePosition;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FootnoteType;
import com.aspose.words.EndnoteOptions;
import com.aspose.words.FootnoteNumberingRule;

@Test
public class WorkingWithFootnoteAndEndnote extends DocsExamplesBase
{
    @Test
    public void setFootNoteColumns() throws Exception
    {
        //ExStart:SetFootNoteColumns
        Document doc = new Document(getMyDir() + "Document.docx");

        // Specify the number of columns with which the footnotes area is formatted.
        doc.getFootnoteOptions().setColumns(3);
        
        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetFootNoteColumns.docx");
        //ExEnd:SetFootNoteColumns
    }

    @Test
    public void setFootnoteAndEndNotePosition() throws Exception
    {
        //ExStart:SetFootnoteAndEndNotePosition
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        
        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetFootnoteAndEndNotePosition.docx");
        //ExEnd:SetFootnoteAndEndNotePosition
    }

    @Test
    public void setEndnoteOptions() throws Exception
    {
        //ExStart:SetEndnoteOptions
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("Some text");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Footnote text.");

        EndnoteOptions option = doc.getEndnoteOptions();
        option.setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        option.setPosition(EndnotePosition.END_OF_SECTION);

        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetEndnoteOptions.docx");
        //ExEnd:SetEndnoteOptions
    }
}
