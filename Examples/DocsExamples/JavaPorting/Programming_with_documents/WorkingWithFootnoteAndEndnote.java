package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.FootnotePosition;
import com.aspose.words.EndnotePosition;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FootnoteType;
import com.aspose.words.EndnoteOptions;
import com.aspose.words.FootnoteNumberingRule;


class WorkingWithFootnotes extends DocsExamplesBase
{
    @Test
    public void setFootnoteColumns() throws Exception
    {
        //ExStart:SetFootnoteColumns
        //GistId:3b39c2019380ee905e7d9596494916a4
        Document doc = new Document(getMyDir() + "Document.docx");

        // Specify the number of columns with which the footnotes area is formatted.
        doc.getFootnoteOptions().setColumns(3);
        
        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetFootnoteColumns.docx");
        //ExEnd:SetFootnoteColumns
    }

    @Test
    public void setFootnoteAndEndnotePosition() throws Exception
    {
        //ExStart:SetFootnoteAndEndnotePosition
        //GistId:3b39c2019380ee905e7d9596494916a4
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        
        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetFootnoteAndEndnotePosition.docx");
        //ExEnd:SetFootnoteAndEndnotePosition
    }

    @Test
    public void setEndnoteOptions() throws Exception
    {
        //ExStart:SetEndnoteOptions
        //GistId:3b39c2019380ee905e7d9596494916a4
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
