package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class WorkingWithFootnoteAndEndnote extends DocsExamplesBase {
    @Test
    public void setFootnoteColumns() throws Exception {
        //ExStart:SetFootnoteColumns
        //GistId:1cb8150f6ed075f764a1b8f2928f318e
        Document doc = new Document(getMyDir() + "Document.docx");

        // Specify the number of columns with which the footnotes area is formatted.
        doc.getFootnoteOptions().setColumns(3);

        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetFootnoteColumns.docx");
        //ExEnd:SetFootnoteColumns
    }

    @Test
    public void setFootnoteAndEndnotePosition() throws Exception {
        //ExStart:SetFootnoteAndEndnotePosition
        //GistId:1cb8150f6ed075f764a1b8f2928f318e
        Document doc = new Document(getMyDir() + "Document.docx");

        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);

        doc.save(getArtifactsDir() + "WorkingWithFootnotes.SetFootnoteAndEndnotePosition.docx");
        //ExEnd:SetFootnoteAndEndnotePosition
    }

    @Test
    public void setEndnoteOptions() throws Exception {
        //ExStart:SetEndnoteOptions
        //GistId:1cb8150f6ed075f764a1b8f2928f318e
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
