package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class WorkingWithMarkdown extends DocsExamplesBase
{
    @Test
    public void createMarkdownDocument() throws Exception
    {
        //ExStart:CreateMarkdownDocument
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify the "Heading 1" style for the paragraph.
        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.writeln("Heading 1");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.getParagraphFormat().setStyleName("Normal");

        // Insert horizontal rule.
        builder.insertHorizontalRule();

        // Specify the ordered list.
        builder.insertParagraph();
        builder.getListFormat().applyNumberDefault();

        // Specify the Italic emphasis for the text.
        builder.getFont().setItalic(true);
        builder.writeln("Italic Text");
        builder.getFont().setItalic(false);

        // Specify the Bold emphasis for the text.
        builder.getFont().setBold(true);
        builder.writeln("Bold Text");
        builder.getFont().setBold(false);

        // Specify the StrikeThrough emphasis for the text.
        builder.getFont().setStrikeThrough(true);
        builder.writeln("StrikeThrough Text");
        builder.getFont().setStrikeThrough(false);

        // Stop paragraphs numbering.
        builder.getListFormat().removeNumbers();

        // Specify the "Quote" style for the paragraph.
        builder.getParagraphFormat().setStyleName("Quote");
        builder.writeln("A Quote block");

        // Specify nesting Quote.
        Style nestedQuote = doc.getStyles().add(StyleType.PARAGRAPH, "Quote1");
        nestedQuote.setBaseStyleName("Quote");
        builder.getParagraphFormat().setStyleName("Quote1");
        builder.writeln("A nested Quote block");

        // Reset paragraph style to Normal to stop Quote blocks. 
        builder.getParagraphFormat().setStyleName("Normal");

        // Specify a Hyperlink for the desired text.
        builder.getFont().setBold(true);
        // Note, the text of hyperlink can be emphasized.
        builder.insertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.getFont().setBold(false);

        // Insert a simple table.
        builder.startTable();
        builder.insertCell();
        builder.write("Cell1");
        builder.insertCell();
        builder.write("Cell2");
        builder.endTable();

        // Save your document as a Markdown file.
        doc.save(getArtifactsDir() + "WorkingWithMarkdown.CreateMarkdownDocument.md");
        //ExEnd:CreateMarkdownDocument
    }

    @Test
    public void readMarkdownDocument() throws Exception
    {
        //ExStart:ReadMarkdownDocument
        Document doc = new Document(getMyDir() + "Quotes.md");

        // Let's remove Heading formatting from a Quote in the very last paragraph.
        Paragraph paragraph = doc.getFirstSection().getBody().getLastParagraph();
        paragraph.getParagraphFormat().setStyle(doc.getStyles().get("Quote"));

        doc.save(getArtifactsDir() + "WorkingWithMarkdown.ReadMarkdownDocument.md");
        //ExEnd:ReadMarkdownDocument
    }

    @Test
    public void emphases() throws Exception
    {
        //ExStart:Emphases
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Markdown treats asterisks (*) and underscores (_) as indicators of emphasis.");
        builder.write("You can write ");

        builder.getFont().setBold(true);
        builder.write("bold");

        builder.getFont().setBold(false);
        builder.write(" or ");

        builder.getFont().setItalic(true);
        builder.write("italic");

        builder.getFont().setItalic(false);
        builder.writeln(" text. ");

        builder.write("You can also write ");
        builder.getFont().setBold(true);

        builder.getFont().setItalic(true);
        builder.write("BoldItalic");

        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);
        builder.write("text.");

        builder.getDocument().save(getArtifactsDir() + "WorkingWithMarkdown.Emphases.md");
        //ExEnd:Emphases
    }

    @Test
    public void headings() throws Exception
    {
        //ExStart:Headings
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default Heading styles in Word may have bold and italic formatting.
        // If we do not want the text to be emphasized, set these properties explicitly to false.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        builder.writeln("The following produces headings:");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Heading1");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 2"));
        builder.writeln("Heading2");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 3"));
        builder.writeln("Heading3");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 4"));
        builder.writeln("Heading4");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 5"));
        builder.writeln("Heading5");
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 6"));
        builder.writeln("Heading6");

        // Note that the emphases are also allowed inside Headings.
        builder.getFont().setBold(true);
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Bold Heading1");

        doc.save(getArtifactsDir() + "WorkingWithMarkdown.Headings.md");
        //ExEnd:Headings
    }

    @Test
    public void blockQuotes() throws Exception
    {
        //ExStart:BlockQuotes
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("We support blockquotes in Markdown:");
        
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Quote"));
        builder.writeln("Lorem");
        builder.writeln("ipsum");
        
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Normal"));
        builder.writeln("The quotes can be of any level and can be nested:");
        
        Style quoteLevel3 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote2");
        builder.getParagraphFormat().setStyle(quoteLevel3);
        builder.writeln("Quote level 3");
        
        Style quoteLevel4 = doc.getStyles().add(StyleType.PARAGRAPH, "Quote3");
        builder.getParagraphFormat().setStyle(quoteLevel4);
        builder.writeln("Nested quote level 4");
        
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Quote"));
        builder.writeln();
        builder.writeln("Back to first level");
        
        Style quoteLevel1WithHeading = doc.getStyles().add(StyleType.PARAGRAPH, "Quote Heading 3");
        builder.getParagraphFormat().setStyle(quoteLevel1WithHeading);
        builder.write("Headings are allowed inside Quotes");

        doc.save(getArtifactsDir() + "WorkingWithMarkdown.BlockQuotes.md");
        //ExEnd:BlockQuotes
    }

    @Test
    public void horizontalRule() throws Exception
    {
        //ExStart:HorizontalRule
        DocumentBuilder builder = new DocumentBuilder(new Document());

        builder.writeln("We support Horizontal rules (Thematic breaks) in Markdown:");
        builder.insertHorizontalRule();

        builder.getDocument().save(getArtifactsDir() + "WorkingWithMarkdown.HorizontalRuleExample.md");
        //ExEnd:HorizontalRule
    }

    @Test
    public void useWarningSource() throws Exception
    {
        //ExStart:UseWarningSourceMarkdown
        Document doc = new Document(getMyDir() + "Emphases markdown warning.docx");

        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.setWarningCallback(warnings);

        doc.save(getArtifactsDir() + "WorkingWithMarkdown.UseWarningSource.md");

        for (WarningInfo warningInfo : warnings)
        {
            if (warningInfo.getSource() == WarningSource.MARKDOWN)
                System.out.println(warningInfo.getDescription());
        }
        //ExEnd:UseWarningSourceMarkdown
    }
}
