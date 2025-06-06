package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class WorkingWithMarkdown extends DocsExamplesBase
{
    @Test
    public void boldText() throws Exception
    {
        //ExStart:BoldText
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Make the text Bold.
        builder.getFont().setBold(true);
        builder.writeln("This text will be Bold");
        //ExEnd:BoldText
    }

    @Test
    public void italicText() throws Exception
    {
        //ExStart:ItalicText
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Make the text Italic.
        builder.getFont().setItalic(true);
        builder.writeln("This text will be Italic");
        //ExEnd:ItalicText
    }

    @Test
    public void strikethrough() throws Exception
    {
        //ExStart:Strikethrough
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Make the text Strikethrough.
        builder.getFont().setStrikeThrough(true);
        builder.writeln("This text will be StrikeThrough");
        //ExEnd:Strikethrough
    }

    @Test
    public void inlineCode() throws Exception
    {
        //ExStart:InlineCode
        //GistId:51b4cb9c451832f23527892e19c7bca6
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Number of backticks is missed, one backtick will be used by default.
        Style inlineCode1BackTicks = builder.getDocument().getStyles().add(StyleType.CHARACTER, "InlineCode");
        builder.getFont().setStyle(inlineCode1BackTicks);
        builder.writeln("Text with InlineCode style with 1 backtick");

        // There will be 3 backticks.
        Style inlineCode3BackTicks = builder.getDocument().getStyles().add(StyleType.CHARACTER, "InlineCode.3");
        builder.getFont().setStyle(inlineCode3BackTicks);
        builder.writeln("Text with InlineCode style with 3 backtick");
        //ExEnd:InlineCode
    }

    @Test
    public void autolink() throws Exception
    {
        //ExStart:Autolink
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert hyperlink.
        builder.insertHyperlink("https://www.aspose.com", "https://www.aspose.com", false);
        builder.insertHyperlink("email@aspose.com", "mailto:email@aspose.com", false);
        //ExEnd:Autolink
    }

    @Test
    public void link() throws Exception
    {
        //ExStart:Link
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert hyperlink.
        builder.insertHyperlink("Aspose", "https://www.aspose.com", false);
        //ExEnd:Link
    }

    @Test
    public void image() throws Exception
    {
        //ExStart:Image
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert image.
        Shape shape = builder.insertImage(getImagesDir() + "Logo.jpg");
        shape.getImageData().setTitle("title");
        //ExEnd:Image
    }

    @Test
    public void horizontalRule() throws Exception
    {
        //ExStart:HorizontalRule
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert horizontal rule.
        builder.insertHorizontalRule();
        //ExEnd:HorizontalRule
    }

    @Test
    public void heading() throws Exception
    {
        //ExStart:Heading
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default Heading styles in Word may have Bold and Italic formatting.
        //If we do not want to be emphasized, set these properties explicitly to false.
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

        // Note, emphases are also allowed inside Headings:
        builder.getFont().setBold(true);
        builder.getParagraphFormat().setStyle(doc.getStyles().get("Heading 1"));
        builder.writeln("Bold Heading1");

        doc.save(getArtifactsDir() + "WorkingWithMarkdown.Heading.md");
        //ExEnd:Heading
    }

    @Test
    public void setextHeading() throws Exception
    {
        //ExStart:SetextHeading
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.writeln("This is an H1 tag");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        Style setexHeading1 = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "SetextHeading1");
        builder.getParagraphFormat().setStyle(setexHeading1);
        builder.getDocument().getStyles().get("SetextHeading1").setBaseStyleName("Heading 1");
        builder.writeln("Setext Heading level 1");

        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 3"));
        builder.writeln("This is an H3 tag");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        Style setexHeading2 = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "SetextHeading2");
        builder.getParagraphFormat().setStyle(setexHeading2);
        builder.getDocument().getStyles().get("SetextHeading2").setBaseStyleName("Heading 3");

        // Setex heading level will be reset to 2 if the base paragraph has a Heading level greater than 2.
        builder.writeln("Setext Heading level 2");
        //ExEnd:SetextHeading

        builder.getDocument().save(getArtifactsDir() + "WorkingWithMarkdown.SetextHeading.md");
    }

    @Test
    public void indentedCode() throws Exception
    {
        //ExStart:IndentedCode
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        Style indentedCode = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "IndentedCode");
        builder.getParagraphFormat().setStyle(indentedCode);
        builder.writeln("This is an indented code");
        //ExEnd:IndentedCode
    }

    @Test
    public void fencedCode() throws Exception
    {
        //ExStart:FencedCode
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        Style fencedCode = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "FencedCode");
        builder.getParagraphFormat().setStyle(fencedCode);
        builder.writeln("This is an fenced code");

        Style fencedCodeWithInfo = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "FencedCode.C#");
        builder.getParagraphFormat().setStyle(fencedCodeWithInfo);
        builder.writeln("This is a fenced code with info string");
        //ExEnd:FencedCode
    }

    @Test
    public void quote() throws Exception
    {
        //ExStart:Quote
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default a document stores blockquote style for the first level.
        builder.getParagraphFormat().setStyleName("Quote");
        builder.writeln("Blockquote");

        // Create styles for nested levels through style inheritance.
        Style quoteLevel2 = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "Quote1");
        builder.getParagraphFormat().setStyle(quoteLevel2);
        builder.getDocument().getStyles().get("Quote1").setBaseStyleName("Quote");
        builder.writeln("1. Nested blockquote");

        doc.save(getArtifactsDir() + "WorkingWithMarkdown.Quote.md");
        //ExEnd:Quote
    }

    @Test
    public void bulletedList() throws Exception
    {
        //ExStart:BulletedList
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        builder.getListFormat().applyBulletDefault();
        builder.getListFormat().getList().getListLevels().get(0).setNumberFormat("-");

        builder.writeln("Item 1");
        builder.writeln("Item 2");

        builder.getListFormat().listIndent();

        builder.writeln("Item 2a");
        builder.writeln("Item 2b");
        //ExEnd:BulletedList
    }

    @Test
    public void orderedList() throws Exception
    {
        //ExStart:OrderedList
        //GistId:0697355b7f872839932388d269ed6a63
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().applyNumberDefault();

        builder.writeln("Item 1");
        builder.writeln("Item 2");

        builder.getListFormat().listIndent();

        builder.writeln("Item 2a");
        builder.writeln("Item 2b");
        //ExEnd:OrderedList
    }

    @Test
    public void table() throws Exception
    {
        //ExStart:Table
        //GistId:0697355b7f872839932388d269ed6a63
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Add the first row.
        builder.insertCell();
        builder.writeln("a");
        builder.insertCell();
        builder.writeln("b");

        builder.endRow();

        // Add the second row.
        builder.insertCell();
        builder.writeln("c");
        builder.insertCell();
        builder.writeln("d");
        //ExEnd:Table
    }

    @Test
    public void readMarkdownDocument() throws Exception
    {
        //ExStart:ReadMarkdownDocument
        //GistId:19de942ef8827201c1dca99f76c59133
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
        //GistId:19de942ef8827201c1dca99f76c59133
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

    @Test
    public void supportedFeatures() throws Exception
    {
        //ExStart:SupportedFeatures
        //GistId:51b4cb9c451832f23527892e19c7bca6
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Specify the "Heading 1" style for the paragraph.
        builder.insertParagraph();
        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.write("Heading 1");

        // Specify the Italic emphasis for the paragraph.
        builder.insertParagraph();
        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.getParagraphFormat().setStyleName("Normal");
        builder.getFont().setItalic(true);
        builder.write("Italic Text");
        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.setItalic(false);

        // Specify a Hyperlink for the desired text.
        builder.insertParagraph();
        builder.insertHyperlink("Aspose", "https://www.aspose.com", false);
        builder.write("Aspose");

        // Save your document as a Markdown file.
        doc.save(getArtifactsDir() + "WorkingWithMarkdown.SupportedFeatures.md");
        //ExEnd:SupportedFeatures
    }
}
