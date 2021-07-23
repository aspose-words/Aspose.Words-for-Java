package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.text.MessageFormat;

public class ConvertDocumentToMarkdown {

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub
		String dataDir = Utils.getDataDir(ConvertDocumentToMarkdown.class);
		
		SaveAsMarkdown(dataDir);
		SupportedMarkdownFeatures(dataDir);
        BoldText();
        ItalicText();
        Strikethrough();
        InlineCode();
        Autolink();
        Link();
        Image();
        HorizontalRule();
        Heading();
        SetextHeading();
        IndentedCode();
        FencedCode();
        Quote();
        BulletedList();
        OrderedList();
        Table();
	}
	
	private static void SaveAsMarkdown(String dataDir) throws Exception {
		// ExStart:SaveAsMD
        // Load the document from disk.
        Document doc = new Document(dataDir + "Test.docx");

        // Save the document to Markdown format.
        doc.save(dataDir + "SaveDocx2Markdown.md");
        // ExEnd:SaveAsMD
	}
	
	private static void SupportedMarkdownFeatures(String dataDir) throws Exception
    {
        // ExStart:SupportedMarkdownFeatures
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
        builder.insertHyperlink("Aspose","https://www.aspose.com", false);
        builder.write("Aspose");

        // Save your document as a Markdown file.
        doc.save(dataDir + "example.md");
        // ExEnd:SupportedMarkdownFeatures
    }

    public static void BoldText() throws Exception {
        //ExStart:BoldText
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Make the text Bold.
        builder.getFont().setBold(true);
        builder.writeln("This text will be Bold");
        //ExEnd:BoldText
    }

    public static void ItalicText() throws Exception {
        //ExStart:ItalicText
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Make the text Italic.
        builder.getFont().setItalic(true);
        builder.writeln("This text will be Italic");
        //ExEnd:ItalicText
    }

    public static void Strikethrough() throws Exception {
        //ExStart:Strikethrough
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Make the text Strikethrough.
        builder.getFont().setStrikeThrough(true);
        builder.writeln("This text will be Strikethrough");
        //ExEnd:Strikethrough
    }

    public static void InlineCode() throws Exception {
        //ExStart:InlineCode
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

    public static void Autolink() throws Exception {
        //ExStart:Autolink
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert hyperlink.
        builder.insertHyperlink("https://www.aspose.com", "https://www.aspose.com", false);
        builder.insertHyperlink("email@aspose.com", "mailto:email@aspose.com", false);
        //ExEnd:Autolink
    }

    public static void Link() throws Exception {
        //ExStart:Link
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert hyperlink.
        builder.insertHyperlink("Aspose", "https://www.aspose.com", false);
        //ExEnd:Link
    }

    public static void Image() throws Exception {
        //ExStart:Image
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert image.
        Shape shape = new Shape(builder.getDocument(), ShapeType.IMAGE);
        shape.setWrapType(WrapType.INLINE);
        shape.getImageData().setSourceFullName("/attachment/1456/pic001.png");
        shape.getImageData().setTitle("title");
        builder.insertNode(shape);
        //ExEnd:Image
    }

    public static void HorizontalRule() throws Exception {
        //ExStart:HorizontalRule
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Insert horizontal rule.
        builder.insertHorizontalRule();
        //ExEnd:HorizontalRule
    }

    public static void Heading() throws Exception {
        //ExStart:Heading
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // By default Heading styles in Word may have Bold and Italic formatting.
        //If we do not want to be emphasized, set these properties explicitly to false.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.writeln("This is an H1 tag");
        //ExEnd:Heading
    }

    public static void SetextHeading() throws Exception {
        //ExStart:SetextHeading
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        builder.getParagraphFormat().setStyleName("Heading 1");
        builder.writeln("This is an H1 tag");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        Style setexHeading1 = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "SetexHeading1");
        builder.getParagraphFormat().setStyle(setexHeading1);
        builder.getDocument().getStyles().get("SetexHeading1").setBaseStyleName("Heading 1");
        builder.writeln("Setex Heading level 1");

        builder.getParagraphFormat().setStyle(builder.getDocument().getStyles().get("Heading 3"));
        builder.writeln("This is an H3 tag");

        // Reset styles from the previous paragraph to not combine styles between paragraphs.
        builder.getFont().setBold(false);
        builder.getFont().setItalic(false);

        Style setexHeading2 = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "SetexHeading2");
        builder.getParagraphFormat().setStyle(setexHeading2);
        builder.getDocument().getStyles().get("SetexHeading2").setBaseStyleName("Heading 3");

        // Setex heading level will be reset to 2 if the base paragraph has a Heading level greater than 2.
        builder.writeln("Setex Heading level 2");
        //ExEnd:SetextHeading
    }

    public static void IndentedCode() throws Exception {
        //ExStart:IndentedCode
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        Style indentedCode = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "IndentedCode");
        builder.getParagraphFormat().setStyle(indentedCode);
        builder.writeln("This is an indented code");
        //ExEnd:IndentedCode
    }

    public static void FencedCode() throws Exception {
        //ExStart:FencedCode
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

    public static void Quote() throws Exception {
        //ExStart:Quote
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // By default a document stores blockquote style for the first level.
        builder.getParagraphFormat().setStyleName("Quote");
        builder.writeln("Blockquote");

        // Create styles for nested levels through style inheritance.
        Style quoteLevel2 = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "Quote1");
        builder.getParagraphFormat().setStyle(quoteLevel2);
        builder.getDocument().getStyles().get("Quote1").setBaseStyleName("Quote");
        builder.writeln("1. Nested blockquote");
        //ExEnd:Quote
    }

    public static void BulletedList() throws Exception {
        //ExStart:BulletedList
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

    public static void OrderedList() throws Exception {
        //ExStart:OrderedList
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().applyBulletDefault();
        builder.getListFormat().getList().getListLevels().get(0).setNumberFormat(MessageFormat.format("{0}.", (char)0));
        builder.getListFormat().getList().getListLevels().get(1).setNumberFormat(MessageFormat.format("{0}.", (char)1));

        builder.writeln("Item 1");
        builder.writeln("Item 2");

        builder.getListFormat().listIndent();

        builder.writeln("Item 2a");
        builder.writeln("Item 2b");
        //ExEnd:OrderedList
    }

    public static void Table() throws Exception {
        //ExStart:Table
        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder();

        // Add the first row.
        builder.insertCell();
        builder.writeln("a");
        builder.insertCell();
        builder.writeln("b");

        // Add the second row.
        builder.insertCell();
        builder.writeln("c");
        builder.insertCell();
        builder.writeln("d");
        //ExEnd:Table
    }
}