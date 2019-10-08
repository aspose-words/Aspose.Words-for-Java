package com.aspose.words.examples.programming_documents.MarkdownFeatures;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class WorkingWithMarkdownFeatures {

    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        String dataDir = Utils.getDataDir(WorkingWithMarkdownFeatures.class);
        MarkdownDocumentWithEmphases(dataDir);
        MarkdownDocumentWithHeadings(dataDir);
        MarkdownDocumentWithBlockQuotes(dataDir);
        MarkdownDocumentWithHorizontalRule(dataDir);
        ReadMarkdownDocument(dataDir);
    }

    private static void MarkdownDocumentWithEmphases(String dataDir) throws Exception {
        // ExStart:MarkdownDocumentWithEmphases
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

        builder.getDocument().save("EmphasesExample.md");
        // ExEnd:MarkdownDocumentWithEmphases
        System.out.println("\nMarkdown Document With Emphases Produced!\nFile saved at " + dataDir);
    }

    private static void MarkdownDocumentWithHeadings(String dataDir) throws Exception {
        //ExStart: MarkdownDocumentWithHeadings
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // By default Heading styles in Word may have bold and italic formatting.
        // If we do not want text to be emphasized, set these properties explicitly to false.
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

        doc.save(dataDir + "HeadingsExample.md");
        // ExEnd:MarkdownDocumentWithHeadings
        System.out.println("\nMarkdown Document With Headings Produced!\nFile saved at " + dataDir);
    }

    private static void MarkdownDocumentWithBlockQuotes(String dataDir) throws Exception {
        // ExStart: MarkdownDocumentWithBlockQuotes
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

        doc.save(dataDir + "QuotesExample.md");
        // ExEnd: MarkdownDocumentWithBlockQuotes
        System.out.println("\nMarkdown Document With BlockQuotes Produced!\nFile saved at " + dataDir);
    }

    private static void MarkdownDocumentWithHorizontalRule(String dataDir) throws Exception {
        // ExStart: MarkdownDocumentWithHorizontalRule
        DocumentBuilder builder = new DocumentBuilder(new Document());

        builder.writeln("We support Horizontal rules (Thematic breaks) in Markdown:");
        builder.insertHorizontalRule();

        builder.getDocument().save(dataDir + "HorizontalRuleExample.md");
        // ExEnd: MarkdownDocumentWithHorizontalRule
        System.out.println("\nMarkdown Document With Horizontal Rule Produced!\nFile saved at " + dataDir);
    }

    private static void ReadMarkdownDocument(String dataDir) throws Exception {
        // ExStart: ReadMarkdownDocument
        // This is Markdown document that was produced in example of MarkdownDocumentWithBlockQuotes.
        Document doc = new Document(dataDir + "QuotesExample.md");

        // Let's remove Heading formatting from a Quote in the very last paragraph.
        Paragraph paragraph = doc.getFirstSection().getBody().getLastParagraph();
        paragraph.getParagraphFormat().setStyle(doc.getStyles().get("Quote"));

        doc.save(dataDir + "QuotesModifiedExample.md");
        // ExEnd: ReadMarkdownDocument
        System.out.println("\nRead Markdown Document!\nFile saved at " + dataDir);
    }

}
