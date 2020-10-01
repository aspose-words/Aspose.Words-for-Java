package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.MarkdownSaveOptions;
import com.aspose.words.SaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class ConvertDocumentToMarkdown {

	public static void main(String[] args) throws Exception{
		// TODO Auto-generated method stub
		String dataDir = Utils.getDataDir(ConvertDocumentToMarkdown.class);
		
		SaveAsMarkdown(dataDir);
		SupportedMarkdownFeatures(dataDir);
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
}
