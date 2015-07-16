package featurescomparison.workingwithdocuments.inserttableofcontents.java;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleIdentifier;

public class AsposeTableOfContents 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/inserttableofcontents/data/";
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Insert a table of contents at the beginning of the document.
		builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

		// The newly inserted table of contents will be initially empty.
		// It needs to be populated by updating the fields in the document.
		doc.updateFields();
		
		doc.save(dataPath + "AsposeTableOfContents.doc");

		// ==================================
		// Table Of Contents with Headings
		// ==================================
		
		Document doc1 = new Document();

		// Create a document builder to insert content with into document.
		builder = new DocumentBuilder(doc1);

		// Insert a table of contents at the beginning of the document.
		builder.insertTableOfContents("\\o \"1-3\" \\h \\z \\u");

		// Start the actual document content on the second page.
		builder.insertBreak(BreakType.PAGE_BREAK);

		// Build a document with complex structure by applying different heading styles thus creating TOC entries.
		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

		builder.writeln("Heading 1");

		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

		builder.writeln("Heading 1.1");
		builder.writeln("Heading 1.2");

		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);

		builder.writeln("Heading 2");
		builder.writeln("Heading 3");

		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

		builder.writeln("Heading 3.1");

		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_3);

		builder.writeln("Heading 3.1.1");
		builder.writeln("Heading 3.1.2");
		builder.writeln("Heading 3.1.3");

		builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_2);

		builder.writeln("Heading 3.2");
		builder.writeln("Heading 3.3");

		// Call the method below to update the TOC.
		doc1.updateFields();
		doc1.save(dataPath + "AsposeTOCHeadings.doc");
		System.out.println("Done.");
	}
}
