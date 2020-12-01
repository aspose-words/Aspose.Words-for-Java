package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Comment;
import com.aspose.words.CompareOptions;
import com.aspose.words.ComparisonTargetType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldDate;
import com.aspose.words.Footnote;
import com.aspose.words.FootnoteType;
import com.aspose.words.Granularity;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Revision;
import com.aspose.words.RevisionType;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.creation.BuildTableFromDataTable;

import java.util.Date;

public class CompareTwoWordDocuments {

	private static final String dataDir = Utils.getSharedDataDir(BuildTableFromDataTable.class) + "Document/";

	public static void main(String[] args) throws Exception {
		// ExStart:CompareTwoWordDocuments
		// Example Shows Normal Comparison Case
		normalComparisonCase();

		// Case when Document has Revisions already so Comparison is not Possible
		caseWhenDocumentHasRevisions(dataDir);

		// Shows how to test that Word Documents are "Equal"
		wordDocumentsAreEqual();

		SpecifyComparisonGranularity(dataDir);
		
		AdvancedComparingProperties(dataDir);
		// ExEnd:CompareTwoWordDocuments
	}

	public static void normalComparisonCase() throws Exception {
		// ExStart:normalComparisonCase
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date()); // docA now contains changes as revisions
		// ExEnd:normalComparisonCase
	}

	public static void caseWhenDocumentHasRevisions(String dataDir) throws Exception {
		// ExStart:caseWhenDocumentHasRevisions
		// The source document doc1
		Document doc1 = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc1);
		builder.writeln("This is the original document.");

		// The target document doc2
		Document doc2 = new Document();
		builder = new DocumentBuilder(doc2);
		builder.writeln("This is the edited document.");

		// If either document has a revision, an exception will be thrown
		if (doc1.getRevisions().getCount() == 0 && doc2.getRevisions().getCount() == 0)
			doc1.compare(doc2, "authorName", new Date());

		// If doc1 and doc2 are different, doc1 now has some revisions after the comparison, which can now be viewed and processed
		if (doc1.getRevisions().getCount() == 2)
			System.out.println("Documents are equal");
		
		for (Revision r : doc1.getRevisions())
		{
			System.out.println("Revision type: " + r.getRevisionType() + ", on a node of type " + r.getParentNode().getNodeType() + "");
			System.out.println("\tChanged text: " + r.getParentNode().getText() + "");
		}

		// All the revisions in doc1 are differences between doc1 and doc2, so accepting them on doc1 transforms doc1 into doc2
		doc1.getRevisions().acceptAll();
		
		// doc1, when saved, now resembles doc2
		doc1.save(dataDir + "Document.Compare.docx");
		doc1 = new Document(dataDir + "Document.Compare.docx");
		
		if (doc1.getRevisions().getCount() == 0)
			System.out.println("Documents are equal");
		
		if (doc2.getText().trim() == doc1.getText().trim())
			System.out.println("Documents are equal");
		// ExEnd:caseWhenDocumentHasRevisions
	}

	public static void wordDocumentsAreEqual() throws Exception {
		// ExStart:wordDocumentsAreEqual
		Document docA = new Document(dataDir + "DocumentA.doc");
		Document docB = new Document(dataDir + "DocumentB.doc");
		docA.compare(docB, "user", new Date());
		if (docA.getRevisions().getCount() == 0)
			System.out.println("Documents are equal");
		else
			System.out.println("Documents are not equal");
		// ExEnd:wordDocumentsAreEqual
	}

	public static void SpecifyComparisonGranularity(String dataDir) throws Exception {
		// ExStart:SpecifyComparisonGranularity
		DocumentBuilder builderA = new DocumentBuilder(new Document());
		DocumentBuilder builderB = new DocumentBuilder(new Document());

		builderA.writeln("This is A simple word");
		builderB.writeln("This is B simple words");

		CompareOptions co = new CompareOptions();
		co.setGranularity(Granularity.CHAR_LEVEL);

		builderA.getDocument().compare(builderB.getDocument(), "author", new Date(), co);
		// ExEnd:SpecifyComparisonGranularity
	}

	public static void AdvancedComparingProperties(String dataDir) throws Exception {
		// ExStart:AdvancedComparingProperties
		// Create the original document
		Document docOriginal = new Document();
		DocumentBuilder builder = new DocumentBuilder(docOriginal);

		// Insert paragraph text with an endnote
		builder.writeln("Hello world! This is the first paragraph.");
		builder.insertFootnote(FootnoteType.ENDNOTE, "Original endnote text.");

		// Insert a table
		builder.startTable();
		builder.insertCell();
		builder.write("Original cell 1 text");
		builder.insertCell();
		builder.write("Original cell 2 text");
		builder.endTable();

		// Insert a textbox
		Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 150, 20);
		builder.moveTo(textBox.getFirstParagraph());
		builder.write("Original textbox contents");

		// Insert a DATE field
		builder.moveTo(docOriginal.getFirstSection().getBody().appendParagraph(""));
		builder.insertField(" DATE ");

		// Insert a comment
		Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", new Date());
		newComment.setText("Original comment.");
		builder.getCurrentParagraph().appendChild(newComment);

		// Insert a header
		builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
		builder.writeln("Original header contents.");

		// Create a clone of our document, which we will edit and later compare to the original
		Document docEdited = (Document)docOriginal.deepClone(true);
		Paragraph firstParagraph = docEdited.getFirstSection().getBody().getFirstParagraph();

		// Change the formatting of the first paragraph, change casing of original characters and add text
		firstParagraph.getRuns().get(0).setText("hello world! this is the first paragraph, after editing.");
		firstParagraph.getParagraphFormat().setStyle(docEdited.getStyles().get(StyleIdentifier.HEADING_1));
		            
		// Edit the footnote
		Footnote footnote = (Footnote)docEdited.getChild(NodeType.FOOTNOTE, 0, true);
		footnote.getFirstParagraph().getRuns().get(1).setText("Edited endnote text.");

		// Edit the table
		Table table = (Table)docEdited.getChild(NodeType.TABLE, 0, true);
		table.getFirstRow().getCells().get(1).getFirstParagraph().getRuns().get(0).setText("Edited Cell 2 contents");

		// Edit the textbox
		textBox = (Shape)docEdited.getChild(NodeType.SHAPE, 0, true);
		textBox.getFirstParagraph().getRuns().get(0).setText("Edited textbox contents");

		// Edit the DATE field
		FieldDate fieldDate = (FieldDate)docEdited.getRange().getFields().get(0);
		fieldDate.setUseLunarCalendar(true);

		// Edit the comment
		Comment comment = (Comment)docEdited.getChild(NodeType.COMMENT, 0, true);
		comment.getFirstParagraph().getRuns().get(0).setText("Edited comment.");

		// Edit the header
		docEdited.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getFirstParagraph().getRuns().get(0).setText("Edited header contents.");

		// Apply different comparing options
		CompareOptions compareOptions = new CompareOptions();
		compareOptions.setIgnoreFormatting(false);
		compareOptions.setIgnoreCaseChanges(false);
		compareOptions.setIgnoreComments(false);
		compareOptions.setIgnoreTables(false);
		compareOptions.setIgnoreFields(false);
		compareOptions.setIgnoreFootnotes(false);
		compareOptions.setIgnoreTextboxes(false);
		compareOptions.setIgnoreHeadersAndFooters(false);
		compareOptions.setTarget(ComparisonTargetType.NEW);

		// compare both documents
		docOriginal.compare(docEdited, "John Doe", new Date(), compareOptions);
		docOriginal.save(dataDir + "Document.CompareOptions.docx");
		// ExEnd:AdvancedComparingProperties
	}
}
