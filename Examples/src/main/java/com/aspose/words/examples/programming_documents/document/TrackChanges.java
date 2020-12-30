package com.aspose.words.examples.programming_documents.document;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class TrackChanges {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(TrackChanges.class) + "Document/";

		WorkWithTrackChanges(dataDir);
		GenerateRevisionsWhenMovingNode(dataDir);
		ApplyDifferentPropertiesWithRevisions(dataDir);
		acceptRevisions(dataDir);
		getRevisionTypes(dataDir);
		getRevisionGroups(dataDir);
		setShowCommentsinPDF(dataDir);
		setShowInBalloons(dataDir);
		getRevisionGroupDetails(dataDir);
		accessRevisedVersion(dataDir);
	}

	private static void WorkWithTrackChanges(String dataDir) throws Exception {
		// ExStart:WorkWithTrackChanges
		Document doc = new Document();
		Body body = doc.getFirstSection().getBody();
		Paragraph para = body.getFirstParagraph();

		// Add text to the first paragraph, then add two more paragraphs.
		para.appendChild(new Run(doc, "Paragraph 1. "));
		body.appendParagraph("Paragraph 2. ");
		body.appendParagraph("Paragraph 3. ");

		// We have three paragraphs, none of which registered as any type of revision
		// If we add/remove any content in the document while tracking revisions,
		// they will be displayed as such in the document and can be accepted/rejected.
		doc.startTrackRevisions("John Doe", new Date());

		// This paragraph is a revision and will have the according "IsInsertRevision" flag set.
		para = body.appendParagraph("Paragraph 4. ");
		if(para.isInsertRevision())
			System.out.println("isInsertRevision:" + para.isInsertRevision());

		// Get the document's paragraph collection and remove a paragraph.
		ParagraphCollection paragraphs = body.getParagraphs();
		if(4 == paragraphs.getCount())
			System.out.println("count:" + paragraphs.getCount());

		para = paragraphs.get(2);
		para.remove();

		// Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
		// and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions.
		if(4 == paragraphs.getCount())
			System.out.println("count:" + paragraphs.getCount());

		if(para.isDeleteRevision())
			System.out.println("isDeleteRevision:" + para.isDeleteRevision());

		// The delete revision paragraph is removed once we accept changes.
		doc.acceptAllRevisions();
		if(3 == paragraphs.getCount())
			System.out.println("count:" + paragraphs.getCount());

		// Stopping the tracking of revisions makes this text appear as normal text.
		// Revisions are not counted when the document is changed.
		doc.stopTrackRevisions();

		// Save the document.
		doc.save(dataDir + "Document.Revisions.docx");
		// ExEnd:WorkWithTrackChanges
	}
	
	private static void GenerateRevisionsWhenMovingNode(String dataDir) throws Exception {
		// ExStart:GenerateRevisionsWhenMovingNode
		// Generate document contents.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.writeln("Paragraph 1");
		builder.writeln("Paragraph 2");
		builder.writeln("Paragraph 3");
		builder.writeln("Paragraph 4");
		builder.writeln("Paragraph 5");
		builder.writeln("Paragraph 6");
		Body body = doc.getFirstSection().getBody();
		System.out.println("Paragraph count:" + body.getParagraphs().getCount());

		// Start tracking revisions.
		doc.startTrackRevisions("Author", new Date());

		// Generate revisions when moving a node from one location to another.
		Node node = body.getParagraphs().get(3);
		Node endNode = body.getParagraphs().get(5).getNextSibling();
		Node referenceNode = body.getParagraphs().get(0);
		while (node != endNode)
		{
			Node nextNode = node.getNextSibling();
			body.insertBefore(node, referenceNode);
			node = nextNode;
		}

		// Stop the process of tracking revisions.
		doc.stopTrackRevisions();

		// There are 3 additional paragraphs in the move-from range.
		System.out.println("Paragraph count: " + body.getParagraphs().getCount());
		doc.save(dataDir + "out.docx");
		// ExEnd:GenerateRevisionsWhenMovingNode
	}

	private static void ApplyDifferentPropertiesWithRevisions(String dataDir) throws Exception {
		// ExStart:ApplyDifferentPropertiesWithRevisions
		// Open a blank document.
		Document doc = new Document();

		// Insert an inline shape without tracking revisions.
		Shape shape = new Shape(doc, ShapeType.CUBE);
		shape.setWrapType(WrapType.INLINE);
		shape.setWidth(100.0);
		shape.setHeight(100.0);
		doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

		// Start tracking revisions and then insert another shape.
		doc.startTrackRevisions("John Doe");
		shape = new Shape(doc, ShapeType.SUN);
		shape.setWrapType(WrapType.INLINE);
		shape.setWidth(100.0);
		shape.setHeight(100.0);
		doc.getFirstSection().getBody().getFirstParagraph().appendChild(shape);

		// Get the document's shape collection which includes just the two shapes we added.
		Node[] shapes = doc.getChildNodes(NodeType.SHAPE, true).toArray();
		if(2 == shapes.length)
			System.out.println("Shapes Count:" + shapes.length);

		// Remove the first shape.
		shapes[0].remove();

		Shape sh = (Shape) shapes[0];
		// Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
		if(ShapeType.CUBE == sh.getShapeType())
			System.out.println("Shape is CUBE");
		
		if(sh.isDeleteRevision())
			System.out.println("isDeleteRevision:" + sh.isDeleteRevision());

		// And we inserted another shape while tracking changes, so that shape will count as an insert revision.
		sh = (Shape) shapes[1];
		if(ShapeType.SUN == sh.getShapeType())
			System.out.println("Shape is SUN");
		
		if(sh.isInsertRevision())
			System.out.println("IsInsertRevision:" + sh.isInsertRevision());

		// The document has one shape that was moved, but shape move revisions will have two instances of that shape.
		// One will be the shape at its arrival destination and the other will be the shape at its original location.
		Node[] nc = doc.getChildNodes(NodeType.SHAPE, true).toArray();
		if(2 == nc.length)
			System.out.println("Shapes Count:" + nc.length);

		Shape mvr = (Shape) nc[0];
		// This is the move to revision, also the shape at its arrival destination.
		if(mvr.isMoveFromRevision())
			System.out.println("isMoveFromRevision:" + mvr.isMoveFromRevision());
		
		if(mvr.isMoveToRevision())
			System.out.println("isMoveToRevision:" + mvr.isMoveToRevision());
		
		mvr = (Shape) nc[1];
		// This is the move from revision, which is the shape at its original location.
		if(mvr.isMoveFromRevision())
			System.out.println("isMoveFromRevision:" + mvr.isMoveFromRevision());
		
		if(mvr.isMoveToRevision())
			System.out.println("isMoveToRevision:" + mvr.isMoveToRevision());
		// ExEnd:ApplyDifferentPropertiesWithRevisions
	}
	
	private static void acceptRevisions(String dataDir) throws Exception {
		// ExStart:AcceptAllRevisions
		Document doc = new Document(dataDir + "Document.doc");

		// Start tracking and make some revisions.
		doc.startTrackRevisions("Author");
		doc.getFirstSection().getBody().appendParagraph("Hello world!");

		// Revisions will now show up as normal text in the output document.
		doc.acceptAllRevisions();

		dataDir = dataDir + "Document.AcceptedRevisions_out.doc";
		doc.save(dataDir);
		// ExEnd:AcceptAllRevisions
		System.out.println("\nAll revisions accepted.\nFile saved at " + dataDir);
	}

	private static void getRevisionTypes(String dataDir) throws Exception {
		// ExStart:GetRevisionTypes
		Document doc = new Document(dataDir + "Revisions.docx");

		ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
		for (int i = 0; i < paragraphs.getCount(); i++) {
			if (paragraphs.get(i).isMoveFromRevision())
				System.out.println("The paragraph " + i + " has been moved (deleted).");
			if (paragraphs.get(i).isMoveToRevision())
				System.out.println("The paragraph " + i + " has been moved (inserted).");
		}
		// ExEnd:GetRevisionTypes
	}

	private static void getRevisionGroups(String dataDir) throws Exception {
		// ExStart:GetRevisionGroups
		Document doc = new Document(dataDir + "Revisions.docx");

		for (RevisionGroup group : (Iterable<RevisionGroup>) doc.getRevisions().getGroups()) {
			System.out.println(group.getAuthor() + ", " + RevisionType.getName(group.getRevisionType()) + ": ");
			System.out.println(group.getText());
		}
		// ExEnd:GetRevisionGroups
	}

	private static void setShowCommentsinPDF(String dataDir) throws Exception {
		// ExStart:SetShowCommentsinPDF
		Document doc = new Document(dataDir + "Revisions.docx");

		// Do not render the comments in PDF
		doc.getLayoutOptions().setShowComments(false);
		doc.save(dataDir + "RemoveCommentsinPDF_out.pdf");
		// ExEnd:SetShowCommentsinPDF
		System.out.println("\nFile saved at " + dataDir);
	}

	private static void setShowInBalloons(String dataDir) throws Exception {
		// ExStart:SetShowInBalloons
		Document doc = new Document(dataDir + "Revisions.docx");

		// Get the RevisionOptions object that controls the appearance of revisions
		RevisionOptions revisionOptions = doc.getLayoutOptions().getRevisionOptions();

		// Show deletion revisions in balloon
		revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);

		doc.save(dataDir + "Revisions.SetShowInBalloons_out.pdf");
		// ExEnd:SetShowInBalloons
		System.out.println("\nFile saved at " + dataDir);
	}

	private static void getRevisionGroupDetails(String dataDir) throws Exception {
		// ExStart:GetRevisionGroupDetails
		Document doc = new Document(dataDir + "TestFormatDescription.docx");

		for (Revision revision : (Iterable<Revision>) doc.getRevisions()) {
			String groupText = revision.getGroup() != null ? "Revision group text: " + revision.getGroup().getText()
					: "Revision has no group";

			System.out.println("Type: " + revision.getRevisionType());
			System.out.println("Author: " + revision.getAuthor());
			System.out.println("Date: " + revision.getDateTime());
			System.out.println("Revision text: " + revision.getParentNode().toString(SaveFormat.TEXT));
			System.out.println(groupText);
		}
		// ExEnd:GetRevisionGroupDetails
	}

	private static void accessRevisedVersion(String dataDir) throws Exception {
		// ExStart:AccessRevisedVersion
		Document doc = new Document(dataDir + "Test.docx");
		doc.updateListLabels();

		// Switch to the revised version of the document.
		doc.setRevisionsView(RevisionsView.FINAL);

		for (Revision revision : (Iterable<Revision>) doc.getRevisions()) {
			if (revision.getParentNode().getNodeType() == NodeType.PARAGRAPH) {
				Paragraph paragraph = (Paragraph) revision.getParentNode();
				if (paragraph.isListItem()) {
					// Print revised version of LabelString and ListLevel.
					System.out.println(paragraph.getListLabel().getLabelString());
					System.out.println(paragraph.getListFormat().getListLevel());
				}
			}
		}
		// ExEnd:AccessRevisedVersion
	}
}
