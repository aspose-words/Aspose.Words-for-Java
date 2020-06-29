package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class TrackChanges {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(TrackChanges.class) + "Document/";

		acceptRevisions(dataDir);
		getRevisionTypes(dataDir);
		getRevisionGroups(dataDir);
		setShowCommentsinPDF(dataDir);
		setShowInBalloons(dataDir);
		getRevisionGroupDetails(dataDir);
		accessRevisedVersion(dataDir);
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
