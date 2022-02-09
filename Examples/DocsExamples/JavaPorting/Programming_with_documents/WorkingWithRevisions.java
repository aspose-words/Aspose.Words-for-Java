package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Body;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import org.testng.Assert;
import com.aspose.words.ParagraphCollection;
import com.aspose.ms.System.msConsole;
import com.aspose.words.RevisionGroup;
import com.aspose.words.CommentDisplayMode;
import com.aspose.words.ShowInBalloons;
import com.aspose.words.MeasurementUnits;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.Revision;
import com.aspose.words.SaveFormat;
import com.aspose.words.RevisionsView;
import com.aspose.words.NodeType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Node;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.WrapType;
import java.util.ArrayList;


class WorkingWithRevisions extends DocsExamplesBase
{
    @Test
    public void acceptRevisions() throws Exception
    {
        //ExStart:AcceptAllRevisions
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
        doc.startTrackRevisionsInternal("John Doe", new Date());

        // This paragraph is a revision and will have the according "IsInsertRevision" flag set.
        para = body.appendParagraph("Paragraph 4. ");
        Assert.assertTrue(para.isInsertRevision());

        // Get the document's paragraph collection and remove a paragraph.
        ParagraphCollection paragraphs = body.getParagraphs();
        Assert.assertEquals(4, paragraphs.getCount());
        para = paragraphs.get(2);
        para.remove();

        // Since we are tracking revisions, the paragraph still exists in the document, will have the "IsDeleteRevision" set
        // and will be displayed as a revision in Microsoft Word, until we accept or reject all revisions.
        Assert.assertEquals(4, paragraphs.getCount());
        Assert.assertTrue(para.isDeleteRevision());

        // The delete revision paragraph is removed once we accept changes.
        doc.acceptAllRevisions();
        Assert.assertEquals(3, paragraphs.getCount());
        Assert.That(para, Is.Empty);

        // Stopping the tracking of revisions makes this text appear as normal text.
        // Revisions are not counted when the document is changed.
        doc.stopTrackRevisions();

        // Save the document.
        doc.save(getArtifactsDir() + "WorkingWithRevisions.AcceptRevisions.docx");
        //ExEnd:AcceptAllRevisions
    }

    @Test
    public void getRevisionTypes() throws Exception
    {
        //ExStart:GetRevisionTypes
        Document doc = new Document(getMyDir() + "Revisions.docx");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        for (int i = 0; i < paragraphs.getCount(); i++)
        {
            if (paragraphs.get(i).isMoveFromRevision())
                System.out.println("The paragraph {0} has been moved (deleted).",i);
            if (paragraphs.get(i).isMoveToRevision())
                System.out.println("The paragraph {0} has been moved (inserted).",i);
        }
        //ExEnd:GetRevisionTypes
    }

    @Test
    public void getRevisionGroups() throws Exception
    {
        //ExStart:GetRevisionGroups
        Document doc = new Document(getMyDir() + "Revisions.docx");

        for (RevisionGroup group : doc.getRevisions().getGroups())
        {
            System.out.println("{0}, {1}:",group.getAuthor(),group.getRevisionType());
            System.out.println(group.getText());
        }
        //ExEnd:GetRevisionGroups
    }

    @Test
    public void removeCommentsInPdf() throws Exception
    {
        //ExStart:RemoveCommentsInPDF
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Do not render the comments in PDF.
        doc.getLayoutOptions().setCommentDisplayMode(CommentDisplayMode.HIDE);

        doc.save(getArtifactsDir() + "WorkingWithRevisions.RemoveCommentsInPdf.pdf");
        //ExEnd:RemoveCommentsInPDF
    }

    @Test
    public void showRevisionsInBalloons() throws Exception
    {
        //ExStart:ShowRevisionsInBalloons
        //ExStart:SetMeasurementUnit
        //ExStart:SetRevisionBarsPosition
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Renders insert revisions inline, delete and format revisions in balloons.
        doc.getLayoutOptions().getRevisionOptions().setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);
        doc.getLayoutOptions().getRevisionOptions().setMeasurementUnit(MeasurementUnits.INCHES);
        // Renders revision bars on the right side of a page.
        doc.getLayoutOptions().getRevisionOptions().setRevisionBarsPosition(HorizontalAlignment.RIGHT);
        
        doc.save(getArtifactsDir() + "WorkingWithRevisions.ShowRevisionsInBalloons.pdf");
        //ExEnd:SetRevisionBarsPosition
        //ExEnd:SetMeasurementUnit
        //ExEnd:ShowRevisionsInBalloons
    }

    @Test
    public void getRevisionGroupDetails() throws Exception
    {
        //ExStart:GetRevisionGroupDetails
        Document doc = new Document(getMyDir() + "Revisions.docx");

        for (Revision revision : doc.getRevisions())
        {
            String groupText = revision.getGroup() != null
                ? "Revision group text: " + revision.getGroup().getText()
                : "Revision has no group";

            System.out.println("Type: " + revision.getRevisionType());
            System.out.println("Author: " + revision.getAuthor());
            System.out.println("Date: " + revision.getDateTimeInternal());
            System.out.println("Revision text: " + revision.getParentNode().toString(SaveFormat.TEXT));
            System.out.println(groupText);
        }
        //ExEnd:GetRevisionGroupDetails
    }

    @Test
    public void accessRevisedVersion() throws Exception
    {
        //ExStart:AccessRevisedVersion
        Document doc = new Document(getMyDir() + "Revisions.docx");
        doc.updateListLabels();

        // Switch to the revised version of the document.
        doc.setRevisionsView(RevisionsView.FINAL);

        for (Revision revision : doc.getRevisions())
        {
            if (revision.getParentNode().getNodeType() == NodeType.PARAGRAPH)
            {
                Paragraph paragraph = (Paragraph) revision.getParentNode();
                if (paragraph.isListItem())
                {
                    System.out.println(paragraph.getListLabel().getLabelString());
                    System.out.println(paragraph.getListFormat().getListLevel());
                }
            }
        }
        //ExEnd:AccessRevisedVersion
    }

    @Test
    public void moveNodeInTrackedDocument() throws Exception
    {
        //ExStart:MoveNodeInTrackedDocument
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Paragraph 1");
        builder.writeln("Paragraph 2");
        builder.writeln("Paragraph 3");
        builder.writeln("Paragraph 4");
        builder.writeln("Paragraph 5");
        builder.writeln("Paragraph 6");
        Body body = doc.getFirstSection().getBody();
        System.out.println("Paragraph count: {0}",body.getParagraphs().getCount());

        // Start tracking revisions.
        doc.startTrackRevisionsInternal("Author", new DateTime(2020, 12, 23, 14, 0, 0));

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
        System.out.println("Paragraph count: {0}",body.getParagraphs().getCount());
        doc.save(getArtifactsDir() + "WorkingWithRevisions.MoveNodeInTrackedDocument.docx");
        //ExEnd:MoveNodeInTrackedDocument
    }

    @Test
    public void shapeRevision() throws Exception
    {
        //ExStart:ShapeRevision
        Document doc = new Document();

        // Insert an inline shape without tracking revisions.
        Assert.assertFalse(doc.getTrackRevisions());
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
        ArrayList<Shape> shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(2, shapes.size());

        // Remove the first shape.
        shapes.get(0).remove();

        // Because we removed that shape while changes were being tracked, the shape counts as a delete revision.
        Assert.assertEquals(ShapeType.CUBE, shapes.get(0).getShapeType());
        Assert.assertTrue(shapes.get(0).isDeleteRevision());

        // And we inserted another shape while tracking changes, so that shape will count as an insert revision.
        Assert.assertEquals(ShapeType.SUN, shapes.get(1).getShapeType());
        Assert.assertTrue(shapes.get(1).isInsertRevision());

        // The document has one shape that was moved, but shape move revisions will have two instances of that shape.
        // One will be the shape at its arrival destination and the other will be the shape at its original location.
        doc = new Document(getMyDir() + "Revision shape.docx");
        
        shapes = doc.getChildNodes(NodeType.SHAPE, true).<Shape>Cast().ToList();
        Assert.assertEquals(2, shapes.size());

        // This is the move to revision, also the shape at its arrival destination.
        Assert.assertFalse(shapes.get(0).isMoveFromRevision());
        Assert.assertTrue(shapes.get(0).isMoveToRevision());

        // This is the move from revision, which is the shape at its original location.
        Assert.assertTrue(shapes.get(1).isMoveFromRevision());
        Assert.assertFalse(shapes.get(1).isMoveToRevision());
        //ExEnd:ShapeRevision
    }
}
