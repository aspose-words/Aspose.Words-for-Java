// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import java.util.Date;
import com.aspose.ms.System.DateTime;
import com.aspose.words.Revision;
import com.aspose.words.RevisionType;
import com.aspose.words.Node;
import com.aspose.words.RevisionCollection;
import com.aspose.ms.System.msConsole;
import java.util.Iterator;
import com.aspose.words.RevisionGroup;
import com.aspose.words.ShowInBalloons;
import com.aspose.words.RevisionOptions;
import com.aspose.words.RevisionColor;
import com.aspose.words.RevisionTextEffect;


@Test
class ExRevision !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void revisions() throws Exception
    {
        //ExStart
        //ExFor:Revision
        //ExFor:Revision.Accept
        //ExFor:Revision.Author
        //ExFor:Revision.DateTime
        //ExFor:Revision.Group
        //ExFor:Revision.Reject
        //ExFor:Revision.RevisionType
        //ExFor:RevisionCollection
        //ExFor:RevisionCollection.Item(Int32)
        //ExFor:RevisionCollection.Count
        //ExFor:RevisionType
        //ExFor:Document.HasRevisions
        //ExFor:Document.TrackRevisions
        //ExFor:Document.Revisions
        //ExSummary:Shows how to work with revisions in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normal editing of the document does not count as a revision.
        builder.write("This does not count as a revision. ");

        Assert.assertFalse(doc.hasRevisions());

        // To register our edits as revisions, we need to declare an author, and then start tracking them.
        doc.startTrackRevisionsInternal("John Doe", new Date());

        builder.write("This is revision #1. ");

        Assert.assertTrue(doc.hasRevisions());
        Assert.assertEquals(1, doc.getRevisions().getCount());

        // This flag corresponds to the "Review" -> "Tracking" -> "Track Changes" option in Microsoft Word.
        // The "StartTrackRevisions" method does not affect its value,
        // and the document is tracking revisions programmatically despite it having a value of "false".
        // If we open this document using Microsoft Word, it will not be tracking revisions.
        Assert.assertFalse(doc.getTrackRevisions());

        // We have added text using the document builder, so the first revision is an insertion-type revision.
        Revision revision = doc.getRevisions().get(0);
        Assert.assertEquals("John Doe", revision.getAuthor());
        Assert.assertEquals("This is revision #1. ", revision.getParentNode().getText());
        Assert.assertEquals(RevisionType.INSERTION, revision.getRevisionType());
        Assert.assertEquals(revision.getDateTimeInternal().getDate(), new Date().getDate());
        Assert.assertEquals(doc.getRevisions().getGroups().get(0), revision.getGroup());

        // Remove a run to create a deletion-type revision.
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).remove();

        // Adding a new revision places it at the beginning of the revision collection.
        Assert.assertEquals(RevisionType.DELETION, doc.getRevisions().get(0).getRevisionType());
        Assert.assertEquals(2, doc.getRevisions().getCount());

        // Insert revisions show up in the document body even before we accept/reject the revision.
        // Rejecting the revision will remove its nodes from the body. Conversely, nodes that make up delete revisions
        // also linger in the document until we accept the revision.
        Assert.assertEquals("This does not count as a revision. This is revision #1.", doc.getText().trim());

        // Accepting the delete revision will remove its parent node from the paragraph text
        // and then remove the collection's revision itself.
        doc.getRevisions().get(0).accept();

        Assert.assertEquals(1, doc.getRevisions().getCount());
        Assert.assertEquals("This is revision #1.", doc.getText().trim());

        builder.writeln("");
        builder.write("This is revision #2.");

        // Now move the node to create a moving revision type.
        Node node = doc.getFirstSection().getBody().getParagraphs().get(1);
        Node endNode = doc.getFirstSection().getBody().getParagraphs().get(1).getNextSibling();
        Node referenceNode = doc.getFirstSection().getBody().getParagraphs().get(0);

        while (node != endNode)
        {
            Node nextNode = node.getNextSibling();
            doc.getFirstSection().getBody().insertBefore(node, referenceNode);
            node = nextNode;
        }

        Assert.assertEquals(RevisionType.MOVING, doc.getRevisions().get(0).getRevisionType());
        Assert.assertEquals(8, doc.getRevisions().getCount());
        Assert.assertEquals("This is revision #2.\rThis is revision #1. \rThis is revision #2.", doc.getText().trim());

        // The moving revision is now at index 1. Reject the revision to discard its contents.
        doc.getRevisions().get(1).reject();

        Assert.assertEquals(6, doc.getRevisions().getCount());
        Assert.assertEquals("This is revision #1. \rThis is revision #2.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void revisionCollection() throws Exception
    {
        //ExStart
        //ExFor:Revision.ParentStyle
        //ExFor:RevisionCollection.GetEnumerator
        //ExFor:RevisionCollection.Groups
        //ExFor:RevisionCollection.RejectAll
        //ExFor:RevisionGroupCollection.GetEnumerator
        //ExSummary:Shows how to work with a document's collection of revisions.
        Document doc = new Document(getMyDir() + "Revisions.docx");
        RevisionCollection revisions = doc.getRevisions();

        // This collection itself has a collection of revision groups.
        // Each group is a sequence of adjacent revisions.
        Assert.assertEquals(7, revisions.getGroups().getCount()); //ExSkip
        System.out.println("{revisions.Groups.Count} revision groups:");

        // Iterate over the collection of groups and print the text that the revision concerns.
        Iterator<RevisionGroup> e = revisions.getGroups().iterator();
        try /*JAVA: was using*/
        {
            while (e.hasNext())
            {
                System.out.println("\tGroup type \"{e.Current.RevisionType}\", " +
                                      $"author: {e.Current.Author}, contents: [{e.Current.Text.Trim()}]");
            }
        }
        finally { if (e != null) e.close(); }

        // Each Run that a revision affects gets a corresponding Revision object.
        // The revisions' collection is considerably larger than the condensed form we printed above,
        // depending on how many Runs we have segmented the document into during Microsoft Word editing.
        Assert.assertEquals(11, revisions.getCount()); //ExSkip
        System.out.println("\n{revisions.Count} revisions:");

        Iterator<Revision> e1 = revisions.iterator();
        try /*JAVA: was using*/
        {
            while (e1.hasNext())
            {
                // A StyleDefinitionChange strictly affects styles and not document nodes. This means the "ParentStyle"
                // property will always be in use, while the ParentNode will always be null.
                // Since all other changes affect nodes, ParentNode will conversely be in use, and ParentStyle will be null.
                if (e1.next().getRevisionType() == RevisionType.STYLE_DEFINITION_CHANGE)
                {
                    System.out.println("\tRevision type \"{e.Current.RevisionType}\", " +
                                          $"author: {e.Current.Author}, style: [{e.Current.ParentStyle.Name}]");
                }
                else
                {
                    System.out.println("\tRevision type \"{e.Current.RevisionType}\", " +
                                          $"author: {e.Current.Author}, contents: [{e.Current.ParentNode.GetText().Trim()}]");
                }
            }
        }
        finally { if (e1 != null) e1.close(); }

        // Reject all revisions via the collection, reverting the document to its original form.
        revisions.rejectAll();

        Assert.assertEquals(0, revisions.getCount());
        //ExEnd
    }

    @Test
    public void getInfoAboutRevisionsInRevisionGroups() throws Exception
    {
        //ExStart
        //ExFor:RevisionGroup
        //ExFor:RevisionGroup.Author
        //ExFor:RevisionGroup.RevisionType
        //ExFor:RevisionGroup.Text
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Count
        //ExSummary:Shows how to print info about a group of revisions in a document.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        Assert.assertEquals(7, doc.getRevisions().getGroups().getCount());

        for (RevisionGroup group : doc.getRevisions().getGroups())
        {
            System.out.println("Revision author: {group.Author}; Revision type: {group.RevisionType} \n\tRevision text: {group.Text}");
        }
        //ExEnd
    }

    @Test
    public void getSpecificRevisionGroup() throws Exception
    {
        //ExStart
        //ExFor:RevisionGroupCollection
        //ExFor:RevisionGroupCollection.Item(Int32)
        //ExSummary:Shows how to get a group of revisions in a document.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        RevisionGroup revisionGroup = doc.getRevisions().getGroups().get(0);
        //ExEnd

        Assert.assertEquals(RevisionType.DELETION, revisionGroup.getRevisionType());
        Assert.assertEquals("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ",
            revisionGroup.getText());
    }

    @Test
    public void showRevisionBalloons() throws Exception
    {
        //ExStart
        //ExFor:RevisionOptions.ShowInBalloons
        //ExSummary:Shows how to display revisions in balloons.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // By default, text that is a revision has a different color to differentiate it from the other non-revision text.
        // Set a revision option to show more details about each revision in a balloon on the page's right margin.
        doc.getLayoutOptions().getRevisionOptions().setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);
        doc.save(getArtifactsDir() + "Revision.ShowRevisionBalloons.pdf");
        //ExEnd
    }

    @Test
    public void revisionOptions() throws Exception
    {
        //ExStart
        //ExFor:ShowInBalloons
        //ExFor:RevisionOptions.ShowInBalloons
        //ExFor:RevisionOptions.CommentColor
        //ExFor:RevisionOptions.DeletedTextColor
        //ExFor:RevisionOptions.DeletedTextEffect
        //ExFor:RevisionOptions.InsertedTextEffect
        //ExFor:RevisionOptions.MovedFromTextColor
        //ExFor:RevisionOptions.MovedFromTextEffect
        //ExFor:RevisionOptions.MovedToTextColor
        //ExFor:RevisionOptions.MovedToTextEffect
        //ExFor:RevisionOptions.RevisedPropertiesColor
        //ExFor:RevisionOptions.RevisedPropertiesEffect
        //ExFor:RevisionOptions.RevisionBarsColor
        //ExFor:RevisionOptions.RevisionBarsWidth
        //ExFor:RevisionOptions.ShowOriginalRevision
        //ExFor:RevisionOptions.ShowRevisionMarks
        //ExFor:RevisionTextEffect
        //ExSummary:Shows how to modify the appearance of revisions.
        Document doc = new Document(getMyDir() + "Revisions.docx");

        // Get the RevisionOptions object that controls the appearance of revisions.
        RevisionOptions revisionOptions = doc.getLayoutOptions().getRevisionOptions();

        // Render insertion revisions in green and italic.
        revisionOptions.setInsertedTextColor(RevisionColor.GREEN);
        revisionOptions.setInsertedTextEffect(RevisionTextEffect.ITALIC);

        // Render deletion revisions in red and bold.
        revisionOptions.setDeletedTextColor(RevisionColor.RED);
        revisionOptions.setDeletedTextEffect(RevisionTextEffect.BOLD);

        // The same text will appear twice in a movement revision:
        // once at the departure point and once at the arrival destination.
        // Render the text at the moved-from revision yellow with a double strike through
        // and double-underlined blue at the moved-to revision.
        revisionOptions.setMovedFromTextColor(RevisionColor.YELLOW);
        revisionOptions.setMovedFromTextEffect(RevisionTextEffect.DOUBLE_STRIKE_THROUGH);
        revisionOptions.setMovedToTextColor(RevisionColor.BLUE);
        revisionOptions.setMovedFromTextEffect(RevisionTextEffect.DOUBLE_UNDERLINE);

        // Render format revisions in dark red and bold.
        revisionOptions.setRevisedPropertiesColor(RevisionColor.DARK_RED);
        revisionOptions.setRevisedPropertiesEffect(RevisionTextEffect.BOLD);

        // Place a thick dark blue bar on the left side of the page next to lines affected by revisions.
        revisionOptions.setRevisionBarsColor(RevisionColor.DARK_BLUE);
        revisionOptions.setRevisionBarsWidth(15.0f);

        // Show revision marks and original text.
        revisionOptions.setShowOriginalRevision(true);
        revisionOptions.setShowRevisionMarks(true);

        // Get movement, deletion, formatting revisions, and comments to show up in green balloons
        // on the right side of the page.
        revisionOptions.setShowInBalloons(ShowInBalloons.FORMAT);
        revisionOptions.setCommentColor(RevisionColor.BRIGHT_GREEN);

        // These features are only applicable to formats such as .pdf or .jpg.
        doc.save(getArtifactsDir() + "Revision.RevisionOptions.pdf");
        //ExEnd
    }
}

