package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.Date;
import java.util.Iterator;

public class ExRevision extends ApiExampleBase {
    @Test
    public void revisions() throws Exception {
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
        doc.startTrackRevisions("John Doe", new Date());

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
        Assert.assertEquals(revision.getDateTime().getDate(), new Date().getDate());
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
    public void revisionCollection() throws Exception {
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
        System.out.println(MessageFormat.format("{0} revision groups:", revisions.getGroups().getCount()));

        // Iterate over the collection of groups and print the text that the revision concerns.
        Iterator<RevisionGroup> e = revisions.getGroups().iterator();

        while (e.hasNext()) {
            RevisionGroup revisionGroup = e.next();

            System.out.println(MessageFormat.format("\tGroup type \"{0}\", ", revisionGroup.getRevisionType()) +
                    MessageFormat.format("author: {0}, contents: [{1}]", revisionGroup.getAuthor(), revisionGroup.getText().trim()));
        }

        // Each Run that a revision affects gets a corresponding Revision object.
        // The revisions' collection is considerably larger than the condensed form we printed above,
        // depending on how many Runs we have segmented the document into during Microsoft Word editing.
        Assert.assertEquals(11, revisions.getCount()); //ExSkip
        System.out.println("\n{revisions.Count} revisions:");

        Iterator<Revision> e1 = revisions.iterator();

        while (e1.hasNext()) {
            Revision revision = e1.next();

            // A StyleDefinitionChange strictly affects styles and not document nodes. This means the "ParentStyle"
            // property will always be in use, while the ParentNode will always be null.
            // Since all other changes affect nodes, ParentNode will conversely be in use, and ParentStyle will be null.
            if (revision.getRevisionType() == RevisionType.STYLE_DEFINITION_CHANGE) {
                System.out.println(MessageFormat.format("\tRevision type \"{0}\", ", revision.getRevisionType()) +
                        MessageFormat.format("author: {0}, style: [{1}]", revision.getAuthor(), revision.getParentStyle().getName()));
            } else {
                System.out.println(MessageFormat.format("\tRevision type \"{0}\", ", revision.getRevisionType()) +
                        MessageFormat.format("author: {0}, contents: [{1}]", revision.getAuthor(), revision.getParentNode().getText().trim()));
            }
        }

        // Reject all revisions via the collection, reverting the document to its original form.
        revisions.rejectAll();

        Assert.assertEquals(0, revisions.getCount());
        //ExEnd
    }

    @Test
    public void getInfoAboutRevisionsInRevisionGroups() throws Exception {
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

        for (RevisionGroup group : doc.getRevisions().getGroups()) {
            System.out.println(MessageFormat.format("Revision author: {0}; Revision type: {1} \n\tRevision text: {2}", group.getAuthor(), group.getRevisionType(), group.getText()));
        }
        //ExEnd
    }

    @Test
    public void getSpecificRevisionGroup() throws Exception {
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
    public void showRevisionBalloons() throws Exception {
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
    public void revisionOptions() throws Exception {
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
        revisionOptions.setMovedToTextColor(RevisionColor.CLASSIC_BLUE);
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

    //ExStart:RevisionSpecifiedCriteria
    //GistId:66dd22f0854357e394a013b536e2181b
    //ExFor:RevisionCollection.Accept(IRevisionCriteria)
    //ExFor:RevisionCollection.Reject(IRevisionCriteria)
    //ExFor:IRevisionCriteria
    //ExFor:IRevisionCriteria.IsMatch(Revision)
    //ExSummary:Shows how to accept or reject revision based on criteria.
    @Test //ExSkip
    public void revisionSpecifiedCriteria() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("This does not count as a revision. ");

        // To register our edits as revisions, we need to declare an author, and then start tracking them.
        doc.startTrackRevisions("John Doe", new Date());
        builder.write("This is insertion revision #1. ");
        doc.stopTrackRevisions();

        doc.startTrackRevisions("Jane Doe", new Date());
        builder.write("This is insertion revision #2. ");
        // Remove a run "This does not count as a revision.".
        doc.getFirstSection().getBody().getFirstParagraph().getRuns().get(0).remove();
        doc.stopTrackRevisions();

        Assert.assertEquals(3, doc.getRevisions().getCount());
        // We have two revisions from different authors, so we need to accept only one.
        doc.getRevisions().accept(new RevisionCriteria("John Doe", RevisionType.INSERTION));
        Assert.assertEquals(2, doc.getRevisions().getCount());
        // Reject revision with different author name and revision type.
        doc.getRevisions().reject(new RevisionCriteria("Jane Doe", RevisionType.DELETION));
        Assert.assertEquals(1, doc.getRevisions().getCount());

        doc.save(getArtifactsDir() + "Revision.RevisionSpecifiedCriteria.docx");
    }

    /// <summary>
    /// Control when certain revision should be accepted/rejected.
    /// </summary>
    public static class RevisionCriteria implements IRevisionCriteria
    {
        private String AuthorName;
        private int _RevisionType;

        public RevisionCriteria(String authorName, int revisionType)
        {
            AuthorName = authorName;
            _RevisionType = revisionType;
        }

        public boolean isMatch(Revision revision)
        {
            return AuthorName.equals(revision.getAuthor()) && revision.getRevisionType() == _RevisionType;
        }
    }
    //ExEnd:RevisionSpecifiedCriteria

    @Test
    public void trackRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.StartTrackRevisions(String)
        //ExFor:Document.StartTrackRevisions(String, DateTime)
        //ExFor:Document.StopTrackRevisions
        //ExSummary:Shows how to track revisions while editing a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Editing a document usually does not count as a revision until we begin tracking them.
        builder.write("Hello world! ");

        Assert.assertEquals(0, doc.getRevisions().getCount());
        Assert.assertFalse(doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(0).isInsertRevision());

        doc.startTrackRevisions("John Doe");

        builder.write("Hello again! ");

        Assert.assertEquals(1, doc.getRevisions().getCount());
        Assert.assertTrue(doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(1).isInsertRevision());
        Assert.assertEquals("John Doe", doc.getRevisions().get(0).getAuthor());

        // Stop tracking revisions to not count any future edits as revisions.
        doc.stopTrackRevisions();
        builder.write("Hello again! ");

        Assert.assertEquals(1, doc.getRevisions().getCount());
        Assert.assertFalse(doc.getFirstSection().getBody().getParagraphs().get(0).getRuns().get(2).isInsertRevision());

        // Creating revisions gives them a date and time of the operation.
        // We can disable this by passing DateTime.MinValue when we start tracking revisions.
        doc.startTrackRevisions("John Doe", new Date());
        builder.write("Hello again! ");

        Assert.assertEquals(2, doc.getRevisions().getCount());
        Assert.assertEquals("John Doe", doc.getRevisions().get(1).getAuthor());
        Assert.assertEquals(new Date(), doc.getRevisions().get(1).getDateTime());

        // We can accept/reject these revisions programmatically
        // by calling methods such as Document.AcceptAllRevisions, or each revision's Accept method.
        // In Microsoft Word, we can process them manually via "Review" -> "Changes".
        doc.save(getArtifactsDir() + "Document.StartTrackRevisions.docx");
        //ExEnd
    }

    @Test
    public void acceptAllRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.AcceptAllRevisions
        //ExSummary:Shows how to accept all tracking changes in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Edit the document while tracking changes to create a few revisions.
        doc.startTrackRevisions("John Doe");
        builder.write("Hello world! ");
        builder.write("Hello again! ");
        builder.write("This is another revision.");
        doc.stopTrackRevisions();

        Assert.assertEquals(3, doc.getRevisions().getCount());

        // We can iterate through every revision and accept/reject it as a part of our document.
        // If we know we wish to accept every revision, we can do it more straightforwardly so by calling this method.
        doc.acceptAllRevisions();

        Assert.assertEquals(0, doc.getRevisions().getCount());
        Assert.assertEquals("Hello world! Hello again! This is another revision.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void getRevisedPropertiesOfList() throws Exception
    {
        //ExStart
        //ExFor:RevisionsView
        //ExFor:Document.RevisionsView
        //ExSummary:Shows how to switch between the revised and the original view of a document.
        Document doc = new Document(getMyDir() + "Revisions at list levels.docx");
        doc.updateListLabels();

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        Assert.assertEquals("1.", paragraphs.get(0).getListLabel().getLabelString());
        Assert.assertEquals("a.", paragraphs.get(1).getListLabel().getLabelString());
        Assert.assertEquals("", paragraphs.get(2).getListLabel().getLabelString());

        // View the document object as if all the revisions are accepted. Currently supports list labels.
        doc.setRevisionsView(RevisionsView.FINAL);

        Assert.assertEquals("", paragraphs.get(0).getListLabel().getLabelString());
        Assert.assertEquals("1.", paragraphs.get(1).getListLabel().getLabelString());
        Assert.assertEquals("a.", paragraphs.get(2).getListLabel().getLabelString());
        //ExEnd

        doc.setRevisionsView(RevisionsView.ORIGINAL);
        doc.acceptAllRevisions();

        Assert.assertEquals("a.", paragraphs.get(0).getListLabel().getLabelString());
        Assert.assertEquals("", paragraphs.get(1).getListLabel().getLabelString());
        Assert.assertEquals("b.", paragraphs.get(2).getListLabel().getLabelString());
    }

    @Test
    public void compare() throws Exception
    {
        //ExStart
        //ExFor:Document.Compare(Document, String, DateTime)
        //ExFor:RevisionCollection.AcceptAll
        //ExSummary:Shows how to compare documents.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.writeln("This is the original document.");

        Document docEdited = new Document();
        builder = new DocumentBuilder(docEdited);
        builder.writeln("This is the edited document.");

        // Comparing documents with revisions will throw an exception.
        if (docOriginal.getRevisions().getCount() == 0 && docEdited.getRevisions().getCount() == 0)
            docOriginal.compare(docEdited, "authorName", new Date());

        // After the comparison, the original document will gain a new revision
        // for every element that is different in the edited document.
        Assert.assertEquals(2, docOriginal.getRevisions().getCount()); //ExSkip
        for (Revision r : docOriginal.getRevisions())
        {
            System.out.println("Revision type: {r.RevisionType}, on a node of type \"{r.ParentNode.NodeType}\"");
            System.out.println("\tChanged text: \"{r.ParentNode.GetText()}\"");
        }

        // Accepting these revisions will transform the original document into the edited document.
        docOriginal.getRevisions().acceptAll();

        Assert.assertEquals(docOriginal.getText(), docEdited.getText());
        //ExEnd

        docOriginal = DocumentHelper.saveOpen(docOriginal);
        Assert.assertEquals(0, docOriginal.getRevisions().getCount());
    }

    @Test
    public void compareDocumentWithRevisions() throws Exception
    {
        Document doc1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc1);
        builder.writeln("Hello world! This text is not a revision.");

        Document docWithRevision = new Document();
        builder = new DocumentBuilder(docWithRevision);

        docWithRevision.startTrackRevisions("John Doe");
        builder.writeln("This is a revision.");

        Assert.assertThrows(IllegalStateException.class, () -> docWithRevision.compare(doc1, "John Doe", new Date()));
    }

    @Test
    public void compareOptions() throws Exception
    {
        //ExStart
        //ExFor:CompareOptions
        //ExFor:CompareOptions.IgnoreFormatting
        //ExFor:CompareOptions.IgnoreCaseChanges
        //ExFor:CompareOptions.IgnoreComments
        //ExFor:CompareOptions.IgnoreTables
        //ExFor:CompareOptions.IgnoreFields
        //ExFor:CompareOptions.IgnoreFootnotes
        //ExFor:CompareOptions.IgnoreTextboxes
        //ExFor:CompareOptions.IgnoreHeadersAndFooters
        //ExFor:CompareOptions.Target
        //ExFor:ComparisonTargetType
        //ExFor:Document.Compare(Document, String, DateTime, CompareOptions)
        //ExSummary:Shows how to filter specific types of document elements when making a comparison.
        // Create the original document and populate it with various kinds of elements.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);

        // Paragraph text referenced with an endnote:
        builder.writeln("Hello world! This is the first paragraph.");
        builder.insertFootnote(FootnoteType.ENDNOTE, "Original endnote text.");

        // Table:
        builder.startTable();
        builder.insertCell();
        builder.write("Original cell 1 text");
        builder.insertCell();
        builder.write("Original cell 2 text");
        builder.endTable();

        // Textbox:
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 150.0, 20.0);
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("Original textbox contents");

        // DATE field:
        builder.moveTo(docOriginal.getFirstSection().getBody().appendParagraph(""));
        builder.insertField(" DATE ");

        // Comment:
        Comment newComment = new Comment(docOriginal, "John Doe", "J.D.", new Date());
        newComment.setText("Original comment.");
        builder.getCurrentParagraph().appendChild(newComment);

        // Header:
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("Original header contents.");

        // Create a clone of our document and perform a quick edit on each of the cloned document's elements.
        Document docEdited = (Document)docOriginal.deepClone(true);
        Paragraph firstParagraph = docEdited.getFirstSection().getBody().getFirstParagraph();

        firstParagraph.getRuns().get(0).setText("hello world! this is the first paragraph, after editing.");
        firstParagraph.getParagraphFormat().setStyle(docEdited.getStyles().getByStyleIdentifier(StyleIdentifier.HEADING_1));
        ((Footnote)docEdited.getChild(NodeType.FOOTNOTE, 0, true)).getFirstParagraph().getRuns().get(1).setText("Edited endnote text.");
        ((Table)docEdited.getChild(NodeType.TABLE, 0, true)).getFirstRow().getCells().get(1).getFirstParagraph().getRuns().get(0).setText("Edited Cell 2 contents");
        ((Shape)docEdited.getChild(NodeType.SHAPE, 0, true)).getFirstParagraph().getRuns().get(0).setText("Edited textbox contents");
        ((FieldDate)docEdited.getRange().getFields().get(0)).setUseLunarCalendar(true);
        ((Comment)docEdited.getChild(NodeType.COMMENT, 0, true)).getFirstParagraph().getRuns().get(0).setText("Edited comment.");
        docEdited.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getFirstParagraph().getRuns().get(0).setText("Edited header contents.");

        // Comparing documents creates a revision for every edit in the edited document.
        // A CompareOptions object has a series of flags that can suppress revisions
        // on each respective type of element, effectively ignoring their change.
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

        docOriginal.compare(docEdited, "John Doe", new Date(), compareOptions);
        docOriginal.save(getArtifactsDir() + "Document.CompareOptions.docx");
        //ExEnd

        docOriginal = new Document(getArtifactsDir() + "Document.CompareOptions.docx");

        TestUtil.verifyFootnote(FootnoteType.ENDNOTE, true, "",
                "OriginalEdited endnote text.", (Footnote)docOriginal.getChild(NodeType.FOOTNOTE, 0, true));
    }

    @Test (dataProvider = "ignoreDmlUniqueIdDataProvider")
    public void ignoreDmlUniqueId(boolean isIgnoreDmlUniqueId) throws Exception
    {
        //ExStart
        //ExFor:CompareOptions.AdvancedOptions
        //ExFor:AdvancedCompareOptions.IgnoreDmlUniqueId
        //ExFor:CompareOptions.IgnoreDmlUniqueId
        //ExSummary:Shows how to compare documents ignoring DML unique ID.
        Document docA = new Document(getMyDir() + "DML unique ID original.docx");
        Document docB = new Document(getMyDir() + "DML unique ID compare.docx");

        // By default, Aspose.Words do not ignore DML's unique ID, and the revisions count was 2.
        // If we are ignoring DML's unique ID, and revisions count were 0.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.getAdvancedOptions().setIgnoreDmlUniqueId(isIgnoreDmlUniqueId);

        docA.compare(docB, "Aspose.Words", new Date(), compareOptions);

        Assert.assertEquals(isIgnoreDmlUniqueId ? 0 : 2, docA.getRevisions().getCount());
        //ExEnd
    }

    //JAVA-added data provider for test method
    @DataProvider(name = "ignoreDmlUniqueIdDataProvider")
    public static Object[][] ignoreDmlUniqueIdDataProvider() throws Exception
    {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test
    public void layoutOptionsRevisions() throws Exception
    {
        //ExStart
        //ExFor:Document.LayoutOptions
        //ExFor:LayoutOptions
        //ExFor:LayoutOptions.RevisionOptions
        //ExFor:RevisionColor
        //ExFor:RevisionOptions
        //ExFor:RevisionOptions.InsertedTextColor
        //ExFor:RevisionOptions.ShowRevisionBars
        //ExFor:RevisionOptions.RevisionBarsPosition
        //ExSummary:Shows how to alter the appearance of revisions in a rendered output document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a revision, then change the color of all revisions to green.
        builder.writeln("This is not a revision.");
        doc.startTrackRevisions("John Doe", new Date());
        Assert.assertEquals(RevisionColor.BY_AUTHOR, doc.getLayoutOptions().getRevisionOptions().getInsertedTextColor()); //ExSkip
        Assert.assertTrue(doc.getLayoutOptions().getRevisionOptions().getShowRevisionBars()); //ExSkip
        builder.writeln("This is a revision.");
        doc.stopTrackRevisions();
        builder.writeln("This is not a revision.");

        // Remove the bar that appears to the left of every revised line.
        doc.getLayoutOptions().getRevisionOptions().setInsertedTextColor(RevisionColor.BRIGHT_GREEN);
        doc.getLayoutOptions().getRevisionOptions().setShowRevisionBars(false);
        doc.getLayoutOptions().getRevisionOptions().setRevisionBarsPosition(HorizontalAlignment.RIGHT);

        doc.save(getArtifactsDir() + "Document.LayoutOptionsRevisions.pdf");
        //ExEnd
    }

    @Test (dataProvider = "granularityCompareOptionDataProvider")
    public void granularityCompareOption(/*Granularity*/int granularity) throws Exception
    {
        //ExStart
        //ExFor:CompareOptions.Granularity
        //ExFor:Granularity
        //ExSummary:Shows to specify a granularity while comparing documents.
        Document docA = new Document();
        DocumentBuilder builderA = new DocumentBuilder(docA);
        builderA.writeln("Alpha Lorem ipsum dolor sit amet, consectetur adipiscing elit");

        Document docB = new Document();
        DocumentBuilder builderB = new DocumentBuilder(docB);
        builderB.writeln("Lorems ipsum dolor sit amet consectetur - \"adipiscing\" elit");

        // Specify whether changes are tracking
        // by character ('Granularity.CharLevel'), or by word ('Granularity.WordLevel').
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setGranularity(granularity);

        docA.compare(docB, "author", new Date(), compareOptions);

        // The first document's collection of revision groups contains all the differences between documents.
        RevisionGroupCollection groups = docA.getRevisions().getGroups();
        Assert.assertEquals(5, groups.getCount());
        //ExEnd

        if (granularity == Granularity.CHAR_LEVEL)
        {
            Assert.assertEquals(RevisionType.DELETION, groups.get(0).getRevisionType());
            Assert.assertEquals("Alpha ", groups.get(0).getText());

            Assert.assertEquals(RevisionType.DELETION, groups.get(1).getRevisionType());
            Assert.assertEquals(",", groups.get(1).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(2).getRevisionType());
            Assert.assertEquals("s", groups.get(2).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(3).getRevisionType());
            Assert.assertEquals("- \"", groups.get(3).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(4).getRevisionType());
            Assert.assertEquals("\"", groups.get(4).getText());
        }
        else
        {
            Assert.assertEquals(RevisionType.DELETION, groups.get(0).getRevisionType());
            Assert.assertEquals("Alpha Lorem", groups.get(0).getText());

            Assert.assertEquals(RevisionType.DELETION, groups.get(1).getRevisionType());
            Assert.assertEquals(",", groups.get(1).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(2).getRevisionType());
            Assert.assertEquals("Lorems", groups.get(2).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(3).getRevisionType());
            Assert.assertEquals("- \"", groups.get(3).getText());

            Assert.assertEquals(RevisionType.INSERTION, groups.get(4).getRevisionType());
            Assert.assertEquals("\"", groups.get(4).getText());
        }
    }

    @DataProvider(name = "granularityCompareOptionDataProvider")
    public static Object[][] granularityCompareOptionDataProvider() {
        return new Object[][]
                {
                        {Granularity.CHAR_LEVEL},
                        {Granularity.WORD_LEVEL},
                };
    }

    @Test
    public void ignoreStoreItemId() throws Exception
    {
        //ExStart:IgnoreStoreItemId
        //GistId:a76df4b18bee76d169e55cdf6af8129c
        //ExFor:AdvancedCompareOptions
        //ExFor:AdvancedCompareOptions.IgnoreStoreItemId
        //ExSummary:Shows how to compare SDT with same content but different store item id.
        Document docA = new Document(getMyDir() + "Document with SDT 1.docx");
        Document docB = new Document(getMyDir() + "Document with SDT 2.docx");

        // Configure options to compare SDT with same content but different store item id.
        CompareOptions compareOptions = new CompareOptions();
        compareOptions.getAdvancedOptions().setIgnoreStoreItemId(false);

        docA.compare(docB, "user", new Date(), compareOptions);
        Assert.assertEquals(8, docA.getRevisions().getCount());

        compareOptions.getAdvancedOptions().setIgnoreStoreItemId(true);

        docA.getRevisions().rejectAll();
        docA.compare(docB, "user", new Date(), compareOptions);
        Assert.assertEquals(0, docA.getRevisions().getCount());
        //ExEnd:IgnoreStoreItemId
    }

    @Test
    public void revisionCellColor() throws Exception
    {
        //ExStart:RevisionCellColor
        //GistId:366eb64fd56dec3c2eaa40410e594182
        //ExFor:RevisionOptions.InsertCellColor
        //ExFor:RevisionOptions.DeleteCellColor
        //ExSummary:Shows how to work with insert/delete cell revision color.
        Document doc = new Document(getMyDir() + "Cell revisions.docx");

        doc.getLayoutOptions().getRevisionOptions().setInsertCellColor(RevisionColor.BLUE);
        doc.getLayoutOptions().getRevisionOptions().setDeleteCellColor(RevisionColor.DARK_RED);

        doc.save(getArtifactsDir() + "Revision.RevisionCellColor.pdf");
        //ExEnd:RevisionCellColor
    }
}

