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
import com.aspose.words.ProtectionType;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.EditableRangeStart;
import com.aspose.words.EditableRangeEnd;
import com.aspose.words.EditableRange;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.words.EditorType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.ms.System.msString;
import com.aspose.words.Run;
import com.aspose.words.NodeCollection;


@Test
class ExEditableRange !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void createAndRemove() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.StartEditableRange
        //ExFor:DocumentBuilder.EndEditableRange
        //ExFor:EditableRange
        //ExFor:EditableRange.EditableRangeEnd
        //ExFor:EditableRange.EditableRangeStart
        //ExFor:EditableRange.Id
        //ExFor:EditableRange.Remove
        //ExFor:EditableRangeEnd.EditableRangeStart
        //ExFor:EditableRangeEnd.Id
        //ExFor:EditableRangeEnd.NodeType
        //ExFor:EditableRangeStart.EditableRange
        //ExFor:EditableRangeStart.Id
        //ExFor:EditableRangeStart.NodeType
        //ExSummary:Shows how to work with an editable range.
        Document doc = new Document();
        doc.protect(ProtectionType.READ_ONLY, "MyPassword");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
                        " we cannot edit this paragraph without the password.");

        // Editable ranges allow us to leave parts of protected documents open for editing.
        EditableRangeStart editableRangeStart = builder.startEditableRange();
        builder.writeln("This paragraph is inside an editable range, and can be edited.");
        EditableRangeEnd editableRangeEnd = builder.endEditableRange();

        // A well-formed editable range has a start node, and end node.
        // These nodes have matching IDs and encompass editable nodes.
        EditableRange editableRange = editableRangeStart.getEditableRange();

        Assert.assertEquals(editableRangeStart.getId(), editableRange.getId());
        Assert.assertEquals(editableRangeEnd.getId(), editableRange.getId());
        
        // Different parts of the editable range link to each other.
        Assert.assertEquals(editableRangeStart.getId(), editableRange.getEditableRangeStart().getId());
        Assert.assertEquals(editableRangeStart.getId(), editableRangeEnd.getEditableRangeStart().getId());
        Assert.assertEquals(editableRange.getId(), editableRangeStart.getEditableRange().getId());
        Assert.assertEquals(editableRangeEnd.getId(), editableRange.getEditableRangeEnd().getId());

        // We can access the node types of each part like this. The editable range itself is not a node,
        // but an entity which consists of a start, an end, and their enclosed contents.
        Assert.assertEquals(NodeType.EDITABLE_RANGE_START, editableRangeStart.getNodeType());
        Assert.assertEquals(NodeType.EDITABLE_RANGE_END, editableRangeEnd.getNodeType());

        builder.writeln("This paragraph is outside the editable range, and cannot be edited.");

        doc.save(getArtifactsDir() + "EditableRange.CreateAndRemove.docx");

        // Remove an editable range. All the nodes that were inside the range will remain intact.
        editableRange.remove();
        //ExEnd

        Assert.assertEquals("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                        "This paragraph is inside an editable range, and can be edited.\r" +
                        "This paragraph is outside the editable range, and cannot be edited.", doc.getText().trim());
        Assert.assertEquals(0, doc.getChildNodes(NodeType.EDITABLE_RANGE_START, true).getCount());

        doc = new Document(getArtifactsDir() + "EditableRange.CreateAndRemove.docx");

        Assert.assertEquals(ProtectionType.READ_ONLY, doc.getProtectionType());
        Assert.assertEquals("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                        "This paragraph is inside an editable range, and can be edited.\r" +
                        "This paragraph is outside the editable range, and cannot be edited.", doc.getText().trim());

        editableRange = ((EditableRangeStart)doc.getChild(NodeType.EDITABLE_RANGE_START, 0, true)).getEditableRange();

        TestUtil.verifyEditableRange(0, "", EditorType.UNSPECIFIED, editableRange);
    }


    @Test
    public void nested() throws Exception
    {
        //ExStart
        //ExFor:DocumentBuilder.StartEditableRange
        //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
        //ExFor:EditableRange.EditorGroup
        //ExSummary:Shows how to create nested editable ranges.
        Document doc = new Document();
        doc.protect(ProtectionType.READ_ONLY, "MyPassword");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world! Since we have set the document's protection level to read-only, " +
                        "we cannot edit this paragraph without the password.");
         
        // Create two nested editable ranges.
        EditableRangeStart outerEditableRangeStart = builder.startEditableRange();
        builder.writeln("This paragraph inside the outer editable range and can be edited.");

        EditableRangeStart innerEditableRangeStart = builder.startEditableRange();
        builder.writeln("This paragraph inside both the outer and inner editable ranges and can be edited.");

        // Currently, the document builder's node insertion cursor is in more than one ongoing editable range.
        // When we want to end an editable range in this situation,
        // we need to specify which of the ranges we wish to end by passing its EditableRangeStart node.
        builder.endEditableRange(innerEditableRangeStart);

        builder.writeln("This paragraph inside the outer editable range and can be edited.");

        builder.endEditableRange(outerEditableRangeStart);

        builder.writeln("This paragraph is outside any editable ranges, and cannot be edited.");

        // If a region of text has two overlapping editable ranges with specified groups,
        // the combined group of users excluded by both groups are prevented from editing it.
        outerEditableRangeStart.getEditableRange().setEditorGroup(EditorType.EVERYONE);
        innerEditableRangeStart.getEditableRange().setEditorGroup(EditorType.CONTRIBUTORS);

        doc.save(getArtifactsDir() + "EditableRange.Nested.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "EditableRange.Nested.docx");

        Assert.assertEquals("Hello world! Since we have set the document's protection level to read-only, we cannot edit this paragraph without the password.\r" +
                        "This paragraph inside the outer editable range and can be edited.\r" +
                        "This paragraph inside both the outer and inner editable ranges and can be edited.\r" +
                        "This paragraph inside the outer editable range and can be edited.\r" +
                        "This paragraph is outside any editable ranges, and cannot be edited.", doc.getText().trim());

        EditableRange editableRange = ((EditableRangeStart)doc.getChild(NodeType.EDITABLE_RANGE_START, 0, true)).getEditableRange();

        TestUtil.verifyEditableRange(0, "", EditorType.EVERYONE, editableRange);

        editableRange = ((EditableRangeStart)doc.getChild(NodeType.EDITABLE_RANGE_START, 1, true)).getEditableRange();

        TestUtil.verifyEditableRange(1, "", EditorType.CONTRIBUTORS, editableRange);
    }

    //ExStart

    //ExFor:EditableRange
    //ExFor:EditableRange.EditorGroup
    //ExFor:EditableRange.SingleUser
    //ExFor:EditableRangeEnd
    //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
    //ExFor:EditableRangeStart
    //ExFor:EditableRangeStart.Accept(DocumentVisitor)
    //ExFor:EditorType
    //ExSummary:Shows how to limit the editing rights of editable ranges to a specific group/user.
    @Test //ExSkip
    public void visitor() throws Exception
    {
        Document doc = new Document();
        doc.protect(ProtectionType.READ_ONLY, "MyPassword");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world! Since we have set the document's protection level to read-only," +
                        " we cannot edit this paragraph without the password.");

        // When we write-protect documents, editable ranges allow us to pick specific areas that users may edit.
        // There are two mutually exclusive ways to narrow down the list of allowed editors.
        // 1 -  Specify a user:
        EditableRange editableRange = builder.startEditableRange().getEditableRange();
        editableRange.setSingleUser("john.doe@myoffice.com");
        builder.writeln($"This paragraph is inside the first editable range, can only be edited by {editableRange.SingleUser}.");
        builder.endEditableRange();

        Assert.assertEquals(EditorType.UNSPECIFIED, editableRange.getEditorGroup());

        // 2 -  Specify a group that allowed users are associated with:
        editableRange = builder.startEditableRange().getEditableRange();
        editableRange.setEditorGroup(EditorType.ADMINISTRATORS);
        builder.writeln($"This paragraph is inside the first editable range, can only be edited by {editableRange.EditorGroup}.");
        builder.endEditableRange();

        Assert.assertEquals("", editableRange.getSingleUser());

        builder.writeln("This paragraph is outside the editable range, and cannot be edited by anybody.");

        // Print details and contents of every editable range in the document.
        EditableRangePrinter editableRangePrinter = new EditableRangePrinter();

        doc.accept(editableRangePrinter);

        System.out.println(editableRangePrinter.toText());
    }

    /// <summary>
    /// Collects properties and contents of visited editable ranges in a string.
    /// </summary>
    public static class EditableRangePrinter extends DocumentVisitor
    {
        public EditableRangePrinter()
        {
            mBuilder = new StringBuilder();
        }

        public String toText()
        {
            return mBuilder.toString();
        }

        public void reset()
        {
            mBuilder.Clear();
            mInsideEditableRange = false;
        }

        /// <summary>
        /// Called when an EditableRangeStart node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitEditableRangeStart(EditableRangeStart editableRangeStart)
        {
            msStringBuilder.appendLine(mBuilder, " -- Editable range found! -- ");
            msStringBuilder.appendLine(mBuilder, "\tID:\t\t" + editableRangeStart.getId());
            if (msString.equals(editableRangeStart.getEditableRange().getSingleUser(), ""))
                msStringBuilder.appendLine(mBuilder, "\tGroup:\t" + editableRangeStart.getEditableRange().getEditorGroup());
            else
                msStringBuilder.appendLine(mBuilder, "\tUser:\t" + editableRangeStart.getEditableRange().getSingleUser());
            msStringBuilder.appendLine(mBuilder, "\tContents:");

            mInsideEditableRange = true;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an EditableRangeEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
        {
            msStringBuilder.appendLine(mBuilder, " -- End of editable range --\n");

            mInsideEditableRange = false;

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document. This visitor only records runs that are inside editable ranges.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            if (mInsideEditableRange) msStringBuilder.appendLine(mBuilder, "\t\"" + run.getText() + "\"");

            return VisitorAction.CONTINUE;
        }

        private boolean mInsideEditableRange;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd

    @Test
    public void incorrectStructureException() throws Exception
    {
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Assert that isn't valid structure for the current document.
        Assert.That(() => builder.endEditableRange(), Throws.<IllegalStateException>TypeOf());

        builder.startEditableRange();
    }

    @Test
    public void incorrectStructureDoNotAdded() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();
        DocumentBuilder builder = new DocumentBuilder(doc);

        EditableRangeStart startRange1 = builder.startEditableRange();

        builder.writeln("EditableRange_1_1");
        builder.writeln("EditableRange_1_2");

        startRange1.getEditableRange().setEditorGroup(EditorType.EVERYONE);
        doc = DocumentHelper.saveOpen(doc);

        // Assert that it's not valid structure and editable ranges aren't added to the current document.
        NodeCollection startNodes = doc.getChildNodes(NodeType.EDITABLE_RANGE_START, true);
        Assert.assertEquals(0, startNodes.getCount());

        NodeCollection endNodes = doc.getChildNodes(NodeType.EDITABLE_RANGE_END, true);
        Assert.assertEquals(0, endNodes.getCount());
    }
}
