// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.EditableRangeStart;
import com.aspose.words.EditableRange;
import com.aspose.words.EditableRangeEnd;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.words.Run;
import com.aspose.words.EditorType;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.SaveFormat;
import com.aspose.words.NodeCollection;


@Test
class ExEditableRange !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void removesEditableRange() throws Exception
    {
        //ExStart
        //ExFor:EditableRange.Remove
        //ExSummary:Shows how to remove an editable range from a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an EditableRange so we can remove it. Does not have to be well-formed
        EditableRangeStart edRange1Start = builder.startEditableRange();
        EditableRange editableRange1 = edRange1Start.getEditableRange();
        builder.writeln("Paragraph inside editable range");
        EditableRangeEnd edRange1End = builder.endEditableRange();

        // Remove the range that was just made
        editableRange1.remove();
        //ExEnd
    }

    //ExStart
    //ExFor:DocumentBuilder.StartEditableRange
    //ExFor:DocumentBuilder.EndEditableRange
    //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
    //ExFor:EditableRange
    //ExFor:EditableRange.EditableRangeEnd
    //ExFor:EditableRange.EditableRangeStart
    //ExFor:EditableRange.Id
    //ExFor:EditableRange.SingleUser
    //ExFor:EditableRangeEnd
    //ExFor:EditableRangeEnd.Accept(DocumentVisitor)
    //ExFor:EditableRangeEnd.EditableRangeStart
    //ExFor:EditableRangeEnd.Id
    //ExFor:EditableRangeEnd.NodeType
    //ExFor:EditableRangeStart
    //ExFor:EditableRangeStart.Accept(DocumentVisitor)
    //ExFor:EditableRangeStart.EditableRange
    //ExFor:EditableRangeStart.Id
    //ExFor:EditableRangeStart.NodeType
    //ExSummary:Shows how to start and end an editable range.
    @Test //ExSkip
    public void createEditableRanges() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start an editable range
        EditableRangeStart edRange1Start = builder.startEditableRange();

        // An EditableRange object is created for the EditableRangeStart that we just made
        EditableRange editableRange1 = edRange1Start.getEditableRange();

        // Put something inside the editable range
        builder.writeln("Paragraph inside first editable range");

        // An editable range is well-formed if it has a start and an end
        // Multiple editable ranges can be nested and overlapping 
        EditableRangeEnd edRange1End = builder.endEditableRange();

        // Explicitly state which EditableRangeStart a new EditableRangeEnd should be paired with
        EditableRangeStart edRange2Start = builder.startEditableRange();
        builder.writeln("Paragraph inside second editable range");
        EditableRange editableRange2 = edRange2Start.getEditableRange();
        EditableRangeEnd edRange2End = builder.endEditableRange(edRange2Start);

        // Editable range starts and ends have their own respective node types
        msAssert.areEqual(NodeType.EDITABLE_RANGE_START, edRange1Start.getNodeType());
        msAssert.areEqual(NodeType.EDITABLE_RANGE_END, edRange1End.getNodeType());

        // Editable range IDs are unique and set automatically
        msAssert.areEqual(0, editableRange1.getId());
        msAssert.areEqual(1, editableRange2.getId());

        // Editable range starts and ends always belong to a range
        msAssert.areEqual(edRange1Start, editableRange1.getEditableRangeStart());
        msAssert.areEqual(edRange1End, editableRange1.getEditableRangeEnd());

        // They also inherit the ID of the entire editable range that they belong to
        msAssert.areEqual(editableRange1.getId(), edRange1Start.getId());
        msAssert.areEqual(editableRange1.getId(), edRange1End.getId());
        msAssert.areEqual(editableRange2.getId(), edRange2Start.getEditableRange().getId());
        msAssert.areEqual(editableRange2.getId(), edRange2End.getEditableRangeStart().getEditableRange().getId());

        // If the editable range was found in a document, it will probably have something in the single user property
        // But if we make one programmatically, the property is empty by default
        msAssert.areEqual("", editableRange1.getSingleUser());

        // We have to set it ourselves if we want the ranges to belong to somebody
        editableRange1.setSingleUser("john.doe@myoffice.com");
        editableRange2.setSingleUser("jane.doe@myoffice.com");

        // Initialize a custom visitor for editable ranges that will print their contents 
        EditableRangeInfoPrinter editableRangeReader = new EditableRangeInfoPrinter();

        // Both the start and end of an editable range can accept visitors, but not the editable range itself
        edRange1Start.accept(editableRangeReader);
        edRange2End.accept(editableRangeReader);

        // Or, if we want to go over all the editable ranges in a document, we can get the document to accept the visitor
        editableRangeReader.reset();
        doc.accept(editableRangeReader);

        msConsole.writeLine(editableRangeReader.toText());
    }

    /// <summary>
    /// Visitor implementation that prints attributes and contents of ranges.
    /// </summary>
    public static class EditableRangeInfoPrinter extends DocumentVisitor
    {
        public EditableRangeInfoPrinter()
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
            msStringBuilder.appendLine(mBuilder, "\tID: " + editableRangeStart.getId());
            msStringBuilder.appendLine(mBuilder, "\tUser: " + editableRangeStart.getEditableRange().getSingleUser());
            msStringBuilder.appendLine(mBuilder, "\tContents: ");

            mInsideEditableRange = true;

            // Let the visitor continue visiting other nodes
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an EditableRangeEnd node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
        {
            msStringBuilder.appendLine(mBuilder, " -- End of editable range -- ");

            mInsideEditableRange = false;

            // Let the visitor continue visiting other nodes
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when a Run node is encountered in the document. Only runs within editable ranges have their contents recorded.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            if (mInsideEditableRange) msStringBuilder.appendLine(mBuilder, "\t\"" + run.getText() + "\"");

            // Let the visitor continue visiting other nodes
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

        // Checking that isn't valid structure for the current document
        Assert.That(() => builder.endEditableRange(), Throws.<IllegalStateException>TypeOf());

        builder.startEditableRange();
    }

    @Test
    public void incorrectStructureDoNotAdded() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //ExStart
        //ExFor:EditableRange.EditorGroup
        //ExFor:EditorType
        //ExSummary:Shows how to add editing group for editable ranges
        EditableRangeStart startRange1 = builder.startEditableRange();

        builder.writeln("EditableRange_1_1");
        builder.writeln("EditableRange_1_2");

        // Sets the editor for editable range region
        startRange1.getEditableRange().setEditorGroup(EditorType.EVERYONE);
        //ExEnd

        MemoryStream dstStream = new MemoryStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // Assert that it's not valid structure and editable ranges aren't added to the current document
        NodeCollection startNodes = doc.getChildNodes(NodeType.EDITABLE_RANGE_START, true);
        msAssert.areEqual(0, startNodes.getCount());

        NodeCollection endNodes = doc.getChildNodes(NodeType.EDITABLE_RANGE_END, true);
        msAssert.areEqual(0, endNodes.getCount());
    }
}
