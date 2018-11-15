//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.EditableRangeStart;
import com.aspose.words.EditableRange;
import com.aspose.words.EditableRangeEnd;
import org.testng.Assert;
import com.aspose.words.EditorType;
import com.aspose.words.SaveFormat;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;

import java.io.ByteArrayOutputStream;

public class ExEditableRange extends ApiExampleBase
{
    @Test
    public void removesEditableRange() throws Exception
    {
        //ExStart
        //ExFor:EditableRange.Remove
        //ExSummary:Shows how to remove an editable range from a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create an EditableRange so we can remove it. Does not have to be well-formed.
        EditableRangeStart edRange1Start = builder.startEditableRange();
        EditableRange editableRange1 = edRange1Start.getEditableRange();
        builder.writeln("Paragraph inside editable range");
        EditableRangeEnd edRange1End = builder.endEditableRange();

        // Remove the range that was just made.
        editableRange1.remove();
        //ExEnd
    }

    @Test
    public void createEditableRanges() throws Exception
        {
        //ExStart
        //ExFor:DocumentBuilder.StartEditableRange
        //ExFor:DocumentBuilder.EndEditableRange
        //ExFor:DocumentBuilder.EndEditableRange(EditableRangeStart)
        //ExSummary:Shows how to start and end an editable range.
        Document doc = new Document(getMyDir() + "Document.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start an editable range.
        EditableRangeStart edRange1Start = builder.startEditableRange();

        // An EditableRange object is created for the EditableRangeStart that we just made.
        EditableRange editableRange1 = edRange1Start.getEditableRange();

        // Put something inside the editable range.
        builder.writeln("Paragraph inside first editable range");

        // An editable range is well-formed if it has a start and an end. 
        // Multiple editable ranges can be nested and overlapping. 
        EditableRangeEnd edRange1End = builder.endEditableRange();

        // Both the start and end automatically belong to editableRange1.
        System.out.println(editableRange1.getEditableRangeStart().equals(edRange1Start)); // True
        System.out.println(editableRange1.getEditableRangeEnd().equals(edRange1End)); // True

        // Explicitly state which EditableRangeStart a new EditableRangeEnd should be paired with.
        EditableRangeStart edRange2Start = builder.startEditableRange();
        builder.writeln("Paragraph inside second editable range");
        EditableRange editableRange2 = edRange2Start.getEditableRange();
        EditableRangeEnd edRange2End = builder.endEditableRange(edRange2Start);

        // Both the start and end automatically belong to editableRange2.
        System.out.println(editableRange2.getEditableRangeStart().equals(edRange2Start)); // True
        System.out.println(editableRange2.getEditableRangeEnd().equals(edRange2End)); // True
        //ExEnd
        }

    @Test
    public void incorrectStructureException() throws Exception
    {
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);

        //Is not valid structure for the current document
        try
        {
            builder.endEditableRange();
        } catch (Exception e)
        {
            Assert.assertTrue(e instanceof IllegalStateException);
        }

        builder.startEditableRange();
    }

    @Test
    public void incorrectStructureDoNotAdded() throws Exception
    {
        Document doc = DocumentHelper.createDocumentFillWithDummyText();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //ExStart
        //ExFor:EditableRange.EditorGroup
        //ExSummary:Shows how to add editing group for editable ranges
        //Add EditableRangeStart
        EditableRangeStart startRange1 = builder.startEditableRange();

        builder.writeln("EditableRange_1_1");
        builder.writeln("EditableRange_1_2");

        // Sets the editor for editable range region
        startRange1.getEditableRange().setEditorGroup(EditorType.EVERYONE);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, SaveFormat.DOCX);

        // Assert that it's not valid structure and editable ranges aren't added to the current document
        NodeCollection startNodes = doc.getChildNodes(NodeType.EDITABLE_RANGE_START, true);
        Assert.assertEquals(startNodes.getCount(), 0);

        NodeCollection endNodes = doc.getChildNodes(NodeType.EDITABLE_RANGE_END, true);
        Assert.assertEquals(endNodes.getCount(), 0);
    }
}
