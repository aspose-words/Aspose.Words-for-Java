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
import org.testng.Assert;
import com.aspose.words.Run;
import com.aspose.words.Paragraph;
import com.aspose.words.RunCollection;


@Test
class ExInline !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void inlineRevisions() throws Exception
    {
        //ExStart
        //ExFor:Inline
        //ExFor:Inline.IsDeleteRevision
        //ExFor:Inline.IsFormatRevision
        //ExFor:Inline.IsInsertRevision
        //ExFor:Inline.IsMoveFromRevision
        //ExFor:Inline.IsMoveToRevision
        //ExFor:Inline.ParentParagraph
        //ExFor:Paragraph.Runs
        //ExFor:Revision.ParentNode
        //ExFor:RunCollection
        //ExFor:RunCollection.Item(Int32)
        //ExFor:RunCollection.ToArray
        //ExSummary:Shows how to view revision-related properties of Inline nodes.
        Document doc = new Document(getMyDir() + "Revision runs.docx");

        // This document has 6 revisions
        Assert.assertEquals(6, doc.getRevisions().getCount());

        // The parent node of a revision is the run that the revision concerns, which is an Inline node
        Run run = (Run)doc.getRevisions().get(0).getParentNode();

        // Get the parent paragraph
        Paragraph firstParagraph = run.getParentParagraph();
        RunCollection runs = firstParagraph.getRuns();

        Assert.assertEquals(6, runs.toArray().length);

        // The text in the run at index #2 was typed after revisions were tracked, so it will count as an insert revision
        // The font was changed, so it will also be a format revision
        Assert.assertTrue(runs.get(2).isInsertRevision());
        Assert.assertTrue(runs.get(2).isFormatRevision());

        // If one node was moved from one place to another while changes were tracked,
        // the node will be placed at the departure location as a "move to revision",
        // and a "move from revision" node will be left behind at the origin, in case we want to reject changes
        // Highlighting text and dragging it to another place with the mouse and cut-and-pasting (but not copy-pasting) both count as "move revisions"
        // The node with the "IsMoveToRevision" flag is the arrival of the move operation, and the node with the "IsMoveFromRevision" flag is the departure point
        Assert.assertTrue(runs.get(1).isMoveToRevision());
        Assert.assertTrue(runs.get(4).isMoveFromRevision());

        // If an Inline node gets deleted while changes are being tracked, it will leave behind a node with the IsDeleteRevision flag set to true until changes are accepted
        Assert.assertTrue(runs.get(5).isDeleteRevision());
        //ExEnd
    }
}

