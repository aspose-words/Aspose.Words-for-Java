package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;

public class ExTabStopCollection extends ApiExampleBase {
    @Test
    public void clearAll() throws Exception {
        //ExStart
        //ExFor:TabStopCollection.Clear
        //ExSummary:Shows how to remove all tab stops from a document.
        Document doc = new Document(getMyDir() + "Table of contents.docx");

        // Clear all tab stops from every paragraph
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            para.getParagraphFormat().getTabStops().clear();
        }

        doc.save(getArtifactsDir() + "TabStopCollection.ClearAll.docx");
        //ExEnd
    }

    @Test
    public void tabStops() throws Exception {
        //ExStart
        //ExFor:TabStop.#ctor
        //ExFor:TabStop.#ctor(Double)
        //ExFor:TabStop.#ctor(Double,TabAlignment,TabLeader)
        //ExFor:TabStop.Equals(TabStop)
        //ExFor:TabStop.IsClear
        //ExFor:TabStopCollection
        //ExFor:TabStopCollection.After(Double)
        //ExFor:TabStopCollection.Before(Double)
        //ExFor:TabStopCollection.Count
        //ExFor:TabStopCollection.Equals(TabStopCollection)
        //ExFor:TabStopCollection.Equals(Object)
        //ExFor:TabStopCollection.GetHashCode
        //ExFor:TabStopCollection.Item(Double)
        //ExFor:TabStopCollection.Item(Int32)
        //ExSummary:Shows how to add tab stops to a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Access the collection of tab stops and add some tab stops to it
        TabStopCollection tabStops = builder.getParagraphFormat().getTabStops();

        // 72 points is one "inch" on the Microsoft Word tab stop ruler
        tabStops.add(new TabStop(72.0));
        tabStops.add(new TabStop(432.0, TabAlignment.RIGHT, TabLeader.DASHES));

        Assert.assertEquals(tabStops.getCount(), 2);
        Assert.assertFalse(tabStops.get(0).isClear());
        Assert.assertFalse(tabStops.get(0).equals(tabStops.get(1)));

        builder.writeln("Start\tTab 1\tTab 2");

        // Get the collection of paragraphs that we've created
        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();
        Assert.assertEquals(paragraphs.getCount(), 2);

        // Each paragraph gets its own TabStopCollection which gets values from the DocumentBuilder's collection
        Assert.assertEquals(paragraphs.get(0).getParagraphFormat().getTabStops(), paragraphs.get(1).getParagraphFormat().getTabStops());
        Assert.assertNotSame(paragraphs.get(0).getParagraphFormat().getTabStops(), paragraphs.get(1).getParagraphFormat().getTabStops());
        Assert.assertEquals(paragraphs.get(0).getParagraphFormat().getTabStops().hashCode(),
                paragraphs.get(1).getParagraphFormat().getTabStops().hashCode());

        // A TabStopCollection can point us to TabStops before and after certain positions
        Assert.assertEquals(72.0, tabStops.before(100.0).getPosition());
        Assert.assertEquals(432.0, tabStops.after(100.0).getPosition());

        doc.save(getArtifactsDir() + "TabStopCollection.TabStops.docx");
        //ExEnd
    }

    @Test
    public void addTabStops() throws Exception {
        //ExStart
        //ExFor:TabStopCollection.Add(TabStop)
        //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
        //ExSummary:Shows how to create tab stops and add them to a document.
        Document doc = new Document(getMyDir() + "Document.docx");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Create a TabStop object and add it to the document
        TabStop tabStop = new TabStop(ConvertUtil.inchToPoint(3.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(tabStop);

        // Add a tab stop without explicitly creating new TabStop objects
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(100.0), TabAlignment.LEFT, TabLeader.DASHES);

        // Add tab stops at 5 cm to all paragraphs
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true)) {
            para.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(50.0), TabAlignment.LEFT, TabLeader.DASHES);
        }

        doc.save(getArtifactsDir() + "TabStopCollection.AddTabStops.doc");
        //ExEnd
    }

    @Test
    public void removeByIndex() throws Exception {
        //ExStart
        //ExFor:TabStopCollection.RemoveByIndex
        //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
        Document doc = new Document(getMyDir() + "Document.docx");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(60.0), TabAlignment.LEFT, TabLeader.DASHES);

        // Tab stop placed at 30 mm is removed
        paragraph.getParagraphFormat().getTabStops().removeByIndex(0);

        System.out.println(paragraph.getParagraphFormat().getTabStops().getCount());

        doc.save(getArtifactsDir() + "TabStopCollection.RemoveByIndex.doc");
        //ExEnd
    }

    @Test
    public void getPositionByIndexEx() throws Exception {
        //ExStart
        //ExFor:TabStopCollection.GetPositionByIndex
        //ExSummary:Shows how to find a tab stop by it's index and get its position.
        Document doc = new Document(getMyDir() + "Document.docx");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(60.0), TabAlignment.LEFT, TabLeader.DASHES);

        System.out.println(MessageFormat.format("Tab stop at index {0} of the first paragraph is at {1} points.", 1, paragraph.getParagraphFormat().getTabStops().getPositionByIndex(1)));
        //ExEnd
    }

    @Test
    public void getIndexByPositionEx() throws Exception {
        //ExStart
        //ExFor:TabStopCollection.GetIndexByPosition
        //ExSummary:Shows how to look up a position to see if a tab stop exists there, and if so, obtain its index.
        Document doc = new Document(getMyDir() + "Document.docx");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);

        // An output of -1 signifies that there is no tab stop at that position.
        System.out.println(paragraph.getParagraphFormat().getTabStops().getIndexByPosition(ConvertUtil.millimeterToPoint(30.0))); // 0
        System.out.println(paragraph.getParagraphFormat().getTabStops().getIndexByPosition(ConvertUtil.millimeterToPoint(60.0))); // -1
        //ExEnd
    }
}
