//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.TabStop;
import com.aspose.words.ConvertUtil;
import com.aspose.words.TabAlignment;
import com.aspose.words.TabLeader;

import java.text.MessageFormat;


@Test
public class ExTabStopCollection extends ApiExampleBase
{
    @Test
    public void clearEx() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.Clear
        //ExSummary:Shows how to remove all tab stops from a document.
        Document doc = new Document(getMyDir() + "Document.TableOfContents.doc");

        // Clear all tab stops from every paragraph.
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().getTabStops().clear();
        }

        doc.save(getMyDir() + "\\Artifacts\\Document.AllTabStopsRemoved.doc");
        //ExEnd
    }

    @Test
    public void addEx() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.Add(TabStop)
        //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
        //ExSummary:Shows how to create tab stops and add them to a document.
        Document doc = new Document(getMyDir() + "Document.doc");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Create a TabStop object and add it to the document.
        TabStop tabStop = new TabStop(ConvertUtil.inchToPoint(3.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(tabStop);

        // Add a tab stop without explicitly creating new TabStop objects.
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(100.0), TabAlignment.LEFT, TabLeader.DASHES);

        // Add tab stops at 5 cm to all paragraphs.
        for (Paragraph para : (Iterable<Paragraph>) doc.getChildNodes(NodeType.PARAGRAPH, true))
        {
            para.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(50.0), TabAlignment.LEFT, TabLeader.DASHES);
        }

        doc.save(getMyDir() + "\\Artifacts\\Document.AddedTabStops.doc");
        //ExEnd
    }

    @Test
    public void removeByIndexEx() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.RemoveByIndex
        //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
        Document doc = new Document(getMyDir() + "Document.doc");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(60.0), TabAlignment.LEFT, TabLeader.DASHES);

        // Tab stop placed at 30 mm is removed
        paragraph.getParagraphFormat().getTabStops().removeByIndex(0);

        System.out.println(paragraph.getParagraphFormat().getTabStops().getCount());

        doc.save(getMyDir() + "\\Artifacts\\Document.RemovedTabStopsByIndex.doc");
        //ExEnd
    }

    @Test
    public void getPositionByIndexEx() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.GetPositionByIndex
        //ExSummary:Shows how to find a tab stop by it's index and get its position.
        Document doc = new Document(getMyDir() + "Document.doc");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(60.0), TabAlignment.LEFT, TabLeader.DASHES);

        System.out.println(MessageFormat.format("Tab stop at index {0} of the first paragraph is at {1} points.", 1, paragraph.getParagraphFormat().getTabStops().getPositionByIndex(1)));
        //ExEnd
    }

    @Test
    public void getIndexByPositionEx() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.GetIndexByPosition
        //ExSummary:Shows how to look up a position to see if a tab stop exists there, and if so, obtain its index.
        Document doc = new Document(getMyDir() + "Document.doc");
        Paragraph paragraph = (Paragraph) doc.getChild(NodeType.PARAGRAPH, 0, true);

        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);

        // An output of -1 signifies that there is no tab stop at that position.
        System.out.println(paragraph.getParagraphFormat().getTabStops().getIndexByPosition(ConvertUtil.millimeterToPoint(30.0))); // 0
        System.out.println(paragraph.getParagraphFormat().getTabStops().getIndexByPosition(ConvertUtil.millimeterToPoint(60.0))); // -1
        //ExEnd
    }
}
