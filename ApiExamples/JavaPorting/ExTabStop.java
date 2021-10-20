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
import com.aspose.words.Paragraph;
import com.aspose.words.NodeType;
import com.aspose.words.TabStop;
import com.aspose.words.ConvertUtil;
import com.aspose.words.TabAlignment;
import com.aspose.words.TabLeader;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.TabStopCollection;
import org.testng.Assert;
import com.aspose.words.ParagraphCollection;


@Test
public class ExTabStop extends ApiExampleBase
{
    @Test
    public void addTabStops() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.Add(TabStop)
        //ExFor:TabStopCollection.Add(Double, TabAlignment, TabLeader)
        //ExSummary:Shows how to add custom tab stops to a document.
        Document doc = new Document();
        Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);

        // Below are two ways of adding tab stops to a paragraph's collection of tab stops via the "ParagraphFormat" property.
        // 1 -  Create a "TabStop" object, and then add it to the collection:
        TabStop tabStop = new TabStop(ConvertUtil.inchToPoint(3.0), TabAlignment.LEFT, TabLeader.DASHES);
        paragraph.getParagraphFormat().getTabStops().add(tabStop);

        // 2 -  Pass the values for properties of a new tab stop to the "Add" method:
        paragraph.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(100.0), TabAlignment.LEFT,
            TabLeader.DASHES);

        // Add tab stops at 5 cm to all paragraphs.
        for (Paragraph para : doc.getChildNodes(NodeType.PARAGRAPH, true).<Paragraph>OfType() !!Autoporter error: Undefined expression type )
        {
            para.getParagraphFormat().getTabStops().add(ConvertUtil.millimeterToPoint(50.0), TabAlignment.LEFT,
                TabLeader.DASHES);
        }

        // Every "tab" character takes the builder's cursor to the location of the next tab stop.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Start\tTab 1\tTab 2\tTab 3\tTab 4");

        doc.save(getArtifactsDir() + "TabStopCollection.AddTabStops.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "TabStopCollection.AddTabStops.docx");
        TabStopCollection tabStops = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getTabStops();

        TestUtil.verifyTabStop(141.75d, TabAlignment.LEFT, TabLeader.DASHES, false, tabStops.get(0));
        TestUtil.verifyTabStop(216.0d, TabAlignment.LEFT, TabLeader.DASHES, false, tabStops.get(1));
        TestUtil.verifyTabStop(283.45d, TabAlignment.LEFT, TabLeader.DASHES, false, tabStops.get(2));
    }

    @Test
    public void tabStopCollection() throws Exception
    {
        //ExStart
        //ExFor:TabStop.#ctor
        //ExFor:TabStop.#ctor(Double)
        //ExFor:TabStop.#ctor(Double,TabAlignment,TabLeader)
        //ExFor:TabStop.Equals(TabStop)
        //ExFor:TabStop.IsClear
        //ExFor:TabStopCollection
        //ExFor:TabStopCollection.After(Double)
        //ExFor:TabStopCollection.Before(Double)
        //ExFor:TabStopCollection.Clear
        //ExFor:TabStopCollection.Count
        //ExFor:TabStopCollection.Equals(TabStopCollection)
        //ExFor:TabStopCollection.Equals(Object)
        //ExFor:TabStopCollection.GetHashCode
        //ExFor:TabStopCollection.Item(Double)
        //ExFor:TabStopCollection.Item(Int32)
        //ExSummary:Shows how to work with a document's collection of tab stops.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        TabStopCollection tabStops = builder.getParagraphFormat().getTabStops();

        // 72 points is one "inch" on the Microsoft Word tab stop ruler.
        tabStops.add(new TabStop(72.0));
        tabStops.add(new TabStop(432.0, TabAlignment.RIGHT, TabLeader.DASHES));

        Assert.assertEquals(2, tabStops.getCount());
        Assert.assertFalse(tabStops.get(0).isClear());
        Assert.assertFalse(tabStops.get(0).equals(tabStops.get(1)));

        // Every "tab" character takes the builder's cursor to the location of the next tab stop.
        builder.writeln("Start\tTab 1\tTab 2");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(2, paragraphs.getCount());

        // Each paragraph gets its tab stop collection, which clones its values from the document builder's tab stop collection.
        Assert.assertEquals(paragraphs.get(0).getParagraphFormat().getTabStops(), paragraphs.get(1).getParagraphFormat().getTabStops());
        Assert.assertNotSame(paragraphs.get(0).getParagraphFormat().getTabStops(), paragraphs.get(1).getParagraphFormat().getTabStops());

        // A tab stop collection can point us to TabStops before and after certain positions.
        Assert.assertEquals(72.0, tabStops.before(100.0).getPosition());
        Assert.assertEquals(432.0, tabStops.after(100.0).getPosition());

        // We can clear a paragraph's tab stop collection to revert to the default tabbing behavior.
        paragraphs.get(1).getParagraphFormat().getTabStops().clear();

        Assert.assertEquals(0, paragraphs.get(1).getParagraphFormat().getTabStops().getCount());

        doc.save(getArtifactsDir() + "TabStopCollection.TabStopCollection.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "TabStopCollection.TabStopCollection.docx");
        tabStops = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getTabStops();

        Assert.assertEquals(2, tabStops.getCount());
        TestUtil.verifyTabStop(72.0d, TabAlignment.LEFT, TabLeader.NONE, false, tabStops.get(0));
        TestUtil.verifyTabStop(432.0d, TabAlignment.RIGHT, TabLeader.DASHES, false, tabStops.get(1));

        tabStops = doc.getFirstSection().getBody().getParagraphs().get(1).getParagraphFormat().getTabStops();

        Assert.assertEquals(0, tabStops.getCount());
    }

    @Test
    public void removeByIndex() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.RemoveByIndex
        //ExSummary:Shows how to select a tab stop in a document by its index and remove it.
        Document doc = new Document();
        TabStopCollection tabStops = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getTabStops();

        tabStops.add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);
        tabStops.add(ConvertUtil.millimeterToPoint(60.0), TabAlignment.LEFT, TabLeader.DASHES);

        Assert.assertEquals(2, tabStops.getCount());

        // Remove the first tab stop.
        tabStops.removeByIndex(0);

        Assert.assertEquals(1, tabStops.getCount());

        doc.save(getArtifactsDir() + "TabStopCollection.RemoveByIndex.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "TabStopCollection.RemoveByIndex.docx");

        TestUtil.verifyTabStop(170.1d, TabAlignment.LEFT, TabLeader.DASHES, false, doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getTabStops().get(0));
    }

    @Test
    public void getPositionByIndex() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.GetPositionByIndex
        //ExSummary:Shows how to find a tab, stop by its index and verify its position.
        Document doc = new Document();
        TabStopCollection tabStops = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getTabStops();

        tabStops.add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);
        tabStops.add(ConvertUtil.millimeterToPoint(60.0), TabAlignment.LEFT, TabLeader.DASHES);

        // Verify the position of the second tab stop in the collection.
        Assert.assertEquals(ConvertUtil.millimeterToPoint(60.0), tabStops.getPositionByIndex(1), 0.1d);
        //ExEnd
    }

    @Test
    public void getIndexByPosition() throws Exception
    {
        //ExStart
        //ExFor:TabStopCollection.GetIndexByPosition
        //ExSummary:Shows how to look up a position to see if a tab stop exists there and obtain its index.
        Document doc = new Document();
        TabStopCollection tabStops = doc.getFirstSection().getBody().getParagraphs().get(0).getParagraphFormat().getTabStops();

        // Add a tab stop at a position of 30mm.
        tabStops.add(ConvertUtil.millimeterToPoint(30.0), TabAlignment.LEFT, TabLeader.DASHES);

        // A result of "0" returned by "GetIndexByPosition" confirms that a tab stop
        // at 30mm exists in this collection, and it is at index 0.
        Assert.assertEquals(0, tabStops.getIndexByPosition(ConvertUtil.millimeterToPoint(30.0)));

        // A "-1" returned by "GetIndexByPosition" confirms that
        // there is no tab stop in this collection with a position of 60mm.
        Assert.assertEquals(-1, tabStops.getIndexByPosition(ConvertUtil.millimeterToPoint(60.0)));
        //ExEnd
    }
}
