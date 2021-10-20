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
import com.aspose.words.LayoutCollector;
import org.testng.Assert;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BreakType;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Node;
import com.aspose.ms.System.msConsole;
import com.aspose.words.LayoutEnumerator;
import com.aspose.words.LayoutEntityType;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.Drawing.RectangleF;
import com.aspose.words.IPageLayoutCallback;
import com.aspose.words.PageLayoutCallbackArgs;
import com.aspose.words.PageLayoutEvent;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.SaveFormat;
import com.aspose.words.PageSet;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.ContinuousSectionRestart;


@Test
class ExLayout !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void layoutCollector() throws Exception
    {
        //ExStart
        //ExFor:Layout.LayoutCollector
        //ExFor:Layout.LayoutCollector.#ctor(Document)
        //ExFor:Layout.LayoutCollector.Clear
        //ExFor:Layout.LayoutCollector.Document
        //ExFor:Layout.LayoutCollector.GetEndPageIndex(Node)
        //ExFor:Layout.LayoutCollector.GetEntity(Node)
        //ExFor:Layout.LayoutCollector.GetNumPagesSpanned(Node)
        //ExFor:Layout.LayoutCollector.GetStartPageIndex(Node)
        //ExFor:Layout.LayoutEnumerator.Current
        //ExSummary:Shows how to see the the ranges of pages that a node spans.
        Document doc = new Document();
        LayoutCollector layoutCollector = new LayoutCollector(doc);
        
        // Call the "GetNumPagesSpanned" method to count how many pages the content of our document spans.
        // Since the document is empty, that number of pages is currently zero.
        Assert.assertEquals(doc, layoutCollector.getDocument());
        Assert.assertEquals(0, layoutCollector.getNumPagesSpanned(doc));

        // Populate the document with 5 pages of content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Section 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.SECTION_BREAK_EVEN_PAGE);
        builder.write("Section 2");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.insertBreak(BreakType.PAGE_BREAK);

        // Before the layout collector, we need to call the "UpdatePageLayout" method to give us
        // an accurate figure for any layout-related metric, such as the page count.
        Assert.assertEquals(0, layoutCollector.getNumPagesSpanned(doc));

        layoutCollector.clear();
        doc.updatePageLayout();

        Assert.assertEquals(5, layoutCollector.getNumPagesSpanned(doc));

        // We can see the numbers of the start and end pages of any node and their overall page spans.
        NodeCollection nodes = doc.getChildNodes(NodeType.ANY, true);
        for (Node node : (Iterable<Node>) nodes)
        {
            System.out.println("->  NodeType.{node.NodeType}: ");
            System.out.println("\tStarts on page {layoutCollector.GetStartPageIndex(node)}, ends on page {layoutCollector.GetEndPageIndex(node)}," +
                    $" spanning {layoutCollector.GetNumPagesSpanned(node)} pages.");
        }

        // We can iterate over the layout entities using a LayoutEnumerator.
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

        Assert.assertEquals(LayoutEntityType.PAGE, layoutEnumerator.getType());

        // The LayoutEnumerator can traverse the collection of layout entities like a tree.
        // We can also apply it to any node's corresponding layout entity.
        layoutEnumerator.setCurrent(layoutCollector.getEntity(doc.getChild(NodeType.PARAGRAPH, 1, true)));

        Assert.assertEquals(LayoutEntityType.SPAN, layoutEnumerator.getType());
        Assert.assertEquals("¶", layoutEnumerator.getText());
        //ExEnd
    }

    //ExStart
    //ExFor:Layout.LayoutEntityType
    //ExFor:Layout.LayoutEnumerator
    //ExFor:Layout.LayoutEnumerator.#ctor(Document)
    //ExFor:Layout.LayoutEnumerator.Document
    //ExFor:Layout.LayoutEnumerator.Kind
    //ExFor:Layout.LayoutEnumerator.MoveFirstChild
    //ExFor:Layout.LayoutEnumerator.MoveLastChild
    //ExFor:Layout.LayoutEnumerator.MoveNext
    //ExFor:Layout.LayoutEnumerator.MoveNextLogical
    //ExFor:Layout.LayoutEnumerator.MoveParent
    //ExFor:Layout.LayoutEnumerator.MoveParent(Layout.LayoutEntityType)
    //ExFor:Layout.LayoutEnumerator.MovePrevious
    //ExFor:Layout.LayoutEnumerator.MovePreviousLogical
    //ExFor:Layout.LayoutEnumerator.PageIndex
    //ExFor:Layout.LayoutEnumerator.Rectangle
    //ExFor:Layout.LayoutEnumerator.Reset
    //ExFor:Layout.LayoutEnumerator.Text
    //ExFor:Layout.LayoutEnumerator.Type
    //ExSummary:Shows ways of traversing a document's layout entities.
    @Test //ExSkip
    public void layoutEnumerator() throws Exception
    {
        // Open a document that contains a variety of layout entities.
        // Layout entities are pages, cells, rows, lines, and other objects included in the LayoutEntityType enum.
        // Each layout entity has a rectangular space that it occupies in the document body.
        Document doc = new Document(getMyDir() + "Layout entities.docx");

        // Create an enumerator that can traverse these entities like a tree.
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

        Assert.assertEquals(doc, layoutEnumerator.getDocument());

        layoutEnumerator.moveParent(LayoutEntityType.PAGE);

        Assert.assertEquals(LayoutEntityType.PAGE, layoutEnumerator.getType());
        Assert.<IllegalStateException>Throws(() => System.out.println(layoutEnumerator.getText()));

        // We can call this method to make sure that the enumerator will be at the first layout entity.
        layoutEnumerator.reset();

        // There are two orders that determine how the layout enumerator continues traversing layout entities
        // when it encounters entities that span across multiple pages.
        // 1 -  In visual order:
        // When moving through an entity's children that span multiple pages,
        // page layout takes precedence, and we move to other child elements on this page and avoid the ones on the next.
        System.out.println("Traversing from first to last, elements between pages separated:");
        traverseLayoutForward(layoutEnumerator, 1);

        // Our enumerator is now at the end of the collection. We can traverse the layout entities backwards to go back to the beginning.
        System.out.println("Traversing from last to first, elements between pages separated:");
        traverseLayoutBackward(layoutEnumerator, 1);

        // 2 -  In logical order:
        // When moving through an entity's children that span multiple pages,
        // the enumerator will move between pages to traverse all the child entities.
        System.out.println("Traversing from first to last, elements between pages mixed:");
        traverseLayoutForwardLogical(layoutEnumerator, 1);

        System.out.println("Traversing from last to first, elements between pages mixed:");
        traverseLayoutBackwardLogical(layoutEnumerator, 1);
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
    /// in a depth-first manner, and in the "Visual" order.
    /// </summary>
    private static void traverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveFirstChild())
            {
                traverseLayoutForward(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNext());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
    /// in a depth-first manner, and in the "Visual" order.
    /// </summary>
    private static void traverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveLastChild())
            {
                traverseLayoutBackward(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.movePrevious());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
    /// in a depth-first manner, and in the "Logical" order.
    /// </summary>
    private static void traverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveFirstChild())
            {
                traverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNextLogical());
    }

    /// <summary>
    /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
    /// in a depth-first manner, and in the "Logical" order.
    /// </summary>
    private static void traverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth) throws Exception
    {
        do
        {
            printCurrentEntity(layoutEnumerator, depth);

            if (layoutEnumerator.moveLastChild())
            {
                traverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.movePreviousLogical());
    }

    /// <summary>
    /// Print information about layoutEnumerator's current entity to the console, while indenting the text with tab characters
    /// based on its depth relative to the root node that we provided in the constructor LayoutEnumerator instance.
    /// The rectangle that we process at the end represents the area and location that the entity takes up in the document.
    /// </summary>
    private static void printCurrentEntity(LayoutEnumerator layoutEnumerator, int indent) throws Exception
    {
        String tabs = msString.newString('\t', indent);

        System.out.println(msString.equals(layoutEnumerator.getKind(), "")
                ? $"{tabs}-> Entity type: {layoutEnumerator.Type}"
                : $"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");

        // Only spans can contain text.
        if (layoutEnumerator.getType() == LayoutEntityType.SPAN)
            System.out.println("{tabs}   Span contents: \"{layoutEnumerator.Text}\"");

        RectangleF leRect = layoutEnumerator.getRectangleInternal();
        System.out.println("{tabs}   Rectangle dimensions {leRect.Width}x{leRect.Height}, X={leRect.X} Y={leRect.Y}");
        System.out.println("{tabs}   Page {layoutEnumerator.PageIndex}");
    }
    //ExEnd

    //ExStart
    //ExFor:IPageLayoutCallback
    //ExFor:IPageLayoutCallback.Notify(PageLayoutCallbackArgs)
    //ExFor:PageLayoutCallbackArgs.Event
    //ExFor:PageLayoutCallbackArgs.Document
    //ExFor:PageLayoutCallbackArgs.PageIndex
    //ExFor:PageLayoutEvent
    //ExSummary:Shows how to track layout changes with a layout callback.
    @Test
    public void pageLayoutCallback() throws Exception
    {
        Document doc = new Document();
        doc.getBuiltInDocumentProperties().setTitle("My Document");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        doc.getLayoutOptions().setCallback(new RenderPageLayoutCallback());
        doc.updatePageLayout();

        doc.save(getArtifactsDir() + "Layout.PageLayoutCallback.pdf");
    }

    /// <summary>
    /// Notifies us when we save the document to a fixed page format
    /// and renders a page that we perform a page reflow on to an image in the local file system.
    /// </summary>
    private static class RenderPageLayoutCallback implements IPageLayoutCallback
    {
        public void notify(PageLayoutCallbackArgs a) throws Exception
        {
            switch (a.getEvent())
            {
                case PageLayoutEvent.PART_REFLOW_FINISHED:
                    notifyPartFinished(a);
                    break;
                case PageLayoutEvent.CONVERSION_FINISHED:
                    notifyConversionFinished(a);
                    break;
            }
        }

        private void notifyPartFinished(PageLayoutCallbackArgs a) throws Exception
        {
            System.out.println("Part at page {a.PageIndex + 1} reflow.");
            renderPage(a, a.getPageIndex());
        }

        private void notifyConversionFinished(PageLayoutCallbackArgs a)
        {
            System.out.println("Document \"{a.Document.BuiltInDocumentProperties.Title}\" converted to page format.");
        }

        private void renderPage(PageLayoutCallbackArgs a, int pageIndex) throws Exception
        {
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.PNG); { saveOptions.setPageSet(new PageSet(pageIndex)); }

            FileStream stream =
                new FileStream(getArtifactsDir() + $"PageLayoutCallback.page-{pageIndex + 1} {++mNum}.png",
                    FileMode.CREATE);
            try /*JAVA: was using*/
        	{
                a.getDocument().save(stream, saveOptions);
        	}
            finally { if (stream != null) stream.close(); }
        }

        private int mNum;
    }
    //ExEnd

    @Test
    public void restartPageNumberingInContinuousSection() throws Exception
    {
        //ExStart
        //ExFor:LayoutOptions.ContinuousSectionPageNumberingRestart
        //ExFor:ContinuousSectionRestart
        //ExSummary:Shows how to control page numbering in a continuous section.
        Document doc = new Document(getMyDir() + "Continuous section page numbering.docx");

        // By default Aspose.Words behavior matches the Microsoft Word 2019.
        // If you need old Aspose.Words behavior, repetitive Microsoft Word 2016, use 'ContinuousSectionRestart.FromNewPageOnly'.
        // Page numbering restarts only if there is no other content before the section on the page where the section starts,
        // because of that the numbering will reset to 2 from the second page.
        doc.getLayoutOptions().setContinuousSectionPageNumberingRestart(ContinuousSectionRestart.FROM_NEW_PAGE_ONLY);
        doc.updatePageLayout();

        doc.save(getArtifactsDir() + "Layout.RestartPageNumberingInContinuousSection.pdf");
        //ExEnd
    }
}

