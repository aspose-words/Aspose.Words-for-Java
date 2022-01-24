package DocsExamples.Programming_with_Documents.Split_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.ms.System.IO.Path;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Document;
import java.util.ArrayList;
import com.aspose.ms.System.IO.Directory;
import com.aspose.ms.System.IO.SearchOption;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import java.util.Map;
import java.util.HashMap;
import com.aspose.words.LayoutCollector;
import com.aspose.words.Paragraph;
import com.aspose.ms.System.Collections.msDictionary;
import com.aspose.words.Run;
import com.aspose.ms.System.msString;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.Table;
import com.aspose.words.Row;
import com.aspose.words.Cell;
import com.aspose.words.StructuredDocumentTag;
import com.aspose.words.SmartTag;
import com.aspose.words.Section;
import com.aspose.words.HeaderFooterCollection;
import com.aspose.words.HeaderFooter;
import com.aspose.words.SectionStart;
import com.aspose.words.CompositeNode;
import java.util.Collections;
import com.aspose.words.Body;


class PageSplitter extends DocsExamplesBase
{
    @Test
    public void splitDocuments() throws Exception
    {
        splitAllDocumentsToPages(getMyDir());
    }

    public void splitDocumentToPages(String docName) throws Exception
    {
        String fileName = Path.getFileNameWithoutExtension(docName);
        String extensionName = Path.getExtension(docName);

        System.out.println("Processing document: " + fileName + extensionName);

        Document doc = new Document(docName);

        // Split nodes in the document into separate pages.
        DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

        // Save each page to the disk as a separate document.
        for (int page = 1; page <= doc.getPageCount(); page++)
        {
            Document pageDoc = splitter.getDocumentOfPage(page);
            pageDoc.save(Path.combine(getArtifactsDir(),
                $"{fileName} - page{page} Out{extensionName}"));
        }
    }

    public void splitAllDocumentsToPages(String folderName) throws Exception
    {
        ArrayList<String> fileNames = Directory.getFiles(folderName, "*.doc", SearchOption.TOP_DIRECTORY_ONLY)
            .Where(item => item.EndsWith(".doc")).ToList();

        for (String fileName : fileNames)
        {
            splitDocumentToPages(fileName);
        }
    }
}

/// <summary>
/// Splits a document into multiple documents, one per page.
/// </summary>
class DocumentPageSplitter
{
    private /*final*/ PageNumberFinder pageNumberFinder;

    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentPageSplitter"/> class.
    /// This method splits the document into sections so that each page begins and ends at a section boundary.
    /// It is recommended not to modify the document afterwards.
    /// </summary>
    /// <param name="source">Source document</param>
    public DocumentPageSplitter(Document source) throws Exception
    {
        pageNumberFinder = PageNumberFinderFactory.create(source);
    }

    /// <summary>
    /// Gets the document this instance works with.
    /// </summary>
    private Document Document => private pageNumberFinder.DocumentpageNumberFinder;

    /// <summary>
    /// Gets the document of a page.
    /// </summary>
    /// <param name="pageIndex">
    /// 1-based index of a page.
    /// </param>
    /// <returns>
    /// The <see cref="Document"/>.
    /// </returns>
    public Document getDocumentOfPage(int pageIndex)
    {
        return getDocumentOfPageRange(pageIndex, pageIndex);
    }

    /// <summary>
    /// Gets the document of a page range.
    /// </summary>
    /// <param name="startIndex">
    /// 1-based index of the start page.
    /// </param>
    /// <param name="endIndex">
    /// 1-based index of the end page.
    /// </param>
    /// <returns>
    /// The <see cref="Document"/>.
    /// </returns>
    public Document getDocumentOfPageRange(int startIndex, int endIndex)
    {
        Document result = (Document) Document.deepClone(false);
        for (Node section : pageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex,
            NodeType.SECTION) !!Autoporter error: Undefined expression type )
        {
            result.appendChild(result.importNode(section, true));
        }

        return result;
    }
}

/// <summary>
/// Provides methods for extracting nodes of a document which are rendered on a specified pages.
/// </summary>
public class PageNumberFinder
{
    // Maps node to a start/end page numbers.
    // This is used to override baseline page numbers provided by the collector when the document is split.
    private /*final*/ Map<Node, Integer> nodeStartPageLookup = new HashMap<Node, Integer>();
    private /*final*/ Map<Node, Integer> nodeEndPageLookup = new HashMap<Node, Integer>();
    private /*final*/ LayoutCollector collector;

    // Maps page number to a list of nodes found on that page.
    private Map<Integer, ArrayList<Node>> reversePageLookup;

    /// <summary>
    /// Initializes a new instance of the <see cref="PageNumberFinder"/> class.
    /// </summary>
    /// <param name="collector">A collector instance that has layout model records for the document.</param>
    public PageNumberFinder(LayoutCollector collector)
    {
        this.collector = collector;
    }

    /// <summary>
    /// Gets the document this instance works with.
    /// </summary>
    public Document Document => private collector.Documentcollector;

    /// <summary>
    /// Retrieves 1-based index of a page that the node begins on.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    public int getPage(Node node)
    {
        return nodeStartPageLookup.containsKey(node)
            ? nodeStartPageLookup.get(node)
            : collector.GetStartPageIndex(node);
    }

    /// <summary>
    /// Retrieves 1-based index of a page that the node ends on.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    public int getPageEnd(Node node)
    {
        return nodeEndPageLookup.containsKey(node)
            ? nodeEndPageLookup.get(node)
            : collector.GetEndPageIndex(node);
    }

    /// <summary>
    /// Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    public int pageSpan(Node node)
    {
        return getPageEnd(node) - getPage(node) + 1;
    }

    /// <summary>
    /// Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
    /// </summary>
    /// <param name="startPage">
    /// The start Page.
    /// </param>
    /// <param name="endPage">
    /// The end Page.
    /// </param>
    /// <param name="nodeType">
    /// The node Type.
    /// </param>
    /// <returns>
    /// The <see cref="IList{T}"/>.
    /// </returns>
    public ArrayList<Node> retrieveAllNodesOnPages(int startPage, int endPage, /*NodeType*/int nodeType) throws Exception
    {
        if (startPage < 1 || startPage > Document.getPageCount())
        {
            throw new IllegalStateException("'startPage' is out of range");
        }

        if (endPage < 1 || endPage > Document.getPageCount() || endPage < startPage)
        {
            throw new IllegalStateException("'endPage' is out of range");
        }

        checkPageListsPopulated();

        ArrayList<Node> pageNodes = new ArrayList<Node>();
        for (int page = startPage; page <= endPage; page++)
        {
            // Some pages can be empty.
            if (!reversePageLookup.containsKey(page))
            {
                continue;
            }

            for (Node node : reversePageLookup.get(page))
            {
                if (node.getParentNode() != null
                    && (nodeType == NodeType.ANY || node.getNodeType() == nodeType)
                    && !pageNodes.contains(node))
                {
                    pageNodes.add(node);
                }
            }
        }

        return pageNodes;
    }

    /// <summary>
    /// Splits nodes that appear over two or more pages into separate nodes so that they still appear in the same way
    /// but no longer appear across a page.
    /// </summary>
    public void splitNodesAcrossPages() throws Exception
    {
        for (Paragraph paragraph : (Iterable<Paragraph>) Document.getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (getPage(paragraph) != getPageEnd(paragraph))
            {
                splitRunsByWords(paragraph);
            }
        }

        clearCollector();

        // Visit any composites which are possibly split across pages and split them into separate nodes.
        Document.accept(new SectionSplitter(this));
    }

    /// <summary>
    /// This is called by <see cref="SectionSplitter"/> to update page numbers of split nodes.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <param name="startPage">
    /// The start Page.
    /// </param>
    /// <param name="endPage">
    /// The end Page.
    /// </param>
    void addPageNumbersForNode(Node node, int startPage, int endPage)
    {
        if (startPage > 0)
        {
            nodeStartPageLookup.put(node, startPage);
        }

        if (endPage > 0)
        {
            nodeEndPageLookup.put(node, endPage);
        }
    }

    private boolean isHeaderFooterType(Node node)
    {
        return node.getNodeType() == NodeType.HEADER_FOOTER || node.getAncestor(NodeType.HEADER_FOOTER) != null;
    }

    private void checkPageListsPopulated()
    {
        if (reversePageLookup != null)
        {
            return;
        }

        reversePageLookup = new HashMap<Integer, ArrayList<Node>>();

        // Add each node to a list that represent the nodes found on each page.
        for (Node node : (Iterable<Node>) Document.getChildNodes(NodeType.ANY, true))
        {
            // Headers/Footers follow sections and are not split by themselves.
            if (isHeaderFooterType(node))
            {
                continue;
            }

            int startPage = getPage(node);
            int endPage = getPageEnd(node);
            for (int page = startPage; page <= endPage; page++)
            {
                if (!reversePageLookup.containsKey(page))
                {
                    msDictionary.add(reversePageLookup, page, new ArrayList<Node>());
                }

                reversePageLookup.get(page).add(node);
            }
        }
    }

    private void splitRunsByWords(Paragraph paragraph)
    {
        for (Run run : (Iterable<Run>) paragraph.getRuns())
        {
            if (getPage(run) == getPageEnd(run))
            {
                continue;
            }

            splitRunByWords(run);
        }
    }

    private void splitRunByWords(Run run)
    {
        Iterable<String> words = msString.split(run.getText(), ' ').Reverse();

        for (String word : words)
        {
            int pos = run.getText().length() - word.length() - 1;
            if (pos > 1)
            {
                splitRun(run, run.getText().length() - word.length() - 1);
            }
        }
    }

    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    private void splitRun(Run run, int position)
    {
        Run afterRun = (Run) run.deepClone(true);
        afterRun.setText(run.getText().substring(position));
        run.setText(run.getText().substring((0), (0) + (position)));
        run.getParentNode().insertAfter(afterRun, run);
    }

    private void clearCollector() throws Exception
    {
        collector.Clear();
        Document.updatePageLayout();

        nodeStartPageLookup.clear();
        nodeEndPageLookup.clear();
    }
}

class PageNumberFinderFactory
{
    public static PageNumberFinder create(Document document) throws Exception
    {
        LayoutCollector layoutCollector = new LayoutCollector(document);
        document.updatePageLayout();
        PageNumberFinder pageNumberFinder = new PageNumberFinder(layoutCollector);
        pageNumberFinder.splitNodesAcrossPages();
        return pageNumberFinder;
    }
}

/// <summary>
/// Splits a document into multiple sections so that each page begins and ends at a section boundary.
/// </summary>
class SectionSplitter extends DocumentVisitor
{
    private /*final*/ PageNumberFinder pageNumberFinder;

    public SectionSplitter(PageNumberFinder pageNumberFinder)
    {
        this.pageNumberFinder = pageNumberFinder;
    }

    public /*override*/ /*VisitorAction*/int visitParagraphStart(Paragraph paragraph)
    {
        return continueIfCompositeAcrossPageElseSkip(paragraph);
    }

    public /*override*/ /*VisitorAction*/int visitTableStart(Table table)
    {
        return continueIfCompositeAcrossPageElseSkip(table);
    }

    public /*override*/ /*VisitorAction*/int visitRowStart(Row row)
    {
        return continueIfCompositeAcrossPageElseSkip(row);
    }

    public /*override*/ /*VisitorAction*/int visitCellStart(Cell cell)
    {
        return continueIfCompositeAcrossPageElseSkip(cell);
    }

    public /*override*/ /*VisitorAction*/int visitStructuredDocumentTagStart(StructuredDocumentTag sdt)
    {
        return continueIfCompositeAcrossPageElseSkip(sdt);
    }

    public /*override*/ /*VisitorAction*/int visitSmartTagStart(SmartTag smartTag)
    {
        return continueIfCompositeAcrossPageElseSkip(smartTag);
    }

    public /*override*/ /*VisitorAction*/int visitSectionStart(Section section)
    {
        Section previousSection = (Section) section.getPreviousSibling();

        // If there is a previous section, attempt to copy any linked header footers.
        // Otherwise, they will not appear in an extracted document if the previous section is missing.
        if (previousSection != null)
        {
            HeaderFooterCollection previousHeaderFooters = previousSection.getHeadersFooters();
            if (!section.getPageSetup().getRestartPageNumbering())
            {
                section.getPageSetup().setRestartPageNumbering(true);
                section.getPageSetup().setPageStartingNumber(previousSection.getPageSetup().getPageStartingNumber() +
                                                       pageNumberFinder.pageSpan(previousSection));
            }

            for (HeaderFooter previousHeaderFooter : (Iterable<HeaderFooter>) previousHeaderFooters)
            {
                if (section.getHeadersFooters().getByHeaderFooterType(previousHeaderFooter.getHeaderFooterType()) == null)
                {
                    HeaderFooter newHeaderFooter =
                        (HeaderFooter) previousHeaderFooters.getByHeaderFooterType(previousHeaderFooter.getHeaderFooterType()).deepClone(true);
                    section.getHeadersFooters().add(newHeaderFooter);
                }
            }
        }

        return continueIfCompositeAcrossPageElseSkip(section);
    }

    public /*override*/ /*VisitorAction*/int visitSmartTagEnd(SmartTag smartTag)
    {
        splitComposite(smartTag);
        return VisitorAction.CONTINUE;
    }

    public /*override*/ /*VisitorAction*/int visitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
    {
        splitComposite(sdt);
        return VisitorAction.CONTINUE;
    }

    public /*override*/ /*VisitorAction*/int visitCellEnd(Cell cell)
    {
        splitComposite(cell);
        return VisitorAction.CONTINUE;
    }

    public /*override*/ /*VisitorAction*/int visitRowEnd(Row row)
    {
        splitComposite(row);
        return VisitorAction.CONTINUE;
    }

    public /*override*/ /*VisitorAction*/int visitTableEnd(Table table)
    {
        splitComposite(table);
        return VisitorAction.CONTINUE;
    }

    public /*override*/ /*VisitorAction*/int visitParagraphEnd(Paragraph paragraph)
    {
        // If the paragraph contains only section break, add fake run into.
        if (paragraph.isEndOfSection() && paragraph.getChildNodes().getCount() == 1 &&
            "\f".equals(paragraph.getChildNodes().get(0).getText()))
        {
            Run run = new Run(paragraph.getDocument());
            paragraph.appendChild(run);
            int currentEndPageNum = pageNumberFinder.getPageEnd(paragraph);
            pageNumberFinder.addPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
        }

        for (Paragraph clonePara : (Iterable<Paragraph>) splitComposite(paragraph))
        {
            // Remove list numbering from the cloned paragraph but leave the indent the same 
            // as the paragraph is supposed to be part of the item before.
            if (paragraph.isListItem())
            {
                double textPosition = clonePara.getListFormat().getListLevel().getTextPosition();
                clonePara.getListFormat().removeNumbers();
                clonePara.getParagraphFormat().setLeftIndent(textPosition);
            }

            // Reset spacing of split paragraphs in tables as additional spacing may cause them to look different.
            if (paragraph.isInCell())
            {
                clonePara.getParagraphFormat().setSpaceBefore(0.0);
                paragraph.getParagraphFormat().setSpaceAfter(0.0);
            }
        }

        return VisitorAction.CONTINUE;
    }

    public /*override*/ /*VisitorAction*/int visitSectionEnd(Section section)
    {
        for (Section cloneSection : (Iterable<Section>) splitComposite(section))
        {
            cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
            cloneSection.getPageSetup().setRestartPageNumbering(true);
            cloneSection.getPageSetup().setPageStartingNumber(section.getPageSetup().getPageStartingNumber() +
                                                        (section.getDocument().indexOf(cloneSection) -
                                                         section.getDocument().indexOf(section)));
            cloneSection.getPageSetup().setDifferentFirstPageHeaderFooter(false);

            // Corrects page break at the end of the section.
            SplitPageBreakCorrector.processSection(cloneSection);
        }

        SplitPageBreakCorrector.processSection(section);

        // Add new page numbering for the body of the section as well.
        pageNumberFinder.addPageNumbersForNode(section.getBody(), pageNumberFinder.getPage(section),
            pageNumberFinder.getPageEnd(section));
        return VisitorAction.CONTINUE;
    }

    private /*VisitorAction*/int continueIfCompositeAcrossPageElseSkip(CompositeNode composite)
    {
        return pageNumberFinder.pageSpan(composite) > 1
            ? VisitorAction.CONTINUE
            : VisitorAction.SKIP_THIS_NODE;
    }

    private ArrayList<Node> splitComposite(CompositeNode composite)
    {
        ArrayList<Node> splitNodes = new ArrayList<Node>();
        for (Node splitNode : findChildSplitPositions(composite))
        {
            splitNodes.add(splitCompositeAtNode(composite, splitNode));
        }

        return splitNodes;
    }

    private Iterable<Node> findChildSplitPositions(CompositeNode node)
    {
        // A node may span across multiple pages, so a list of split positions is returned.
        // The split node is the first node on the next page.
        ArrayList<Node> splitList = new ArrayList<Node>();

        int startingPage = pageNumberFinder.getPage(node);
        
        Node[] childNodes = node.getNodeType() == NodeType.SECTION
            ? ((Section) node).getBody().getChildNodes().toArray()
            : node.getChildNodes().toArray();
        for (Node childNode : childNodes)
        {
            int pageNum = pageNumberFinder.getPage(childNode);

            if (childNode instanceof Run)
            {
                pageNum = pageNumberFinder.getPageEnd(childNode);
            }

            // If the page of the child node has changed, then this is the split position.
            // Add this to the list.
            if (pageNum > startingPage)
            {
                splitList.add(childNode);
                startingPage = pageNum;
            }

            if (pageNumberFinder.pageSpan(childNode) > 1)
            {
                pageNumberFinder.addPageNumbersForNode(childNode, pageNum, pageNum);
            }
        }

        // Split composites backward, so the cloned nodes are inserted in the right order.
        Collections.reverse(splitList);
        return splitList;
    }

    private CompositeNode splitCompositeAtNode(CompositeNode baseNode, Node targetNode)
    {
        CompositeNode cloneNode = (CompositeNode) baseNode.deepClone(false);
        Node node = targetNode;
        int currentPageNum = pageNumberFinder.getPage(baseNode);

        // Move all nodes found on the next page into the copied node. Handle row nodes separately.
        if (baseNode.getNodeType() != NodeType.ROW)
        {
            CompositeNode composite = cloneNode;
            if (baseNode.getNodeType() == NodeType.SECTION)
            {
                cloneNode = (CompositeNode) baseNode.deepClone(true);
                Section section = (Section) cloneNode;
                section.getBody().removeAllChildren();
                composite = section.getBody();
            }

            while (node != null)
            {
                Node nextNode = node.getNextSibling();
                composite.appendChild(node);
                node = nextNode;
            }
        }
        else
        {
            // If we are dealing with a row, we need to add dummy cells for the cloned row.
            int targetPageNum = pageNumberFinder.getPage(targetNode);
            
            Node[] childNodes = baseNode.getChildNodes().toArray();
            for (Node childNode : childNodes)
            {
                int pageNum = pageNumberFinder.getPage(childNode);
                if (pageNum == targetPageNum)
                {
                    if (cloneNode.getNodeType() == NodeType.ROW)
                        ((Row) cloneNode).ensureMinimum();

                    if (cloneNode.getNodeType() == NodeType.CELL)
                        ((Cell) cloneNode).ensureMinimum();

                    cloneNode.getLastChild().remove();
                    cloneNode.appendChild(childNode);
                }
                else if (pageNum == currentPageNum)
                {
                    cloneNode.appendChild(childNode.deepClone(false));
                    if (cloneNode.getLastChild().getNodeType() != NodeType.CELL)
                    {
                        ((CompositeNode) cloneNode.getLastChild()).appendChild(
                            ((CompositeNode) childNode).getFirstChild().deepClone(false));
                    }
                }
            }
        }

        // Insert the split node after the original.
        baseNode.getParentNode().insertAfter(cloneNode, baseNode);

        // Update the new page numbers of the base node and the cloned node, including its descendants.
        // This will only be a single page as the cloned composite is split to be on one page.
        int currentEndPageNum = pageNumberFinder.getPageEnd(baseNode);
        pageNumberFinder.addPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
        pageNumberFinder.addPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);
        for (Node childNode : (Iterable<Node>) cloneNode.getChildNodes(NodeType.ANY, true))
        {
            pageNumberFinder.addPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
        }

        return cloneNode;
    }
}

class SplitPageBreakCorrector
{
    private static final String PAGE_BREAK_STR = "\f";
    private static final char PAGE_BREAK = '\f';

    public static void processSection(Section section)
    {
        if (section.getChildNodes().getCount() == 0)
        {
            return;
        }

        Body lastBody = section.getChildNodes().<Body>OfType().LastOrDefault();

        Run run = lastBody?.GetChildNodes(NodeType.Run, true).OfType<Run>()
            .FirstOrDefault(p => p.Text.EndsWith(PageBreakStr));

        if (run != null)
        {
            removePageBreak(run);
        }
    }

    public void removePageBreakFromParagraph(Paragraph paragraph)
    {
        Run run = (Run) paragraph.getFirstChild();
        if (PAGE_BREAK_STR.equals(run.getText()))
        {
            paragraph.removeChild(run);
        }
    }

    private void processLastParagraph(Paragraph paragraph)
    {
        Node lastNode = paragraph.getChildNodes().get(paragraph.getChildNodes().getCount() - 1);
        if (lastNode.getNodeType() != NodeType.RUN)
        {
            return;
        }

        Run run = (Run) lastNode;
        removePageBreak(run);
    }

    private static void removePageBreak(Run run)
    {
        Paragraph paragraph = run.getParentParagraph();
        
        if (PAGE_BREAK_STR.equals(run.getText()))
        {
            paragraph.removeChild(run);
        }
        else if (run.getText().endsWith(PAGE_BREAK_STR))
        {
            run.setText(msString.trimEnd(run.getText(), PAGE_BREAK));
        }

        if (paragraph.getChildNodes().getCount() == 0)
        {
            CompositeNode parent = paragraph.getParentNode();
            parent.removeChild(paragraph);
        }
    }
}
