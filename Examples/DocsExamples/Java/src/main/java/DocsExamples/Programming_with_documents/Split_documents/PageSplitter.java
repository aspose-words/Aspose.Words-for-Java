package DocsExamples.Programming_with_documents.Split_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.testng.annotations.Test;

import java.io.File;
import java.text.MessageFormat;
import java.util.*;
import java.util.regex.Pattern;

@Test
public class PageSplitter extends DocsExamplesBase
{
    @Test
    public void splitDocuments() throws Exception
    {
        splitAllDocumentsToPages(getMyDir());
    }

    private void splitDocumentToPages(String docName) throws Exception
    {
        String fileName = FilenameUtils.getBaseName(docName);
        String extensionName = FilenameUtils.getExtension(docName);

        System.out.println("Processing document: " + fileName + "." + extensionName);

        Document doc = new Document(docName);

        // Split nodes in the document into separate pages.
        DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

        // Save each page to the disk as a separate document.
        for (int page = 1; page <= doc.getPageCount(); page++)
        {
            Document pageDoc = splitter.getDocumentOfPage(page);
            pageDoc.save(getArtifactsDir()+ MessageFormat.format("{0} - page{1}.{2}", fileName, page, extensionName));
        }
    }

    private void splitAllDocumentsToPages(String folderName) throws Exception
    {
        ArrayList<String> fileNames = directoryGetFiles(folderName, "*.doc");

        for (String fileName : fileNames)
        {
            splitDocumentToPages(fileName);
        }
    }

    static ArrayList<String> directoryGetFiles(final String dirname, final String filenamePattern) {
        File dirFile = new File(dirname);
        Pattern re = Pattern.compile(filenamePattern.replace("*", ".*").replace("?", ".?"));
        ArrayList<String> dirFiles = new ArrayList<>();
        for (File file : dirFile.listFiles()) {
            if (file.isDirectory()) {
                dirFiles.addAll(directoryGetFiles(file.getPath(), filenamePattern));
            } else {
                if (re.matcher(file.getName()).matches()) {
                    dirFiles.add(file.getPath());
                }
            }
        }
        return dirFiles;
    }
}

/// <summary>
/// Splits a document into multiple documents, one per page.
/// </summary>
class DocumentPageSplitter
{
    private PageNumberFinder pageNumberFinder;

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

    private Document getDocument() {
        return pageNumberFinder.getDocument();
    }

    /// <summary>
    /// Gets the document of a page.
    /// </summary>
    /// <param name="pageIndex">
    /// 1-based index of a page.
    /// </param>
    /// <returns>
    /// The <see cref="Document"/>.
    /// </returns>
    public Document getDocumentOfPage(int pageIndex) throws Exception {
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
    public Document getDocumentOfPageRange(int startIndex, int endIndex) throws Exception {
        Document result = (Document) getDocument().deepClone(false);
        for (Node section : pageNumberFinder.retrieveAllNodesOnPages(startIndex, endIndex, NodeType.SECTION))
        {
            result.appendChild(result.importNode(section, true));
        }

        return result;
    }
}

/// <summary>
/// Provides methods for extracting nodes of a document which are rendered on a specified pages.
/// </summary>
class PageNumberFinder
{
    // Maps node to a start/end page numbers.
    // This is used to override baseline page numbers provided by the collector when the document is split.
    private Map<Node, Integer> nodeStartPageLookup = new HashMap<>();
    private Map<Node, Integer> nodeEndPageLookup = new HashMap<>();
    private LayoutCollector collector;

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

    public Document getDocument()
    {
        return collector.getDocument();
    }

    /// <summary>
    /// Retrieves 1-based index of a page that the node begins on.
    /// </summary>
    /// <param name="node">
    /// The node.
    /// </param>
    /// <returns>
    /// Page index.
    /// </returns>
    public int getPage(Node node) throws Exception {
        return nodeStartPageLookup.containsKey(node)
            ? nodeStartPageLookup.get(node)
            : collector.getStartPageIndex(node);
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
    public int getPageEnd(Node node) throws Exception {
        return nodeEndPageLookup.containsKey(node)
            ? nodeEndPageLookup.get(node)
            : collector.getEndPageIndex(node);
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
    public int pageSpan(Node node) throws Exception {
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
        if (startPage < 1 || startPage > collector.getDocument().getPageCount())
        {
            throw new IllegalStateException("'startPage' is out of range");
        }

        if (endPage < 1 || endPage > collector.getDocument().getPageCount() || endPage < startPage)
        {
            throw new IllegalStateException("'endPage' is out of range");
        }

        checkPageListsPopulated();

        ArrayList<Node> pageNodes = new ArrayList<>();
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
        for (Paragraph paragraph : (Iterable<Paragraph>) collector.getDocument().getChildNodes(NodeType.PARAGRAPH, true))
        {
            if (getPage(paragraph) != getPageEnd(paragraph))
            {
                splitRunsByWords(paragraph);
            }
        }

        clearCollector();

        // Visit any composites which are possibly split across pages and split them into separate nodes.
        collector.getDocument().accept(new SectionSplitter(this));
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

    private void checkPageListsPopulated() throws Exception {
        if (reversePageLookup != null)
        {
            return;
        }

        reversePageLookup = new HashMap<Integer, ArrayList<Node>>();

        // Add each node to a list that represent the nodes found on each page.
        for (Node node : (Iterable<Node>) collector.getDocument().getChildNodes(NodeType.ANY, true))
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
                    reversePageLookup.put(page, new ArrayList<Node>());
                }

                reversePageLookup.get(page).add(node);
            }
        }
    }

    private void splitRunsByWords(Paragraph paragraph) throws Exception {
        for (Run run : paragraph.getRuns())
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
        String[] words = reverseWord(run.getText());

        for (String word : words)
        {
            int pos = run.getText().length() - word.length() - 1;
            if (pos > 1)
            {
                splitRun(run, run.getText().length() - word.length() - 1);
            }
        }
    }

    private static String[] reverseWord(String str) {
        String words[] = str.split(" ");
        String reverseWord = "";

        for (String w : words) {
            StringBuilder sb = new StringBuilder(w);
            sb.reverse();
            reverseWord += sb.toString() + " ";
        }

        return reverseWord.split(" ");
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
        collector.clear();
        collector.getDocument().updatePageLayout();

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
    private PageNumberFinder pageNumberFinder;

    public SectionSplitter(PageNumberFinder pageNumberFinder)
    {
        this.pageNumberFinder = pageNumberFinder;
    }

    public int visitParagraphStart(Paragraph paragraph) throws Exception {
        return continueIfCompositeAcrossPageElseSkip(paragraph);
    }

    public int visitTableStart(Table table) throws Exception {
        return continueIfCompositeAcrossPageElseSkip(table);
    }

    public int visitRowStart(Row row) throws Exception {
        return continueIfCompositeAcrossPageElseSkip(row);
    }

    public int visitCellStart(Cell cell) throws Exception {
        return continueIfCompositeAcrossPageElseSkip(cell);
    }

    public int visitStructuredDocumentTagStart(StructuredDocumentTag sdt) throws Exception {
        return continueIfCompositeAcrossPageElseSkip(sdt);
    }

    public int visitSmartTagStart(SmartTag smartTag) throws Exception {
        return continueIfCompositeAcrossPageElseSkip(smartTag);
    }

    public int visitSectionStart(Section section) throws Exception {
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

    public int visitSmartTagEnd(SmartTag smartTag) throws Exception {
        splitComposite(smartTag);
        return VisitorAction.CONTINUE;
    }

    public int visitStructuredDocumentTagEnd(StructuredDocumentTag sdt) throws Exception {
        splitComposite(sdt);
        return VisitorAction.CONTINUE;
    }

    public int visitCellEnd(Cell cell) throws Exception {
        splitComposite(cell);
        return VisitorAction.CONTINUE;
    }

    public int visitRowEnd(Row row) throws Exception {
        splitComposite(row);
        return VisitorAction.CONTINUE;
    }

    public int visitTableEnd(Table table) throws Exception {
        splitComposite(table);
        return VisitorAction.CONTINUE;
    }

    public int visitParagraphEnd(Paragraph paragraph) throws Exception {
        // If the paragraph contains only section break, add fake run into.
        if (paragraph.isEndOfSection() && paragraph.getChildNodes().getCount() == 1 &&
            "\f".equals(paragraph.getChildNodes().get(0).getText()))
        {
            Run run = new Run(paragraph.getDocument());
            paragraph.appendChild(run);
            int currentEndPageNum = pageNumberFinder.getPageEnd(paragraph);
            pageNumberFinder.addPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
        }

        for (Node cloneNode : splitComposite(paragraph))
        {
            Paragraph clonePara = (Paragraph) cloneNode;
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

    public int visitSectionEnd(Section section) throws Exception {
        for (Node cloneNode : splitComposite(section))
        {
            Section cloneSection = (Section) cloneNode;

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

    private /*VisitorAction*/int continueIfCompositeAcrossPageElseSkip(CompositeNode composite) throws Exception {
        return pageNumberFinder.pageSpan(composite) > 1
            ? VisitorAction.CONTINUE
            : VisitorAction.SKIP_THIS_NODE;
    }

    private ArrayList<Node> splitComposite(CompositeNode composite) throws Exception {
        ArrayList<Node> splitNodes = new ArrayList<>();
        for (Node splitNode : findChildSplitPositions(composite))
        {
            splitNodes.add(splitCompositeAtNode(composite, splitNode));
        }

        return splitNodes;
    }

    private Iterable<Node> findChildSplitPositions(CompositeNode node) throws Exception {
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

    private CompositeNode splitCompositeAtNode(CompositeNode baseNode, Node targetNode) throws Exception {
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

        Body lastBody = (Body) Arrays.stream(new Iterator[]{section.getChildNodes().iterator()}).reduce((first, second) -> second)
            .orElse(null);

        RunCollection runs = (RunCollection) lastBody.getChildNodes(NodeType.RUN, true).iterator();
        Run run  = Arrays.stream(runs.toArray()).filter(p -> p.getText().endsWith(PAGE_BREAK_STR)).findFirst().get();

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
            run.setText(StringUtils.stripEnd(run.getText(), String.valueOf(PAGE_BREAK)));
        }

        if (paragraph.getChildNodes().getCount() == 0)
        {
            CompositeNode parent = paragraph.getParentNode();
            parent.removeChild(paragraph);
        }
    }
}
