package com.aspose.words.examples.loading_saving;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.io.File;
import java.text.MessageFormat;
import java.util.*;
import java.util.List;

public class PageSplitter {
    public static void main(String[] args) throws Exception {
        License lic = new License();
        lic.setLicense("D:\\Aspose\\Licenses\\aspose.total.java.lic");

        //ExStart:PageSplitter
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(PageSplitter.class);

        SplitAllDocumentsToPages(dataDir);
        //ExEnd:PageSplitter
        System.out.println("\nDocument split to pages successfully.\nFile saved at " + dataDir + "\\Out");
    }

    //ExStart:SplitDocumentToPages
    public static void SplitDocumentToPages(File docName) throws Exception {
        String folderName = docName.getParent();
        String fileName = docName.getName();
        String extensionName = fileName.substring(fileName.lastIndexOf("."));
        String outFolder = new File(folderName, "Out").getAbsolutePath();
        System.out.println("Processing document: " + fileName);

        Document doc = new Document(docName.getAbsolutePath());

        // Split nodes in the document into separate pages.
        DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

        // Save each page to the disk as a separate document.
        for (int page = 1; page <= doc.getPageCount(); page++) {
            Document pageDoc = splitter.getDocumentOfPage(page);
            pageDoc.save(new File(outFolder, MessageFormat.format("{0} - page{1} Out{2}", fileName, page, extensionName)).getAbsolutePath());
        }
    }
    //ExEnd:SplitDocumentToPages

    //ExStart:SplitAllDocumentsToPages
    public static void SplitAllDocumentsToPages(String folderName) throws Exception {
        File[] files = new File(folderName).listFiles();

        for (File file : files) {
            if (file.isFile()) {
                SplitDocumentToPages(file);
            }
        }
    }
    //ExEnd:SplitAllDocumentsToPages
}

//ExStart:DocumentPageSplitter
class DocumentPageSplitter {
    private PageNumberFinder pageNumberFinder;

    /// <summary>
    /// Initializes a new instance of the <see cref="DocumentPageSplitter"/> class.
    /// This method splits the document into sections so that each page begins and ends at a section boundary.
    /// It is recommended not to modify the document afterwards.
    /// </summary>
    /// <param name="source">source document</param>
    public DocumentPageSplitter(Document source) throws Exception {
        this.pageNumberFinder = PageNumberFinderFactory.create(source);
    }

    /// <summary>
    /// Gets the document this instance works with.
    /// </summary>
    private Document getDocument() {
        return this.pageNumberFinder.getDocument();
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
        return this.getDocumentOfPageRange(pageIndex, pageIndex);
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
        Document result = (Document) this.getDocument().deepClone(false);
        for (Section section : (Iterable<Section>) this.pageNumberFinder.retrieveAllNodesOnPages(startIndex, endIndex, NodeType.SECTION)) {
            result.appendChild(result.importNode(section, true));
        }

        return result;
    }
}

class PageNumberFinder {
    // Maps node to a start/end page numbers. This is used to override baseline page numbers provided by collector when document is split.
    private Hashtable nodeStartPageLookup = new Hashtable();
    private Hashtable nodeEndPageLookup = new Hashtable();
    private LayoutCollector collector;

    // Maps page number to a list of nodes found on that page.
    private Hashtable reversePageLookup;

    /// <summary>
    /// Initializes a new instance of the <see cref="PageNumberFinder"/> class.
    /// </summary>
    /// <param name="collector">A collector instance which has layout model records for the document.</param>
    public PageNumberFinder(LayoutCollector collector) {
        this.collector = collector;
    }

    /// <summary>
    /// Gets the document this instance works with.
    /// </summary>
    public Document getDocument() {
        return this.collector.getDocument();
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
        return this.nodeStartPageLookup.containsKey(node) ?
                (Integer) this.nodeStartPageLookup.get(node) : this.collector.getStartPageIndex(node);
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
        return this.nodeEndPageLookup.containsKey(node) ?
                (Integer) this.nodeEndPageLookup.get(node) :
                this.collector.getEndPageIndex(node);
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
        return this.getPageEnd(node) - this.getPage(node) + 1;
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
    /// The <see cref="IList"/>.
    /// </returns>
    public ArrayList retrieveAllNodesOnPages(int startPage, int endPage, int nodeType) throws Exception {
        if (startPage < 1 || startPage > this.getDocument().getPageCount()) {
            throw new IllegalStateException("'startPage' is out of range");
        }

        if (endPage < 1 || endPage > this.getDocument().getPageCount() || endPage < startPage) {
            throw new IllegalStateException("'endPage' is out of range");
        }

        this.checkPageListsPopulated();
        ArrayList pageNodes = new ArrayList();
        for (int page = startPage; page <= endPage; page++) {
            // Some pages can be empty.
            if (!this.reversePageLookup.containsKey(page)) {
                continue;
            }

            for (Node node : (Iterable<Node>) this.reversePageLookup.get(page)) {
                if (node.getParentNode() != null
                        && (nodeType == NodeType.ANY || node.getNodeType() == nodeType)
                        && !pageNodes.contains(node)) {
                    pageNodes.add(node);
                }
            }
        }

        return pageNodes;
    }

    /// <summary>
    /// Splits nodes which appear over two or more pages into separate nodes so that they still appear in the same way
    /// but no longer appear across a page.
    /// </summary>
    public void splitNodesAcrossPages() throws Exception {
        for (Paragraph paragraph : (Iterable<Paragraph>) this.getDocument().getChildNodes(NodeType.PARAGRAPH, true)) {
            if (this.getPage(paragraph) != this.getPageEnd(paragraph)) {
                this.splitRunsByWords(paragraph);
            }
        }

        this.clearCollector();

        // Visit any composites which are possibly split across pages and split them into separate nodes.
        this.getDocument().accept(new SectionSplitter(this));
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
    void addPageNumbersForNode(Node node, int startPage, int endPage) {
        if (startPage > 0) {
            this.nodeStartPageLookup.put(node, startPage);
        }

        if (endPage > 0) {
            this.nodeEndPageLookup.put(node, endPage);
        }
    }

    private static boolean isHeaderFooterType(Node node) {
        return node.getNodeType() == NodeType.HEADER_FOOTER || node.getAncestor(NodeType.HEADER_FOOTER) != null;
    }

    private void checkPageListsPopulated() throws Exception {
        if (this.reversePageLookup != null) {
            return;
        }

        this.reversePageLookup = new Hashtable();

        // Add each node to a list which represent the nodes found on each page.
        for (Node node : (Iterable<Node>) this.getDocument().getChildNodes(NodeType.ANY, true)) {
            // Headers/Footers follow sections. They are not split by themselves.
            if (isHeaderFooterType(node)) {
                continue;
            }

            int startPage = this.getPage(node);
            int endPage = this.getPageEnd(node);
            for (int page = startPage; page <= endPage; page++) {
                if (!this.reversePageLookup.containsKey(page)) {
                    this.reversePageLookup.put(page, new ArrayList());
                }

                ((ArrayList) this.reversePageLookup.get(page)).add(node);
            }
        }
    }

    private void splitRunsByWords(Paragraph paragraph) throws Exception {
        for (Run run : paragraph.getRuns()) {
            if (this.getPage(run) == this.getPageEnd(run)) {
                continue;
            }

            this.splitRunByWords(run);
        }
    }

    private void splitRunByWords(Run run) {
        String[] words = run.getText().split(" ");
        List<String> list = Arrays.asList(words);
        Collections.reverse(list);
        String[] reversedWords = (String[]) list.toArray();

        for (String word : reversedWords) {
            int pos = run.getText().length() - word.length() - 1;
            if (pos > 1) {
                splitRun(run, run.getText().length() - word.length() - 1);
            }
        }
    }

    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    private static Run splitRun(Run run, int position) {
        Run afterRun = (Run) run.deepClone(true);
        afterRun.setText(run.getText().substring(position));
        run.setText(run.getText().substring(0, position));
        run.getParentNode().insertAfter(afterRun, run);
        return afterRun;
    }

    private void clearCollector() throws Exception {
        this.collector.clear();
        this.getDocument().updatePageLayout();

        this.nodeStartPageLookup.clear();
        this.nodeEndPageLookup.clear();
    }
}

class PageNumberFinderFactory {
    /* Simulation of static class by using private constructor */
    private PageNumberFinderFactory() {
    }

    public static PageNumberFinder create(Document document) throws Exception {
        LayoutCollector layoutCollector = new LayoutCollector(document);
        document.updatePageLayout();
        PageNumberFinder pageNumberFinder = new PageNumberFinder(layoutCollector);
        pageNumberFinder.splitNodesAcrossPages();
        return pageNumberFinder;
    }
}

class SectionSplitter extends DocumentVisitor {
    private PageNumberFinder pageNumberFinder;

    public SectionSplitter(PageNumberFinder pageNumberFinder) {
        this.pageNumberFinder = pageNumberFinder;
    }

    public int visitParagraphStart(Paragraph paragraph) throws Exception {
        return this.continueIfCompositeAcrossPageElseSkip(paragraph);
    }

    public int visitTableStart(Table table) throws Exception {
        return this.continueIfCompositeAcrossPageElseSkip(table);
    }

    public int visitRowStart(Row row) throws Exception {
        return this.continueIfCompositeAcrossPageElseSkip(row);
    }

    public int visitCellStart(Cell cell) throws Exception {
        return this.continueIfCompositeAcrossPageElseSkip(cell);
    }

    public int visitStructuredDocumentTagStart(StructuredDocumentTag sdt) throws Exception {
        return this.continueIfCompositeAcrossPageElseSkip(sdt);
    }

    public int visitSmartTagStart(SmartTag smartTag) throws Exception {
        return this.continueIfCompositeAcrossPageElseSkip(smartTag);
    }

    public int visitSectionStart(Section section) throws Exception {
        Section previousSection = (Section) section.getPreviousSibling();

        // If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an
        // extracted document if the previous section is missing.
        if (previousSection != null) {
            HeaderFooterCollection previousHeaderFooters = previousSection.getHeadersFooters();
            if (!section.getPageSetup().getRestartPageNumbering()) {
                section.getPageSetup().setRestartPageNumbering(true);
                section.getPageSetup().setPageStartingNumber(previousSection.getPageSetup().getPageStartingNumber() + this.pageNumberFinder.pageSpan(previousSection));
            }

            for (HeaderFooter previousHeaderFooter : (Iterable<HeaderFooter>) previousHeaderFooters) {
                if (section.getHeadersFooters().getByHeaderFooterType(previousHeaderFooter.getHeaderFooterType()) == null) {
                    HeaderFooter newHeaderFooter = (HeaderFooter) previousHeaderFooters.getByHeaderFooterType(previousHeaderFooter.getHeaderFooterType()).deepClone(true);
                    section.getHeadersFooters().add(newHeaderFooter);
                }
            }
        }

        return this.continueIfCompositeAcrossPageElseSkip(section);
    }

    public int visitSmartTagEnd(SmartTag smartTag) throws Exception {
        this.splitComposite(smartTag);
        return VisitorAction.CONTINUE;
    }

    public int visitStructuredDocumentTagEnd(StructuredDocumentTag sdt) throws Exception {
        this.splitComposite(sdt);
        return VisitorAction.CONTINUE;
    }

    public int visitCellEnd(Cell cell) throws Exception {
        this.splitComposite(cell);
        return VisitorAction.CONTINUE;
    }

    public int visitRowEnd(Row row) throws Exception {
        this.splitComposite(row);
        return VisitorAction.CONTINUE;
    }

    public int visitTableEnd(Table table) throws Exception {
        this.splitComposite(table);
        return VisitorAction.CONTINUE;
    }

    public int visitParagraphEnd(Paragraph paragraph) throws Exception {
        // If paragraph contains only section break, add fake run into
        if (paragraph.isEndOfSection() && paragraph.getChildNodes().getCount() == 1 && "\f".equals(paragraph.getChildNodes().get(0).getText())) {
            Run run = new Run(paragraph.getDocument());
            paragraph.appendChild(run);
            int currentEndPageNum = this.pageNumberFinder.getPageEnd(paragraph);
            this.pageNumberFinder.addPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
        }

        for (Paragraph clonePara : (Iterable<Paragraph>) splitComposite(paragraph)) {
            // Remove list numbering from the cloned paragraph but leave the indent the same
            // as the paragraph is supposed to be part of the item before.
            if (paragraph.isListItem()) {
                double textPosition = clonePara.getListFormat().getListLevel().getTextPosition();
                clonePara.getListFormat().removeNumbers();
                clonePara.getParagraphFormat().setLeftIndent(textPosition);
            }
            // Reset spacing of split paragraphs in tables as additional spacing may cause them to look different.
            if (paragraph.isInCell()) {
                clonePara.getParagraphFormat().setSpaceBefore(0);
                paragraph.getParagraphFormat().setSpaceAfter(0);
            }
        }

        return VisitorAction.CONTINUE;
    }

    public int visitSectionEnd(Section section) throws Exception {
        for (Section cloneSection : (Iterable<Section>) this.splitComposite(section)) {
            cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
            cloneSection.getPageSetup().setRestartPageNumbering(true);
            cloneSection.getPageSetup().setPageStartingNumber(section.getPageSetup().getPageStartingNumber() +
                    (section.getDocument().indexOf(cloneSection) - section.getDocument().indexOf(section)));
            cloneSection.getPageSetup().setDifferentFirstPageHeaderFooter(false);

            // corrects page break on end of the section
            SplitPageBreakCorrector.processSection(cloneSection);
        }

        // corrects page break on end of the section
        SplitPageBreakCorrector.processSection(section);

        // Add new page numbering for the body of the section as well.
        this.pageNumberFinder.addPageNumbersForNode(section.getBody(), this.pageNumberFinder.getPage(section), this.pageNumberFinder.getPageEnd(section));
        return VisitorAction.CONTINUE;
    }

    private int continueIfCompositeAcrossPageElseSkip(CompositeNode composite) throws Exception {
        return (this.pageNumberFinder.pageSpan(composite) > 1) ? VisitorAction.CONTINUE : VisitorAction.SKIP_THIS_NODE;
    }

    private ArrayList splitComposite(CompositeNode composite) throws Exception {
        ArrayList splitNodes = new ArrayList</* unknown Type use JavaGenericArguments */>();
        for (Node splitNode : (Iterable<Node>) this.findChildSplitPositions(composite)) {
            splitNodes.add(this.splitCompositeAtNode(composite, splitNode));
        }

        return splitNodes;
    }

    private ArrayList findChildSplitPositions(CompositeNode node) throws Exception {
        // A node may span across multiple pages so a list of split positions is returned.
        // The split node is the first node on the next page.
        ArrayList splitList = new ArrayList();
        int startingPage = this.pageNumberFinder.getPage(node);
        Node[] childNodes = node.getNodeType() == NodeType.SECTION
                ? ((Section) node).getBody().getChildNodes().toArray()
                : node.getChildNodes().toArray();
        for (Node childNode : childNodes) {
            int pageNum = this.pageNumberFinder.getPage(childNode);

            if (childNode instanceof Run) {
                pageNum = this.pageNumberFinder.getPageEnd(childNode);
            }

            // If the page of the child node has changed then this is the split position. Add
            // this to the list.
            if (pageNum > startingPage) {
                splitList.add(childNode);
                startingPage = pageNum;
            }

            if (this.pageNumberFinder.pageSpan(childNode) > 1) {
                this.pageNumberFinder.addPageNumbersForNode(childNode, pageNum, pageNum);
            }
        }

        // Split composites backward so the cloned nodes are inserted in the right order.
        Collections.reverse(splitList);
        return splitList;
    }

    private CompositeNode splitCompositeAtNode(CompositeNode baseNode, Node targetNode) throws Exception {
        CompositeNode cloneNode = (CompositeNode) baseNode.deepClone(false);
        Node node = targetNode;
        int currentPageNum = this.pageNumberFinder.getPage(baseNode);

        // Move all nodes found on the next page into the copied node. Handle row nodes separately.
        if (baseNode.getNodeType() != NodeType.ROW) {
            CompositeNode composite = cloneNode;
            if (baseNode.getNodeType() == NodeType.SECTION) {
                cloneNode = (CompositeNode) baseNode.deepClone(true);
                Section section = (Section) cloneNode;
                section.getBody().removeAllChildren();
                composite = section.getBody();
            }

            while (node != null) {
                Node nextNode = node.getNextSibling();
                composite.appendChild(node);
                node = nextNode;
            }
        } else {
            // If we are dealing with a row then we need to add in dummy cells for the cloned row.
            int targetPageNum = this.pageNumberFinder.getPage(targetNode);
            Node[] childNodes = baseNode.getChildNodes().toArray();
            for (Node childNode : childNodes) {
                int pageNum = this.pageNumberFinder.getPage(childNode);
                if (pageNum == targetPageNum) {
                    cloneNode.getLastChild().remove();
                    cloneNode.appendChild(childNode);
                } else if (pageNum == currentPageNum) {
                    cloneNode.appendChild(childNode.deepClone(false));
                    if (cloneNode.getLastChild().getNodeType() != NodeType.CELL) {
                        ((CompositeNode) cloneNode.getLastChild()).appendChild(((CompositeNode) childNode).getFirstChild().deepClone(false));
                    }
                }
            }
        }

        // Insert the split node after the original.
        baseNode.getParentNode().insertAfter(cloneNode, baseNode);

        // Update the new page numbers of the base node and the clone node including its descendents.
        // This will only be a single page as the cloned composite is split to be on one page.
        int currentEndPageNum = this.pageNumberFinder.getPageEnd(baseNode);
        this.pageNumberFinder.addPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
        this.pageNumberFinder.addPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);

        for (Node childNode : (Iterable<Node>) cloneNode.getChildNodes(NodeType.ANY, true)) {
            this.pageNumberFinder.addPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
        }

        return cloneNode;
    }
}

class SplitPageBreakCorrector {
    private static final String PAGE_BREAK_STR = "\f";
    private static final char PAGE_BREAK = '\f';

    public static void processSection(Section section) {
        if (section.getChildNodes().getCount() == 0) {
            return;
        }

        Body lastBody = section.getBody();
        if (lastBody == null) {
            return;
        }

        Run run = null;
        for (Run r : (Iterable<Run>) lastBody.getChildNodes(NodeType.RUN, true)) {
            if (r.getText().endsWith(PAGE_BREAK_STR)) {
                run = r;
                break;
            }
        }

        if (run != null) {
            removePageBreak(run);
        }

        return;
    }

    public static void removePageBreakFromParagraph(Paragraph paragraph) {
        Run run = (Run) paragraph.getFirstChild();
        if (run.getText().equals(PAGE_BREAK_STR)) {
            paragraph.removeChild(run);
        }
    }

    private static void processLastParagraph(Paragraph paragraph) {
        Node lastNode = paragraph.getChildNodes().get(paragraph.getChildNodes().getCount() - 1);
        if (lastNode.getNodeType() != NodeType.RUN) {
            return;
        }

        Run run = (Run) lastNode;
        removePageBreak(run);
    }

    private static void removePageBreak(Run run) {
        Paragraph paragraph = run.getParentParagraph();
        if (run.getText().equals(PAGE_BREAK_STR)) {
            paragraph.removeChild(run);
        } else if (run.getText().endsWith(PAGE_BREAK_STR)) {
            run.setText(run.getText().replaceAll("[" + PAGE_BREAK + "]+$", ""));
        }

        if (paragraph.getChildNodes().getCount() == 0) {
            CompositeNode parent = paragraph.getParentNode();
            parent.removeChild(paragraph);
        }
    }
}

//ExEnd:DocumentPageSplitter

