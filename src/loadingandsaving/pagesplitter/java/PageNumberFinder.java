/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package loadingandsaving.pagesplitter.java;

import com.aspose.words.Document;
import com.aspose.words.LayoutCollector;
import com.aspose.words.Node;
import com.aspose.words.NodeType;

import java.util.ArrayList;
import java.util.Hashtable;

public class PageNumberFinder {

    /**
      * Initializes new instance of this class.
     */
    public PageNumberFinder(LayoutCollector collector) {
        mCollector = collector;
    }

    /**
      * Retrieves 1-based index of a page that the node begins on.
     */
    public int GetPage(Node node) throws Exception {
        if (mNodeStartPageLookup.containsKey(node))
            return (Integer) mNodeStartPageLookup.get(node);

        return mCollector.getStartPageIndex(node);
    }

    /**
      * Retrieves 1-based index of a page that the node ends on.
     */
    public int GetPageEnd(Node node) throws Exception {
        if (mNodeEndPageLookup.containsKey(node))
            return (Integer) mNodeEndPageLookup.get(node);

        return mCollector.getEndPageIndex(node);
    }

    /**
      * Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
     */
    public int PageSpan(Node node) throws Exception {
        return GetPageEnd(node) - GetPage(node) + 1;
    }

    /**
      * Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
     */
    public ArrayList RetrieveAllNodesOnPages(int startPage, int endPage, int nodeType) throws Exception {
        if (startPage < 1 || startPage > getDocument().getPageCount())
            throw new Exception("startPage");

        if (endPage < 1 || endPage > getDocument().getPageCount() || endPage < startPage)
            throw new Exception("endPage");

        CheckPageListsPopulated();

        ArrayList pageNodes = new ArrayList();

        for (int page = startPage; page <= endPage; page++) {
            // Some pages can be empty.
            if (!mReversePageLookup.containsKey(page))
                continue;

            for (Node node : (Iterable<Node>) (ArrayList) mReversePageLookup.get(page)) {
                if (node.getParentNode() != null && ((nodeType == NodeType.ANY) || (nodeType == node.getNodeType())) && !pageNodes.contains(node))
                    pageNodes.add(node);
            }
        }

        return pageNodes;
    }

    /**
     * Splits nodes which appear over two or more pages into separate nodes so that they still appear in the same way
     * but no longer appear across a page.
     */
    public void SplitNodesAcrossPages() throws Exception {
        // Visit any composites which are possibly split across pages and split them into separate nodes.
        getDocument().accept(new SectionSplitter(this));
    }

    /**
     * Gets the document this instance works with.
     */
    public Document getDocument() {
        return mCollector.getDocument();
    }

    /**
     * This is called by <see cref="SectionSplitter"/> to update page numbers of split nodes.
     */
    void AddPageNumbersForNode(Node node, int startPage, int endPage) {
        if (startPage > 0)
            mNodeStartPageLookup.put(node, startPage);

        if (endPage > 0)
            mNodeEndPageLookup.put(node, endPage);
    }

    private void CheckPageListsPopulated() throws Exception {
        if (mReversePageLookup != null)
            return;

        mReversePageLookup = new Hashtable();

        // Add each node to a list which represent the nodes found on each page.
        for (Node node : (Iterable<Node>) getDocument().getChildNodes(NodeType.ANY, true)) {
            // Headers/Footers follow sections. They are not split by themselves.
            if (IsHeaderFooterType(node))
                continue;

            int startPage = GetPage(node);
            int endPage = GetPageEnd(node);

            for (int page = startPage; page <= endPage; page++) {
                if (!mReversePageLookup.containsKey(page))
                    mReversePageLookup.put(page, new ArrayList());

                ((ArrayList) mReversePageLookup.get(page)).add(node);
            }
        }
    }

    private static boolean IsHeaderFooterType(Node node) {
        return node.getNodeType() == NodeType.HEADER_FOOTER || node.getAncestor(NodeType.HEADER_FOOTER) != null;
    }

    // Maps node to a start/end page numbers. This is used to override baseline page numbers provided by collector when document is split.
    private Hashtable mNodeStartPageLookup = new Hashtable();
    private Hashtable mNodeEndPageLookup = new Hashtable();
    // Maps page number to a list of nodes found on that page.
    private Hashtable mReversePageLookup;
    private LayoutCollector mCollector;
}
