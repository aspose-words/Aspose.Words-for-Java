/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package loadingandsaving.pagesplitter.java;

import com.aspose.words.*;

public class DocumentPageSplitter {
    /**
     * Initializes new instance of this class. This method splits the document into sections so that each page
     * begins and ends at a section boundary. It is recommended not to modify the document afterwards.
     */
    public DocumentPageSplitter(LayoutCollector collector) throws Exception {
        mPageNumberFinder = new PageNumberFinder(collector);
        mPageNumberFinder.SplitNodesAcrossPages();
    }

    /**
     * Gets the document of a page.
     */
    public Document GetDocumentOfPage(int pageIndex) throws Exception {
        return GetDocumentOfPageRange(pageIndex, pageIndex);
    }

    /**
     * Gets the document of a page range.
     */
    public Document GetDocumentOfPageRange(int startIndex, int endIndex) throws Exception {
        Document result = (Document) getDocument().deepClone(false);

        for (Section section : (Iterable<Section>) mPageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex, NodeType.SECTION))
            result.appendChild(result.importNode(section, true));

        return result;
    }

    /**
     * Gets the document this instance works with.
     */
    private Document getDocument() {
        return mPageNumberFinder.getDocument();
    }

    private PageNumberFinder mPageNumberFinder;
}
