package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.NodeType;
import com.aspose.words.Section;

public class DocumentPageSplitter {
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

