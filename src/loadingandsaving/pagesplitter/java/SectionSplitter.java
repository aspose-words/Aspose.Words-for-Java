/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package loadingandsaving.pagesplitter.java;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Hashtable;
import java.util.Stack;

import com.aspose.words.*;

public class SectionSplitter extends DocumentVisitor {
    public SectionSplitter(PageNumberFinder pageNumberFinder) {
        mPageNumberFinder = pageNumberFinder;
    }

    public int VisitParagraphStart(Paragraph paragraph) throws Exception {
        if (paragraph.isListItem()) {
            List paraList = paragraph.getListFormat().getList();
            ListLevel currentLevel = paragraph.getListFormat().getListLevel();

            // Since we have encountered a list item we need to check if this will reset
            // any subsequent list levels and if so then update the numbering of the level.
            int currentListLevelNumber = paragraph.getListFormat().getListLevelNumber();
            for (int i = currentListLevelNumber + 1; i < paraList.getListLevels().getCount(); i++) {
                ListLevel paraLevel = paraList.getListLevels().get(i);

                if (paraLevel.getRestartAfterLevel() >= currentListLevelNumber) {
                    // This list level needs to be reset after the current list number.
                    mListLevelToListNumberLookup.put(paraLevel, paraLevel.getStartAt());
                }
            }

            // A list which was used on a previous page is present on a different page, the list
            // needs to be copied so list numbering is retained when extracting individual pages.
            if (ContainsListLevelAndPageChanged(paragraph)) {
                List copyList = paragraph.getDocument().getLists().addCopy(paraList);
                mListLevelToListNumberLookup.put(currentLevel, paragraph.getListLabel().getLabelValue());

                // Set the numbering of each list level to start at the numbering of the level on the previous page.
                for (int i = 0; i < paraList.getListLevels().getCount(); i++) {
                    ListLevel paraLevel = paraList.getListLevels().get(i);

                    if (mListLevelToListNumberLookup.containsKey(paraLevel))
                        copyList.getListLevels().get(i).setStartAt((Integer) mListLevelToListNumberLookup.get(paraLevel));
                }

                mListToReplacementListLookup.put(paraList, copyList);
            }

            if (mListToReplacementListLookup.containsKey(paraList)) {
                // This paragraph belongs to a list from a previous page. Apply the replacement list.
                paragraph.getListFormat().setList((List) mListToReplacementListLookup.get(paraList));
                // This is a trick to get the spacing of the list level to set correctly.
                paragraph.getListFormat().setListLevelNumber(paragraph.getListFormat().getListLevelNumber() + 0);
            }

            mListLevelToPageLookup.put(currentLevel, mPageNumberFinder.GetPage(paragraph));
            mListLevelToListNumberLookup.put(currentLevel, paragraph.getListLabel().getLabelValue());
        }

        Section prevSection = (Section) paragraph.getParentSection().getPreviousSibling();
        Paragraph prevBodyPara = (Paragraph) paragraph.getPreviousSibling();

        Paragraph prevSectionPara = prevSection != null && paragraph == paragraph.getParentSection().getBody().getFirstChild() ? prevSection.getBody().getLastParagraph() : null;
        Paragraph prevParagraph = prevBodyPara != null ? prevBodyPara : prevSectionPara;

        if (paragraph.isEndOfSection() && !paragraph.hasChildNodes())
            paragraph.remove();

        // Paragraphs across pages can merge or remove spacing depending upon the previous paragraph.
        if (prevParagraph != null) {
            if (mPageNumberFinder.GetPage(paragraph) != mPageNumberFinder.GetPageEnd(prevParagraph)) {
                if (paragraph.isListItem() && prevParagraph.isListItem() && !prevParagraph.isEndOfSection())
                    prevParagraph.getParagraphFormat().setSpaceAfter(0);
                else if (prevParagraph.getParagraphFormat().getStyleName() == paragraph.getParagraphFormat().getStyleName() && paragraph.getParagraphFormat().getNoSpaceBetweenParagraphsOfSameStyle())
                    paragraph.getParagraphFormat().setSpaceBefore(0);
                else if (paragraph.getParagraphFormat().getPageBreakBefore() || (prevParagraph.isEndOfSection() && prevSection.getPageSetup().getSectionStart() != SectionStart.NEW_COLUMN))
                    paragraph.getParagraphFormat().setSpaceBefore(Math.max(paragraph.getParagraphFormat().getSpaceBefore() - prevParagraph.getParagraphFormat().getSpaceAfter(), 0));
                else
                    paragraph.getParagraphFormat().setSpaceBefore(0);
            }
        }

        return VisitorAction.CONTINUE;
    }

    public int VisitSectionStart(Section section) throws Exception {
        mSectionCount++;
        Section previousSection = (Section) section.getPreviousSibling();

        // If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an
        // extracted document if the previous section is missing.
        if (previousSection != null) {
            if (!section.getPageSetup().getRestartPageNumbering()) {
                section.getPageSetup().setRestartPageNumbering(true);
                section.getPageSetup().setPageStartingNumber(previousSection.getPageSetup().getPageStartingNumber() + (int) mPageNumberFinder.PageSpan(previousSection));
            }

            for (HeaderFooter previousHeaderFooter : previousSection.getHeadersFooters()) {
                if (section.getHeadersFooters().getByHeaderFooterType(previousHeaderFooter.getHeaderFooterType()) == null) {
                    HeaderFooter newHeaderFooter = (HeaderFooter) previousSection.getHeadersFooters().getByHeaderFooterType(previousHeaderFooter.getHeaderFooterType()).deepClone(true);
                    section.getHeadersFooters().add(newHeaderFooter);
                }
            }
        }

        // Manually set the result of these fields before sections are split.
        for (HeaderFooter headerFooter : section.getHeadersFooters()) {
            for (Field field : headerFooter.getRange().getFields()) {
                if (field.getType() == FieldType.FIELD_SECTION || field.getType() == FieldType.FIELD_SECTION_PAGES) {
                    field.setResult((field.getType() == FieldType.FIELD_SECTION) ? Integer.toString(mSectionCount) :
                            Integer.toString(mPageNumberFinder.PageSpan(section)));

                    field.isLocked(true);
                }
            }
        }

        // All fields in the body should stay the same, this also improves field update time.
        for (Field field : section.getBody().getRange().getFields())
            field.isLocked(true);

        return VisitorAction.CONTINUE;
    }

    public int VisitDocumentEnd(Document doc) throws Exception {
        // All sections have separate headers and footers now, update the fields in all headers and footers
        // to the correct values. This allows each page to maintain the correct field results even when
        // PAGE or IF fields are used.
        doc.updateFields();

        for (HeaderFooter headerFooter : (Iterable<HeaderFooter>) doc.getChildNodes(NodeType.HEADER_FOOTER, true)) {
            for (Field field : headerFooter.getRange().getFields())
                field.isLocked(true);
        }

        return VisitorAction.CONTINUE;
    }

    public int VisitSmartTagEnd(SmartTag smartTag) throws Exception {
        if (IsCompositeAcrossPage(smartTag))
            SplitComposite(smartTag);

        return VisitorAction.CONTINUE;
    }

    public int VisitCustomXmlMarkupEnd(CustomXmlMarkup customXmlMarkup) throws Exception {
        if (IsCompositeAcrossPage(customXmlMarkup))
            SplitComposite(customXmlMarkup);

        return VisitorAction.CONTINUE;
    }

    public int VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt) throws Exception {
        if (IsCompositeAcrossPage(sdt))
            SplitComposite(sdt);

        return VisitorAction.CONTINUE;
    }

    public int VisitCellEnd(Cell cell) throws Exception {
        if (IsCompositeAcrossPage(cell))
            SplitComposite(cell);

        return VisitorAction.CONTINUE;
    }

    public int VisitRowEnd(Row row) throws Exception {
        if (IsCompositeAcrossPage(row))
            SplitComposite(row);

        return VisitorAction.CONTINUE;
    }

    public int VisitTableEnd(Table table) throws Exception {
        if (IsCompositeAcrossPage(table)) {
            // Copy any header rows to other pages.
            Stack stack = new Stack();
            stack.add(table.getRows().toArray());

            for (Table cloneTable : (Iterable<Table>) SplitComposite(table)) {
                for (Row row : (Iterable<Row>) stack) {
                    if (row.getRowFormat().getHeadingFormat())
                        cloneTable.prependChild(row.deepClone(true));
                }
            }
        }

        return VisitorAction.CONTINUE;
    }

    public int VisitParagraphEnd(Paragraph paragraph) throws Exception {
        if (IsCompositeAcrossPage(paragraph)) {
            for (Paragraph clonePara : (Iterable<Paragraph>) SplitComposite(paragraph)) {
                // Remove list numbering from the cloned paragraph but leave the indent the same
                // as the paragraph is supposed to be part of the item before.
                if (paragraph.isListItem()) {
                    double textPosition = clonePara.getListFormat().getListLevel().getTextPosition();
                    clonePara.getListFormat().removeNumbers();
                    clonePara.getParagraphFormat().setLeftIndent(textPosition);
                }

                // Reset spacing of split paragraphs as additional spacing is removed.
                clonePara.getParagraphFormat().setSpaceBefore(0);
                paragraph.getParagraphFormat().setSpaceAfter(0);
            }
        }

        return VisitorAction.CONTINUE;
    }

    public int VisitSectionEnd(Section section) throws Exception {
        if (IsCompositeAcrossPage(section)) {
            // If a TOC field spans across more than one page then the hyperlink formatting may show through.
            // Remove direct formatting to avoid this.
            for (FieldStart start : (Iterable<FieldStart>) section.getChildNodes(NodeType.FIELD_START, true)) {
                if (start.getFieldType() == FieldType.FIELD_TOC) {
                    Field field = start.getField();
                    Node node = field.getSeparator();

                    while ((node = node.nextPreOrder(section)) != field.getEnd())
                        if (node.getNodeType() == NodeType.RUN)
                            ((Run) node).getFont().clearFormatting();
                }
            }

            for (Section cloneSection : (Iterable<Section>) SplitComposite(section)) {
                cloneSection.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
                cloneSection.getPageSetup().setRestartPageNumbering(true);
                cloneSection.getPageSetup().setPageStartingNumber(section.getPageSetup().getPageStartingNumber() + (section.getDocument().indexOf(cloneSection) - section.getDocument().indexOf(section)));
                cloneSection.getPageSetup().setDifferentFirstPageHeaderFooter(false);

                RemovePageBreaksFromParagraph(cloneSection.getBody().getLastParagraph());
            }

            RemovePageBreaksFromParagraph(section.getBody().getLastParagraph());

            // Add new page numbering for the body of the section as well.
            mPageNumberFinder.AddPageNumbersForNode(section.getBody(), mPageNumberFinder.GetPage(section), mPageNumberFinder.GetPageEnd(section));
        }

        return VisitorAction.CONTINUE;
    }

    private boolean IsCompositeAcrossPage(CompositeNode composite) throws Exception {
        return (mPageNumberFinder.PageSpan(composite) > 1);
    }

    private boolean ContainsListLevelAndPageChanged(Paragraph para) throws Exception {
        return mListLevelToPageLookup.containsKey(para.getListFormat().getListLevel()) && (Integer) mListLevelToPageLookup.get(para.getListFormat().getListLevel()) != mPageNumberFinder.GetPage(para);
    }

    private void RemovePageBreaksFromParagraph(Paragraph para) {
        if (para != null) {
            for (Run run : para.getRuns())
                run.setText(run.getText().replace(ControlChar.PAGE_BREAK, ""));
        }
    }

    private ArrayList SplitComposite(CompositeNode composite) throws Exception {
        ArrayList splitNodes = new ArrayList();
        for (Node splitNode : (Iterable<Node>) FindChildSplitPositions(composite))
            splitNodes.add(SplitCompositeAtNode(composite, splitNode));

        return splitNodes;
    }

    private ArrayList FindChildSplitPositions(CompositeNode node) throws Exception {
        // A node may span across multiple pages so a list of split positions is returned.
        // The split node is the first node on the next page.
        ArrayList splitList = new ArrayList();

        int startingPage = mPageNumberFinder.GetPage(node);

        Node[] childNodes = node.getNodeType() == NodeType.SECTION ?
                ((Section) node).getBody().getChildNodes().toArray() : node.getChildNodes().toArray();

        for (Node childNode : childNodes) {
            int pageNum = mPageNumberFinder.GetPage(childNode);

            // If the page of the child node has changed then this is the split position. Add
            // this to the list.
            if (pageNum > startingPage) {
                splitList.add(childNode);
                startingPage = pageNum;
            }

            if (mPageNumberFinder.PageSpan(childNode) > 1)
                mPageNumberFinder.AddPageNumbersForNode(childNode, pageNum, pageNum);
        }

        // Split composites backward so the cloned nodes are inserted in the right order.
        Collections.reverse(splitList);

        return splitList;
    }

    private CompositeNode SplitCompositeAtNode(CompositeNode baseNode, Node targetNode) throws Exception {
        CompositeNode cloneNode = (CompositeNode) baseNode.deepClone(false);

        Node node = targetNode;
        int currentPageNum = mPageNumberFinder.GetPage(baseNode);

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
            int targetPageNum = mPageNumberFinder.GetPage(targetNode);
            Node[] childNodes = baseNode.getChildNodes().toArray();

            for (Node childNode : childNodes) {
                int pageNum = mPageNumberFinder.GetPage(childNode);

                if (pageNum == targetPageNum) {
                    cloneNode.getLastChild().remove();
                    cloneNode.appendChild(childNode);
                } else if (pageNum == currentPageNum) {
                    cloneNode.appendChild(childNode.deepClone(false));
                    if (cloneNode.getLastChild().getNodeType() != NodeType.CELL)
                        ((CompositeNode) cloneNode.getLastChild()).appendChild(((CompositeNode) childNode).getFirstChild().deepClone(false));
                }
            }
        }

        // Insert the split node after the original.
        baseNode.getParentNode().insertAfter(cloneNode, baseNode);

        // Update the new page numbers of the base node and the clone node including its descendents.
        // This will only be a single page as the cloned composite is split to be on one page.
        int currentEndPageNum = mPageNumberFinder.GetPageEnd(baseNode);
        mPageNumberFinder.AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
        mPageNumberFinder.AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);

        for (Node childNode : (Iterable<Node>) cloneNode.getChildNodes(NodeType.ANY, true))
            mPageNumberFinder.AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);

        return cloneNode;
    }

    private Hashtable mListLevelToListNumberLookup = new Hashtable();
    private Hashtable mListToReplacementListLookup = new Hashtable();
    private Hashtable mListLevelToPageLookup = new Hashtable();
    private PageNumberFinder mPageNumberFinder;
    private int mSectionCount;
}
