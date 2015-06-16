<?php
/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use java\util\ArrayList as ArrayList;
use java\util\Collections as Collections;
use java\util\Hashtable as Hashtable;
use java\util\Stack as Stack;

use com\aspose\words\DocumentVisitor as DocumentVisitor;

class SectionSplitter {

    private $mListLevelToListNumberLookup;
    private $mListToReplacementListLookup;
    private $mListLevelToPageLookup;
    private $mPageNumberFinder;
    private $mSectionCount;


    /**
     * Initializes new instance of this class.
     */

    public function __construct(PageNumberFinder $pageNumberFinder) {

        $this->mListLevelToListNumberLookup = new Hashtable();
        $this->mListToReplacementListLookup = new Hashtable();
        $this->mListLevelToPageLookup = new Hashtable();

        $this->mPageNumberFinder = pageNumberFinder;

    }

    public function VisitParagraphStart(Paragraph $paragraph) {

        if ($paragraph->isListItem()) {
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

}