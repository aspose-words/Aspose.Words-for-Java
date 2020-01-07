package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.CompatibilityOptions;
import com.aspose.words.Document;
import com.aspose.words.MsWordVersion;
import org.testng.Assert;
import org.testng.annotations.Test;

@Test
public class ExCompatibilityOptions extends ApiExampleBase {
    @Test
    public void compatibilityOptionsTable() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2002);

        Assert.assertEquals(false, compatibilityOptions.getAdjustLineHeightInTable());
        Assert.assertEquals(false, compatibilityOptions.getAlignTablesRowByRow());
        Assert.assertEquals(true, compatibilityOptions.getAllowSpaceOfSameStyleInTable());
        Assert.assertEquals(true, compatibilityOptions.getDoNotAutofitConstrainedTables());
        Assert.assertEquals(true, compatibilityOptions.getDoNotBreakConstrainedForcedTable());
        Assert.assertEquals(false, compatibilityOptions.getDoNotBreakWrappedTables());
        Assert.assertEquals(false, compatibilityOptions.getDoNotSnapToGridInCell());
        Assert.assertEquals(false, compatibilityOptions.getDoNotUseHTMLParagraphAutoSpacing());
        Assert.assertEquals(true, compatibilityOptions.getDoNotVertAlignCellWithSp());
        Assert.assertEquals(false, compatibilityOptions.getForgetLastTabAlignment());
        Assert.assertEquals(true, compatibilityOptions.getGrowAutofit());
        Assert.assertEquals(false, compatibilityOptions.getLayoutRawTableWidth());
        Assert.assertEquals(false, compatibilityOptions.getLayoutTableRowsApart());
        Assert.assertEquals(false, compatibilityOptions.getNoColumnBalance());
        Assert.assertEquals(false, compatibilityOptions.getOverrideTableStyleFontSizeAndJustification());
        Assert.assertEquals(false, compatibilityOptions.getUseSingleBorderforContiguousCells());
        Assert.assertEquals(true, compatibilityOptions.getUseWord2002TableStyleRules());
        Assert.assertEquals(false, compatibilityOptions.getUseWord2010TableStyleRules());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsTable.docx");
    }

    @Test
    public void compatibilityOptionsBreaks() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getApplyBreakingRules());
        Assert.assertEquals(true, compatibilityOptions.getDoNotUseEastAsianBreakRules());
        Assert.assertEquals(false, compatibilityOptions.getShowBreaksInFrames());
        Assert.assertEquals(true, compatibilityOptions.getSplitPgBreakAndParaMark());
        Assert.assertEquals(true, compatibilityOptions.getUseAltKinsokuLineBreakRules());
        Assert.assertEquals(false, compatibilityOptions.getUseWord97LineBreakRules());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsBreaks.docx");
    }

    @Test
    public void compatibilityOptionsSpacing() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getAutoSpaceLikeWord95());
        Assert.assertEquals(true, compatibilityOptions.getDisplayHangulFixedWidth());
        Assert.assertEquals(false, compatibilityOptions.getNoExtraLineSpacing());
        Assert.assertEquals(false, compatibilityOptions.getNoLeading());
        Assert.assertEquals(false, compatibilityOptions.getNoSpaceRaiseLower());
        Assert.assertEquals(false, compatibilityOptions.getSpaceForUL());
        Assert.assertEquals(false, compatibilityOptions.getSpacingInWholePoints());
        Assert.assertEquals(false, compatibilityOptions.getSuppressBottomSpacing());
        Assert.assertEquals(false, compatibilityOptions.getSuppressSpBfAfterPgBrk());
        Assert.assertEquals(false, compatibilityOptions.getSuppressSpacingAtTopOfPage());
        Assert.assertEquals(false, compatibilityOptions.getSuppressTopSpacing());
        Assert.assertEquals(false, compatibilityOptions.getUlTrailSpace());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsSpacing.docx");
    }

    @Test
    public void compatibilityOptionsWordPerfect() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getSuppressTopSpacingWP());
        Assert.assertEquals(false, compatibilityOptions.getTruncateFontHeightsLikeWP6());
        Assert.assertEquals(false, compatibilityOptions.getWPJustification());
        Assert.assertEquals(false, compatibilityOptions.getWPSpaceWidth());
        Assert.assertEquals(false, compatibilityOptions.getWrapTrailSpaces());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsWordPerfect.docx");
    }

    @Test
    public void compatibilityOptionsAlignment() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(true, compatibilityOptions.getCachedColBalance());
        Assert.assertEquals(true, compatibilityOptions.getDoNotVertAlignInTxbx());
        Assert.assertEquals(true, compatibilityOptions.getDoNotWrapTextWithPunct());
        Assert.assertEquals(false, compatibilityOptions.getNoTabHangInd());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsAlignment.docx");
    }

    @Test
    public void compatibilityOptionsLegacy() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getFootnoteLayoutLikeWW8());
        Assert.assertEquals(false, compatibilityOptions.getLineWrapLikeWord6());
        Assert.assertEquals(false, compatibilityOptions.getMWSmallCaps());
        Assert.assertEquals(false, compatibilityOptions.getShapeLayoutLikeWW8());
        Assert.assertEquals(false, compatibilityOptions.getUICompat97To2003());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsLegacy.docx");
    }

    @Test
    public void compatibilityOptionsList() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(true, compatibilityOptions.getUnderlineTabInNumList());
        Assert.assertEquals(true, compatibilityOptions.getUseNormalStyleForList());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsList.docx");
    }

    @Test
    public void compatibilityOptionsMisc() throws Exception {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getBalanceSingleByteDoubleByteWidth());
        Assert.assertEquals(false, compatibilityOptions.getConvMailMergeEsc());
        Assert.assertEquals(false, compatibilityOptions.getDoNotExpandShiftReturn());
        Assert.assertEquals(false, compatibilityOptions.getDoNotLeaveBackslashAlone());
        Assert.assertEquals(false, compatibilityOptions.getDoNotSuppressParagraphBorders());
        Assert.assertEquals(true, compatibilityOptions.getDoNotUseIndentAsNumberingTabStop());
        Assert.assertEquals(false, compatibilityOptions.getPrintBodyTextBeforeHeader());
        Assert.assertEquals(false, compatibilityOptions.getPrintColBlack());
        Assert.assertEquals(true, compatibilityOptions.getSelectFldWithFirstOrLastChar());
        Assert.assertEquals(false, compatibilityOptions.getSubFontBySize());
        Assert.assertEquals(false, compatibilityOptions.getSwapBordersFacingPgs());
        Assert.assertEquals(false, compatibilityOptions.getTransparentMetafiles());
        Assert.assertEquals(true, compatibilityOptions.getUseAnsiKerningPairs());
        Assert.assertEquals(false, compatibilityOptions.getUseFELayout());
        Assert.assertEquals(false, compatibilityOptions.getUsePrinterMetrics());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsMisc.docx");
    }
}
