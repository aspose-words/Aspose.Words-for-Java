// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.CompatibilityOptions;
import com.aspose.ms.System.msConsole;
import com.aspose.words.MsWordVersion;
import java.util.ArrayList;
import org.testng.Assert;


@Test
public class ExCompatibilityOptions extends ApiExampleBase
{
    //ExStart
    //ExFor:Compatibility
    //ExFor:CompatibilityOptions
    //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
    //ExFor:Document.CompatibilityOptions
    //ExFor:MsWordVersion
    //ExFor:CompatibilityOptions.AdjustLineHeightInTable
    //ExFor:CompatibilityOptions.AlignTablesRowByRow
    //ExFor:CompatibilityOptions.AllowSpaceOfSameStyleInTable
    //ExFor:CompatibilityOptions.ApplyBreakingRules
    //ExFor:CompatibilityOptions.AutofitToFirstFixedWidthCell
    //ExFor:CompatibilityOptions.AutoSpaceLikeWord95
    //ExFor:CompatibilityOptions.BalanceSingleByteDoubleByteWidth
    //ExFor:CompatibilityOptions.CachedColBalance
    //ExFor:CompatibilityOptions.ConvMailMergeEsc
    //ExFor:CompatibilityOptions.DisableOpenTypeFontFormattingFeatures
    //ExFor:CompatibilityOptions.DisplayHangulFixedWidth
    //ExFor:CompatibilityOptions.DoNotAutofitConstrainedTables
    //ExFor:CompatibilityOptions.DoNotBreakConstrainedForcedTable
    //ExFor:CompatibilityOptions.DoNotBreakWrappedTables
    //ExFor:CompatibilityOptions.DoNotExpandShiftReturn
    //ExFor:CompatibilityOptions.DoNotLeaveBackslashAlone
    //ExFor:CompatibilityOptions.DoNotSnapToGridInCell
    //ExFor:CompatibilityOptions.DoNotSuppressIndentation
    //ExFor:CompatibilityOptions.DoNotSuppressParagraphBorders
    //ExFor:CompatibilityOptions.DoNotUseEastAsianBreakRules
    //ExFor:CompatibilityOptions.DoNotUseHTMLParagraphAutoSpacing
    //ExFor:CompatibilityOptions.DoNotUseIndentAsNumberingTabStop
    //ExFor:CompatibilityOptions.DoNotVertAlignCellWithSp
    //ExFor:CompatibilityOptions.DoNotVertAlignInTxbx
    //ExFor:CompatibilityOptions.DoNotWrapTextWithPunct
    //ExFor:CompatibilityOptions.FootnoteLayoutLikeWW8
    //ExFor:CompatibilityOptions.ForgetLastTabAlignment
    //ExFor:CompatibilityOptions.GrowAutofit
    //ExFor:CompatibilityOptions.LayoutRawTableWidth
    //ExFor:CompatibilityOptions.LayoutTableRowsApart
    //ExFor:CompatibilityOptions.LineWrapLikeWord6
    //ExFor:CompatibilityOptions.MWSmallCaps
    //ExFor:CompatibilityOptions.NoColumnBalance
    //ExFor:CompatibilityOptions.NoExtraLineSpacing
    //ExFor:CompatibilityOptions.NoLeading
    //ExFor:CompatibilityOptions.NoSpaceRaiseLower
    //ExFor:CompatibilityOptions.NoTabHangInd
    //ExFor:CompatibilityOptions.OverrideTableStyleFontSizeAndJustification
    //ExFor:CompatibilityOptions.PrintBodyTextBeforeHeader
    //ExFor:CompatibilityOptions.PrintColBlack
    //ExFor:CompatibilityOptions.SelectFldWithFirstOrLastChar
    //ExFor:CompatibilityOptions.ShapeLayoutLikeWW8
    //ExFor:CompatibilityOptions.ShowBreaksInFrames
    //ExFor:CompatibilityOptions.SpaceForUL
    //ExFor:CompatibilityOptions.SpacingInWholePoints
    //ExFor:CompatibilityOptions.SplitPgBreakAndParaMark
    //ExFor:CompatibilityOptions.SubFontBySize
    //ExFor:CompatibilityOptions.SuppressBottomSpacing
    //ExFor:CompatibilityOptions.SuppressSpacingAtTopOfPage
    //ExFor:CompatibilityOptions.SuppressSpBfAfterPgBrk
    //ExFor:CompatibilityOptions.SuppressTopSpacing
    //ExFor:CompatibilityOptions.SuppressTopSpacingWP
    //ExFor:CompatibilityOptions.SwapBordersFacingPgs
    //ExFor:CompatibilityOptions.SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning
    //ExFor:CompatibilityOptions.TransparentMetafiles
    //ExFor:CompatibilityOptions.TruncateFontHeightsLikeWP6
    //ExFor:CompatibilityOptions.UICompat97To2003
    //ExFor:CompatibilityOptions.UlTrailSpace
    //ExFor:CompatibilityOptions.UnderlineTabInNumList
    //ExFor:CompatibilityOptions.UseAltKinsokuLineBreakRules
    //ExFor:CompatibilityOptions.UseAnsiKerningPairs
    //ExFor:CompatibilityOptions.UseFELayout
    //ExFor:CompatibilityOptions.UseNormalStyleForList
    //ExFor:CompatibilityOptions.UsePrinterMetrics
    //ExFor:CompatibilityOptions.UseSingleBorderforContiguousCells
    //ExFor:CompatibilityOptions.UseWord2002TableStyleRules
    //ExFor:CompatibilityOptions.UseWord2010TableStyleRules
    //ExFor:CompatibilityOptions.UseWord97LineBreakRules
    //ExFor:CompatibilityOptions.WPJustification
    //ExFor:CompatibilityOptions.WPSpaceWidth
    //ExFor:CompatibilityOptions.WrapTrailSpaces
    //ExSummary:Shows how to optimize the document for different versions of Microsoft Word.
    @Test //ExSkip
    public void optimizeFor() throws Exception
    {
        Document doc = new Document();

        // This object contains an extensive list of flags unique to each document
        // that allow us to facilitate backward compatibility with older versions of Microsoft Word.
        CompatibilityOptions options = doc.getCompatibilityOptions();

        // Print the default settings for a blank document.
        System.out.println("\nDefault optimization settings:");
        printCompatibilityOptions(options);

        // We can access these settings in Microsoft Word via "File" -> "Options" -> "Advanced" -> "Compatibility options for...".
        doc.save(getArtifactsDir() + "CompatibilityOptions.OptimizeFor.DefaultSettings.docx");

        // We can use the OptimizeFor method to ensure optimal compatibility with a specific Microsoft Word version.
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);
        System.out.println("\nOptimized for Word 2010:");
        printCompatibilityOptions(options);

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2000);
        System.out.println("\nOptimized for Word 2000:");
        printCompatibilityOptions(options);
    }

    /// <summary>
    /// Groups all flags in a document's compatibility options object by state, then prints each group.
    /// </summary>
    private static void printCompatibilityOptions(CompatibilityOptions options)
    {
        ArrayList<String> enabledOptions = new ArrayList<String>();
        ArrayList<String> disabledOptions = new ArrayList<String>();
        addOptionName(options.getAdjustLineHeightInTable(), "AdjustLineHeightInTable", enabledOptions, disabledOptions);
        addOptionName(options.getAlignTablesRowByRow(), "AlignTablesRowByRow", enabledOptions, disabledOptions);
        addOptionName(options.getAllowSpaceOfSameStyleInTable(), "AllowSpaceOfSameStyleInTable", enabledOptions, disabledOptions);
        addOptionName(options.getApplyBreakingRules(), "ApplyBreakingRules", enabledOptions, disabledOptions);
        addOptionName(options.getAutoSpaceLikeWord95(), "AutoSpaceLikeWord95", enabledOptions, disabledOptions);
        addOptionName(options.getAutofitToFirstFixedWidthCell(), "AutofitToFirstFixedWidthCell", enabledOptions, disabledOptions);
        addOptionName(options.getBalanceSingleByteDoubleByteWidth(), "BalanceSingleByteDoubleByteWidth", enabledOptions, disabledOptions);
        addOptionName(options.getCachedColBalance(), "CachedColBalance", enabledOptions, disabledOptions);
        addOptionName(options.getConvMailMergeEsc(), "ConvMailMergeEsc", enabledOptions, disabledOptions);
        addOptionName(options.getDisableOpenTypeFontFormattingFeatures(), "DisableOpenTypeFontFormattingFeatures", enabledOptions, disabledOptions);
        addOptionName(options.getDisplayHangulFixedWidth(), "DisplayHangulFixedWidth", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotAutofitConstrainedTables(), "DoNotAutofitConstrainedTables", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotBreakConstrainedForcedTable(), "DoNotBreakConstrainedForcedTable", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotBreakWrappedTables(), "DoNotBreakWrappedTables", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotExpandShiftReturn(), "DoNotExpandShiftReturn", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotLeaveBackslashAlone(), "DoNotLeaveBackslashAlone", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotSnapToGridInCell(), "DoNotSnapToGridInCell", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotSuppressIndentation(), "DoNotSnapToGridInCell", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotSuppressParagraphBorders(), "DoNotSuppressParagraphBorders", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotUseEastAsianBreakRules(), "DoNotUseEastAsianBreakRules", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotUseHTMLParagraphAutoSpacing(), "DoNotUseHTMLParagraphAutoSpacing", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotUseIndentAsNumberingTabStop(), "DoNotUseIndentAsNumberingTabStop", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotVertAlignCellWithSp(), "DoNotVertAlignCellWithSp", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotVertAlignInTxbx(), "DoNotVertAlignInTxbx", enabledOptions, disabledOptions);
        addOptionName(options.getDoNotWrapTextWithPunct(), "DoNotWrapTextWithPunct", enabledOptions, disabledOptions);
        addOptionName(options.getFootnoteLayoutLikeWW8(), "FootnoteLayoutLikeWW8", enabledOptions, disabledOptions);
        addOptionName(options.getForgetLastTabAlignment(), "ForgetLastTabAlignment", enabledOptions, disabledOptions);
        addOptionName(options.getGrowAutofit(), "GrowAutofit", enabledOptions, disabledOptions);
        addOptionName(options.getLayoutRawTableWidth(), "LayoutRawTableWidth", enabledOptions, disabledOptions);
        addOptionName(options.getLayoutTableRowsApart(), "LayoutTableRowsApart", enabledOptions, disabledOptions);
        addOptionName(options.getLineWrapLikeWord6(), "LineWrapLikeWord6", enabledOptions, disabledOptions);
        addOptionName(options.getMWSmallCaps(), "MWSmallCaps", enabledOptions, disabledOptions);
        addOptionName(options.getNoColumnBalance(), "NoColumnBalance", enabledOptions, disabledOptions);
        addOptionName(options.getNoExtraLineSpacing(), "NoExtraLineSpacing", enabledOptions, disabledOptions);
        addOptionName(options.getNoLeading(), "NoLeading", enabledOptions, disabledOptions);
        addOptionName(options.getNoSpaceRaiseLower(), "NoSpaceRaiseLower", enabledOptions, disabledOptions);
        addOptionName(options.getNoTabHangInd(), "NoTabHangInd", enabledOptions, disabledOptions);
        addOptionName(options.getOverrideTableStyleFontSizeAndJustification(), "OverrideTableStyleFontSizeAndJustification", enabledOptions, disabledOptions);
        addOptionName(options.getPrintBodyTextBeforeHeader(), "PrintBodyTextBeforeHeader", enabledOptions, disabledOptions);
        addOptionName(options.getPrintColBlack(), "PrintColBlack", enabledOptions, disabledOptions);
        addOptionName(options.getSelectFldWithFirstOrLastChar(), "SelectFldWithFirstOrLastChar", enabledOptions, disabledOptions);
        addOptionName(options.getShapeLayoutLikeWW8(), "ShapeLayoutLikeWW8", enabledOptions, disabledOptions);
        addOptionName(options.getShowBreaksInFrames(), "ShowBreaksInFrames", enabledOptions, disabledOptions);
        addOptionName(options.getSpaceForUL(), "SpaceForUL", enabledOptions, disabledOptions);
        addOptionName(options.getSpacingInWholePoints(), "SpacingInWholePoints", enabledOptions, disabledOptions);
        addOptionName(options.getSplitPgBreakAndParaMark(), "SplitPgBreakAndParaMark", enabledOptions, disabledOptions);
        addOptionName(options.getSubFontBySize(), "SubFontBySize", enabledOptions, disabledOptions);
        addOptionName(options.getSuppressBottomSpacing(), "SuppressBottomSpacing", enabledOptions, disabledOptions);
        addOptionName(options.getSuppressSpBfAfterPgBrk(), "SuppressSpBfAfterPgBrk", enabledOptions, disabledOptions);
        addOptionName(options.getSuppressSpacingAtTopOfPage(), "SuppressSpacingAtTopOfPage", enabledOptions, disabledOptions);
        addOptionName(options.getSuppressTopSpacing(), "SuppressTopSpacing", enabledOptions, disabledOptions);
        addOptionName(options.getSuppressTopSpacingWP(), "SuppressTopSpacingWP", enabledOptions, disabledOptions);
        addOptionName(options.getSwapBordersFacingPgs(), "SwapBordersFacingPgs", enabledOptions, disabledOptions);
        addOptionName(options.getSwapInsideAndOutsideForMirrorIndentsAndRelativePositioning(), "SwapInsideAndOutsideForMirrorIndentsAndRelativePositioning", enabledOptions, disabledOptions);
        addOptionName(options.getTransparentMetafiles(), "TransparentMetafiles", enabledOptions, disabledOptions);
        addOptionName(options.getTruncateFontHeightsLikeWP6(), "TruncateFontHeightsLikeWP6", enabledOptions, disabledOptions);
        addOptionName(options.getUICompat97To2003(), "UICompat97To2003", enabledOptions, disabledOptions);
        addOptionName(options.getUlTrailSpace(), "UlTrailSpace", enabledOptions, disabledOptions);
        addOptionName(options.getUnderlineTabInNumList(), "UnderlineTabInNumList", enabledOptions, disabledOptions);
        addOptionName(options.getUseAltKinsokuLineBreakRules(), "UseAltKinsokuLineBreakRules", enabledOptions, disabledOptions);
        addOptionName(options.getUseAnsiKerningPairs(), "UseAnsiKerningPairs", enabledOptions, disabledOptions);
        addOptionName(options.getUseFELayout(), "UseFELayout", enabledOptions, disabledOptions);
        addOptionName(options.getUseNormalStyleForList(), "UseNormalStyleForList", enabledOptions, disabledOptions);
        addOptionName(options.getUsePrinterMetrics(), "UsePrinterMetrics", enabledOptions, disabledOptions);
        addOptionName(options.getUseSingleBorderforContiguousCells(), "UseSingleBorderforContiguousCells", enabledOptions, disabledOptions);
        addOptionName(options.getUseWord2002TableStyleRules(), "UseWord2002TableStyleRules", enabledOptions, disabledOptions);
        addOptionName(options.getUseWord2010TableStyleRules(), "UseWord2010TableStyleRules", enabledOptions, disabledOptions);
        addOptionName(options.getUseWord97LineBreakRules(), "UseWord97LineBreakRules", enabledOptions, disabledOptions);
        addOptionName(options.getWPJustification(), "WPJustification", enabledOptions, disabledOptions);
        addOptionName(options.getWPSpaceWidth(), "WPSpaceWidth", enabledOptions, disabledOptions);
        addOptionName(options.getWrapTrailSpaces(), "WrapTrailSpaces", enabledOptions, disabledOptions);
        System.out.println("\tEnabled options:");
        for (String optionName : enabledOptions)
            System.out.println("\t\t{optionName}");
        System.out.println("\tDisabled options:");
        for (String optionName : disabledOptions)
            System.out.println("\t\t{optionName}");
    }

    private static void addOptionName(boolean option, String optionName, ArrayList<String> enabledOptions, ArrayList<String> disabledOptions)
    {
        if (option)
            enabledOptions.add(optionName);
        else
            disabledOptions.add(optionName);
    }
    //ExEnd

    @Test
    public void tables() throws Exception
    {
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

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.Tables.docx");
    }

    @Test
    public void breaks() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getApplyBreakingRules());
        Assert.assertEquals(true, compatibilityOptions.getDoNotUseEastAsianBreakRules());
        Assert.assertEquals(false, compatibilityOptions.getShowBreaksInFrames());
        Assert.assertEquals(true, compatibilityOptions.getSplitPgBreakAndParaMark());
        Assert.assertEquals(true, compatibilityOptions.getUseAltKinsokuLineBreakRules());
        Assert.assertEquals(false, compatibilityOptions.getUseWord97LineBreakRules());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.Breaks.docx");
    }

    @Test
    public void spacing() throws Exception
    {
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

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.Spacing.docx");
    }

    @Test
    public void wordPerfect() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getSuppressTopSpacingWP());
        Assert.assertEquals(false, compatibilityOptions.getTruncateFontHeightsLikeWP6());
        Assert.assertEquals(false, compatibilityOptions.getWPJustification());
        Assert.assertEquals(false, compatibilityOptions.getWPSpaceWidth());
        Assert.assertEquals(false, compatibilityOptions.getWrapTrailSpaces());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.WordPerfect.docx");
    }

    @Test
    public void alignment() throws Exception
    {
        Document doc = new Document();
        
        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(true, compatibilityOptions.getCachedColBalance());
        Assert.assertEquals(true, compatibilityOptions.getDoNotVertAlignInTxbx());
        Assert.assertEquals(true, compatibilityOptions.getDoNotWrapTextWithPunct());
        Assert.assertEquals(false, compatibilityOptions.getNoTabHangInd());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.Alignment.docx");
    }

    @Test
    public void legacy() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(false, compatibilityOptions.getFootnoteLayoutLikeWW8());
        Assert.assertEquals(false, compatibilityOptions.getLineWrapLikeWord6());
        Assert.assertEquals(false, compatibilityOptions.getMWSmallCaps());
        Assert.assertEquals(false, compatibilityOptions.getShapeLayoutLikeWW8());
        Assert.assertEquals(false, compatibilityOptions.getUICompat97To2003());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.Legacy.docx");
    }

    @Test
    public void list() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        Assert.assertEquals(true, compatibilityOptions.getUnderlineTabInNumList());
        Assert.assertEquals(true, compatibilityOptions.getUseNormalStyleForList());

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.List.docx");
    }

    @Test
    public void misc() throws Exception
    {
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

        // In the output document, these settings can be accessed in Microsoft Word via
        // File -> Options -> Advanced -> Compatibility options for...
        doc.save(getArtifactsDir() + "CompatibilityOptions.Misc.docx");
    }
}
