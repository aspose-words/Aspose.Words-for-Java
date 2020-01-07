// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.Convert;
import java.lang.Class;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;


@Test
public class ExCompatibilityOptions extends ApiExampleBase
{
    //ExStart
    //ExFor:Compatibility
    //ExFor:CompatibilityOptions
    //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
    //ExFor:Document.CompatibilityOptions
    //ExSummary:Shows how to optimize document for different word versions.
    @Test //ExSkip
    public void compatibilityOptionsOptimizeFor() throws Exception
    {
        // Create a blank document and get its CompatibilityOptions object
        Document doc = new Document();
        CompatibilityOptions options = doc.getCompatibilityOptions();

        // By default, the CompatibilityOptions will contain the set of values printed below
        msConsole.writeLine("\nDefault optimization settings:");
        printCompatibilityOptions(options);

        // These attributes can be accessed in the output document via File > Options > Advanced > Compatibility for...
        doc.save(getArtifactsDir() + "DefaultCompatibility.docx");

        // We can use the OptimizeFor method to set these values automatically
        // for maximum compatibility with some Microsoft Word versions
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2010);
        msConsole.writeLine("\nOptimized for Word 2010:");
        printCompatibilityOptions(options);

        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2000);
        msConsole.writeLine("\nOptimized for Word 2000:");
        printCompatibilityOptions(options);
    }

    /// <summary>
    /// Prints all options of a CompatibilityOptions object and indicates whether they are enabled or disabled
    /// </summary>
    private static void printCompatibilityOptions(CompatibilityOptions options)
    {
        for (int i = 1; i >= 0; i--)
        {
            msConsole.writeLine(Convert.toBoolean(i) ? "\tEnabled options:" : "\tDisabled options:");
            SortedSet<String> optionNames = new SortedSet<String>();

            for (PropertyDescriptor descriptor : (Iterable<PropertyDescriptor>) TypeDescriptor.GetProperties(options))
            {
                if (descriptor.PropertyType == Class.GetType("System.Boolean") && i == Convert.toInt32(descriptor.GetValue(options)))
                {
                    optionNames.Add(descriptor.Name);
                }
            }

            for (String s : optionNames)
            {
                msConsole.writeLine($"\t\t{s}");
            }
        }
    }
    //ExEnd

    @Test
    public void compatibilityOptionsTable() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2002);

        msAssert.areEqual(false, compatibilityOptions.getAdjustLineHeightInTable());
        msAssert.areEqual(false, compatibilityOptions.getAlignTablesRowByRow());
        msAssert.areEqual(true, compatibilityOptions.getAllowSpaceOfSameStyleInTable());
        msAssert.areEqual(true, compatibilityOptions.getDoNotAutofitConstrainedTables());
        msAssert.areEqual(true, compatibilityOptions.getDoNotBreakConstrainedForcedTable());
        msAssert.areEqual(false, compatibilityOptions.getDoNotBreakWrappedTables());
        msAssert.areEqual(false, compatibilityOptions.getDoNotSnapToGridInCell());
        msAssert.areEqual(false, compatibilityOptions.getDoNotUseHTMLParagraphAutoSpacing());
        msAssert.areEqual(true, compatibilityOptions.getDoNotVertAlignCellWithSp());
        msAssert.areEqual(false, compatibilityOptions.getForgetLastTabAlignment());
        msAssert.areEqual(true, compatibilityOptions.getGrowAutofit());
        msAssert.areEqual(false, compatibilityOptions.getLayoutRawTableWidth());
        msAssert.areEqual(false, compatibilityOptions.getLayoutTableRowsApart());
        msAssert.areEqual(false, compatibilityOptions.getNoColumnBalance());
        msAssert.areEqual(false, compatibilityOptions.getOverrideTableStyleFontSizeAndJustification());
        msAssert.areEqual(false, compatibilityOptions.getUseSingleBorderforContiguousCells());
        msAssert.areEqual(true, compatibilityOptions.getUseWord2002TableStyleRules());
        msAssert.areEqual(false, compatibilityOptions.getUseWord2010TableStyleRules());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsTable.docx");
    }

    @Test
    public void compatibilityOptionsBreaks() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(false, compatibilityOptions.getApplyBreakingRules());
        msAssert.areEqual(true, compatibilityOptions.getDoNotUseEastAsianBreakRules());
        msAssert.areEqual(false, compatibilityOptions.getShowBreaksInFrames());
        msAssert.areEqual(true, compatibilityOptions.getSplitPgBreakAndParaMark());
        msAssert.areEqual(true, compatibilityOptions.getUseAltKinsokuLineBreakRules());
        msAssert.areEqual(false, compatibilityOptions.getUseWord97LineBreakRules());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsBreaks.docx");
    }

    @Test
    public void compatibilityOptionsSpacing() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(false, compatibilityOptions.getAutoSpaceLikeWord95());
        msAssert.areEqual(true, compatibilityOptions.getDisplayHangulFixedWidth());
        msAssert.areEqual(false, compatibilityOptions.getNoExtraLineSpacing());
        msAssert.areEqual(false, compatibilityOptions.getNoLeading());
        msAssert.areEqual(false, compatibilityOptions.getNoSpaceRaiseLower());
        msAssert.areEqual(false, compatibilityOptions.getSpaceForUL());
        msAssert.areEqual(false, compatibilityOptions.getSpacingInWholePoints());
        msAssert.areEqual(false, compatibilityOptions.getSuppressBottomSpacing());
        msAssert.areEqual(false, compatibilityOptions.getSuppressSpBfAfterPgBrk());
        msAssert.areEqual(false, compatibilityOptions.getSuppressSpacingAtTopOfPage());
        msAssert.areEqual(false, compatibilityOptions.getSuppressTopSpacing());
        msAssert.areEqual(false, compatibilityOptions.getUlTrailSpace());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsSpacing.docx");
    }

    @Test
    public void compatibilityOptionsWordPerfect() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(false, compatibilityOptions.getSuppressTopSpacingWP());
        msAssert.areEqual(false, compatibilityOptions.getTruncateFontHeightsLikeWP6());
        msAssert.areEqual(false, compatibilityOptions.getWPJustification());
        msAssert.areEqual(false, compatibilityOptions.getWPSpaceWidth());
        msAssert.areEqual(false, compatibilityOptions.getWrapTrailSpaces());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsWordPerfect.docx");
    }

    @Test
    public void compatibilityOptionsAlignment() throws Exception
    {
        Document doc = new Document();
        
        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(true, compatibilityOptions.getCachedColBalance());
        msAssert.areEqual(true, compatibilityOptions.getDoNotVertAlignInTxbx());
        msAssert.areEqual(true, compatibilityOptions.getDoNotWrapTextWithPunct());
        msAssert.areEqual(false, compatibilityOptions.getNoTabHangInd());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsAlignment.docx");
    }

    @Test
    public void compatibilityOptionsLegacy() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(false, compatibilityOptions.getFootnoteLayoutLikeWW8());
        msAssert.areEqual(false, compatibilityOptions.getLineWrapLikeWord6());
        msAssert.areEqual(false, compatibilityOptions.getMWSmallCaps());
        msAssert.areEqual(false, compatibilityOptions.getShapeLayoutLikeWW8());
        msAssert.areEqual(false, compatibilityOptions.getUICompat97To2003());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsLegacy.docx");
    }

    @Test
    public void compatibilityOptionsList() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(true, compatibilityOptions.getUnderlineTabInNumList());
        msAssert.areEqual(true, compatibilityOptions.getUseNormalStyleForList());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsList.docx");
    }

    @Test
    public void compatibilityOptionsMisc() throws Exception
    {
        Document doc = new Document();

        CompatibilityOptions compatibilityOptions = doc.getCompatibilityOptions();
        compatibilityOptions.optimizeFor(MsWordVersion.WORD_2000);

        msAssert.areEqual(false, compatibilityOptions.getBalanceSingleByteDoubleByteWidth());
        msAssert.areEqual(false, compatibilityOptions.getConvMailMergeEsc());
        msAssert.areEqual(false, compatibilityOptions.getDoNotExpandShiftReturn());
        msAssert.areEqual(false, compatibilityOptions.getDoNotLeaveBackslashAlone());
        msAssert.areEqual(false, compatibilityOptions.getDoNotSuppressParagraphBorders());
        msAssert.areEqual(true, compatibilityOptions.getDoNotUseIndentAsNumberingTabStop());
        msAssert.areEqual(false, compatibilityOptions.getPrintBodyTextBeforeHeader());
        msAssert.areEqual(false, compatibilityOptions.getPrintColBlack());
        msAssert.areEqual(true, compatibilityOptions.getSelectFldWithFirstOrLastChar());
        msAssert.areEqual(false, compatibilityOptions.getSubFontBySize());
        msAssert.areEqual(false, compatibilityOptions.getSwapBordersFacingPgs());
        msAssert.areEqual(false, compatibilityOptions.getTransparentMetafiles());
        msAssert.areEqual(true, compatibilityOptions.getUseAnsiKerningPairs());
        msAssert.areEqual(false, compatibilityOptions.getUseFELayout());
        msAssert.areEqual(false, compatibilityOptions.getUsePrinterMetrics());

        // These options will become available in File > Options > Advanced > Compatibility Options in the output document
        doc.save(getArtifactsDir() + "CompatibilityOptionsMisc.docx");
    }
}
