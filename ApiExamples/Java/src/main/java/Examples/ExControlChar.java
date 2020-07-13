package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;

@Test
public class ExControlChar extends ApiExampleBase {
    @Test
    public void carriageReturn() throws Exception {
        //ExStart
        //ExFor:ControlChar
        //ExFor:ControlChar.Cr
        //ExSummary:Shows how to use control characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert paragraphs with text with DocumentBuilder
        builder.writeln("Hello world!");
        builder.writeln("Hello again!");

        // The entire document, when in string form, will display some structural features such as breaks with control characters
        Assert.assertEquals(MessageFormat.format("Hello world!{0}Hello again!{1}{2}", ControlChar.CR, ControlChar.CR, ControlChar.PAGE_BREAK), doc.getText());

        // Some of them can be trimmed out
        Assert.assertEquals(MessageFormat.format("Hello world!{0}Hello again!", ControlChar.CR), doc.getText().trim());
        //ExEnd
    }

    @Test
    public void insertControlChars() throws Exception {
        //ExStart
        //ExFor:ControlChar.Cell
        //ExFor:ControlChar.ColumnBreak
        //ExFor:ControlChar.CrLf
        //ExFor:ControlChar.Lf
        //ExFor:ControlChar.LineBreak
        //ExFor:ControlChar.LineFeed
        //ExFor:ControlChar.NonBreakingSpace
        //ExFor:ControlChar.PageBreak
        //ExFor:ControlChar.ParagraphBreak
        //ExFor:ControlChar.SectionBreak
        //ExFor:ControlChar.CellChar
        //ExFor:ControlChar.ColumnBreakChar
        //ExFor:ControlChar.DefaultTextInputChar
        //ExFor:ControlChar.FieldEndChar
        //ExFor:ControlChar.FieldStartChar
        //ExFor:ControlChar.FieldSeparatorChar
        //ExFor:ControlChar.LineBreakChar
        //ExFor:ControlChar.LineFeedChar
        //ExFor:ControlChar.NonBreakingHyphenChar
        //ExFor:ControlChar.NonBreakingSpaceChar
        //ExFor:ControlChar.OptionalHyphenChar
        //ExFor:ControlChar.PageBreakChar
        //ExFor:ControlChar.ParagraphBreakChar
        //ExFor:ControlChar.SectionBreakChar
        //ExFor:ControlChar.SpaceChar
        //ExSummary:Shows how to use various control characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a regular space
        builder.write("Before space." + ControlChar.SPACE_CHAR + "After space.");

        // Add a NBSP, or non-breaking space
        // Unlike the regular space, this space can't have an automatic line break at its position
        builder.write("Before space." + ControlChar.NON_BREAKING_SPACE + "After space.");

        // Add a tab character
        builder.write("Before tab." + ControlChar.TAB + "After tab.");

        // Add a line break
        builder.write("Before line break." + ControlChar.LINE_BREAK + "After line break.");

        // This adds a new line and starts a new paragraph
        // Same value as ControlChar.Lf
        Assert.assertEquals(doc.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount(), 1);
        builder.write("Before line feed." + ControlChar.LINE_FEED + "After line feed.");
        Assert.assertEquals(2, doc.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount());

        // Carriage returns and line feeds can be represented together by one character
        Assert.assertEquals(ControlChar.CR_LF, ControlChar.CR + ControlChar.LF);

        // The line feed character has two versions
        Assert.assertEquals(ControlChar.LINE_FEED, ControlChar.LF);

        // Add a paragraph break, also adding a new paragraph
        builder.write("Before paragraph break." + ControlChar.PARAGRAPH_BREAK + "After paragraph break.");
        Assert.assertEquals(doc.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount(), 3);

        // Add a section break. Note that this does not make a new section or paragraph
        Assert.assertEquals(doc.getSections().getCount(), 1);
        builder.write("Before section break." + ControlChar.SECTION_BREAK + "After section break.");
        Assert.assertEquals(doc.getSections().getCount(), 1);

        // A page break is the same value as a section break
        builder.write("Before page break." + ControlChar.PAGE_BREAK + "After page break.");

        // We can add a new section like this
        doc.appendChild(new Section(doc));
        builder.moveToSection(1);

        // If you have a section with more than one column, you can use a column break to make following text start on a new column
        builder.getCurrentSection().getPageSetup().getTextColumns().setCount(2);
        builder.write("Text at end of column 1." + ControlChar.COLUMN_BREAK + "Text at beginning of column 2.");

        // Save document to see the characters we added
        doc.save(getArtifactsDir() + "ControlChar.InsertControlChars.docx");

        // There are char and string counterparts for most characters
        Assert.assertEquals(ControlChar.CELL.toCharArray()[0], ControlChar.CELL_CHAR);
        Assert.assertEquals(ControlChar.NON_BREAKING_SPACE.toCharArray()[0], ControlChar.NON_BREAKING_SPACE_CHAR);
        Assert.assertEquals(ControlChar.TAB.toCharArray()[0], ControlChar.TAB_CHAR);
        Assert.assertEquals(ControlChar.LINE_BREAK.toCharArray()[0], ControlChar.LINE_BREAK_CHAR);
        Assert.assertEquals(ControlChar.LINE_FEED.toCharArray()[0], ControlChar.LINE_FEED_CHAR);
        Assert.assertEquals(ControlChar.PARAGRAPH_BREAK.toCharArray()[0], ControlChar.PARAGRAPH_BREAK_CHAR);
        Assert.assertEquals(ControlChar.SECTION_BREAK.toCharArray()[0], ControlChar.SECTION_BREAK_CHAR);
        Assert.assertEquals(ControlChar.PAGE_BREAK.toCharArray()[0], ControlChar.SECTION_BREAK_CHAR);
        Assert.assertEquals(ControlChar.COLUMN_BREAK.toCharArray()[0], ControlChar.COLUMN_BREAK_CHAR);
        //ExEnd
    }
}
