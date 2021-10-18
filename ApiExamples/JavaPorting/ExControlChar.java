// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import org.testng.Assert;
import com.aspose.words.ControlChar;
import com.aspose.words.NodeType;
import com.aspose.words.Section;
import com.aspose.ms.System.Convert;


@Test
public class ExControlChar extends ApiExampleBase
{
    @Test
    public void carriageReturn() throws Exception
    {
        //ExStart
        //ExFor:ControlChar
        //ExFor:ControlChar.Cr
        //ExFor:Node.GetText
        //ExSummary:Shows how to use control characters.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert paragraphs with text with DocumentBuilder.
        builder.writeln("Hello world!");
        builder.writeln("Hello again!");

        // Converting the document to text form reveals that control characters
        // represent some of the document's structural elements, such as page breaks.
        Assert.assertEquals($"Hello world!{ControlChar.Cr}" +
                        $"Hello again!{ControlChar.Cr}" +
                        ControlChar.PAGE_BREAK, doc.getText());

        // When converting a document to string form,
        // we can omit some of the control characters with the Trim method.
        Assert.assertEquals($"Hello world!{ControlChar.Cr}" +
                        "Hello again!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void insertControlChars() throws Exception
    {
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
        //ExSummary:Shows how to add various control characters to a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a regular space.
        builder.write("Before space." + ControlChar.SPACE_CHAR + "After space.");

        // Add an NBSP, which is a non-breaking space.
        // Unlike the regular space, this space cannot have an automatic line break at its position.
        builder.write("Before space." + ControlChar.NON_BREAKING_SPACE + "After space.");

        // Add a tab character.
        builder.write("Before tab." + ControlChar.TAB + "After tab.");

        // Add a line break.
        builder.write("Before line break." + ControlChar.LINE_BREAK + "After line break.");

        // Add a new line and starts a new paragraph.
        Assert.assertEquals(1, doc.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount());
        builder.write("Before line feed." + ControlChar.LINE_FEED + "After line feed.");
        Assert.assertEquals(2, doc.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount());

        // The line feed character has two versions.
        Assert.assertEquals(ControlChar.LINE_FEED, ControlChar.LF);

        // Carriage returns and line feeds can be represented together by one character.
        Assert.assertEquals(ControlChar.CR_LF, ControlChar.CR + ControlChar.LF);

        // Add a paragraph break, which will start a new paragraph.
        builder.write("Before paragraph break." + ControlChar.PARAGRAPH_BREAK + "After paragraph break.");
        Assert.assertEquals(3, doc.getFirstSection().getBody().getChildNodes(NodeType.PARAGRAPH, true).getCount());

        // Add a section break. This does not make a new section or paragraph.
        Assert.assertEquals(1, doc.getSections().getCount());
        builder.write("Before section break." + ControlChar.SECTION_BREAK + "After section break.");
        Assert.assertEquals(1, doc.getSections().getCount());

        // Add a page break.
        builder.write("Before page break." + ControlChar.PAGE_BREAK + "After page break.");

        // A page break is the same value as a section break.
        Assert.assertEquals(ControlChar.PAGE_BREAK, ControlChar.SECTION_BREAK);

        // Insert a new section, and then set its column count to two.
        doc.appendChild(new Section(doc));
        builder.moveToSection(1);
        builder.getCurrentSection().getPageSetup().getTextColumns().setCount(2);

        // We can use a control character to mark the point where text moves to the next column.
        builder.write("Text at end of column 1." + ControlChar.COLUMN_BREAK + "Text at beginning of column 2.");

        doc.save(getArtifactsDir() + "ControlChar.InsertControlChars.docx");

        // There are char and string counterparts for most characters.
        Assert.assertEquals(Convert.toChar(ControlChar.CELL), ControlChar.CELL_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.NON_BREAKING_SPACE), ControlChar.NON_BREAKING_SPACE_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.TAB), ControlChar.TAB_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.LINE_BREAK), ControlChar.LINE_BREAK_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.LINE_FEED), ControlChar.LINE_FEED_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.PARAGRAPH_BREAK), ControlChar.PARAGRAPH_BREAK_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.SECTION_BREAK), ControlChar.SECTION_BREAK_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.PAGE_BREAK), ControlChar.SECTION_BREAK_CHAR);
        Assert.assertEquals(Convert.toChar(ControlChar.COLUMN_BREAK), ControlChar.COLUMN_BREAK_CHAR);
        //ExEnd
    }
}
