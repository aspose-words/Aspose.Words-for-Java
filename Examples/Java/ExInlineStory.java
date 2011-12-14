//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Footnote;
import com.aspose.words.FootnoteType;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import org.testng.Assert;
import com.aspose.words.NodeType;
import com.aspose.words.Comment;

import java.util.Date;


public class ExInlineStory extends ExBase
{
    @Test
    public void addFootnote() throws Exception
    {
        //ExStart
        //ExFor:Footnote
        //ExFor:FootnoteType
        //ExFor:InlineStory
        //ExFor:InlineStory.Paragraphs
        //ExFor:InlineStory.FirstParagraph
        //ExFor:Footnote.#ctor
        //ExSummary:Shows how to add a footnote to a paragraph in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Some text is added.");

        Footnote footnote = new Footnote(doc, FootnoteType.FOOTNOTE);
        builder.getCurrentParagraph().appendChild(footnote);
        footnote.getParagraphs().add(new Paragraph(doc));
        footnote.getFirstParagraph().getRuns().add(new Run(doc, "Footnote text."));
        //ExEnd

        Assert.assertEquals(doc.getChildNodes(NodeType.FOOTNOTE, true).get(0).toTxt().trim(), "Footnote text.");
    }

    @Test
    public void addComment() throws Exception
    {
        //ExStart
        //ExFor:Comment
        //ExFor:InlineStory
        //ExFor:InlineStory.Paragraphs
        //ExFor:InlineStory.FirstParagraph
        //ExFor:Comment.#ctor(DocumentBase, String, String, DateTime)
        //ExSummary:Shows how to add a comment to a paragraph in the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Some text is added.");

        Comment comment = new Comment(doc, "Amy Lee", "AL", new Date());
        builder.getCurrentParagraph().appendChild(comment);
        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));
        //ExEnd

        Assert.assertEquals((doc.getChildNodes(NodeType.COMMENT, true).get(0)).getText(), "Comment text.\r");
    }
}


