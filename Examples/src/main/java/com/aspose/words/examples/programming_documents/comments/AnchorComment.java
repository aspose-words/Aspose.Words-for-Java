package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.Date;

@SuppressWarnings("unchecked")
public class AnchorComment {
    public static void main(String[] args) throws Exception {

        String dataDir = Utils.getDataDir(AnchorComment.class);
        Document doc = new Document();

        Paragraph para1 = new Paragraph(doc);
        Run run1 = new Run(doc, "Some ");
        Run run2 = new Run(doc, "text ");
        para1.appendChild(run1);
        para1.appendChild(run2);
        doc.getFirstSection().getBody().appendChild(para1);

        Paragraph para2 = new Paragraph(doc);
        Run run3 = new Run(doc, "is ");
        Run run4 = new Run(doc, "added ");
        para2.appendChild(run3);
        para2.appendChild(run4);
        doc.getFirstSection().getBody().appendChild(para2);

        Comment comment = new Comment(doc, "Awais Hafeez", "AH", new Date());
        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));

        CommentRangeStart commentRangeStart = new CommentRangeStart(doc, comment.getId());
        CommentRangeEnd commentRangeEnd = new CommentRangeEnd(doc, comment.getId());

        run1.getParentNode().insertAfter(commentRangeStart, run1);
        run3.getParentNode().insertAfter(commentRangeEnd, run3);
        commentRangeEnd.getParentNode().insertAfter(comment, commentRangeEnd);

        doc.save(dataDir + "output.doc");


    }
}