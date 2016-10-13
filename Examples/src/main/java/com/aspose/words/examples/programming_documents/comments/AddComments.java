package com.aspose.words.examples.programming_documents.comments;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

import java.util.Date;

@SuppressWarnings("unchecked")
public class AddComments {
    public static void main(String[] args) throws Exception {

        String dataDir = Utils.getDataDir(AddComments.class);
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Some text is added.");

        Comment comment = new Comment(doc, "Awais Hafeez", "AH", new Date());
        builder.getCurrentParagraph().appendChild(comment);
        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));
        doc.save(dataDir + "output.doc");

    }
}