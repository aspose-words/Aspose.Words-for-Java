package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

@Test
public class WorkingWithComments extends DocsExamplesBase
{
    @Test
    public void addComments() throws Exception
    {
        //ExStart:AddComments
        //ExStart:CreateSimpleDocumentUsingDocumentBuilder
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Some text is added.");
        //ExEnd:CreateSimpleDocumentUsingDocumentBuilder
        
        Comment comment = new Comment(doc, "Awais Hafeez", "AH", new Date());

        builder.getCurrentParagraph().appendChild(comment);

        comment.getParagraphs().add(new Paragraph(doc));
        comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));

        doc.save(getArtifactsDir() + "WorkingWithComments.AddComments.docx");
        //ExEnd:AddComments
    }

    @Test
    public void anchorComment() throws Exception
    {
        //ExStart:AnchorComment
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

        doc.save(getArtifactsDir() + "WorkingWithComments.AnchorComment.doc");
        //ExEnd:AnchorComment
    }

    @Test
    public void addRemoveCommentReply() throws Exception
    {
        //ExStart:AddRemoveCommentReply
        Document doc = new Document(getMyDir() + "Comments.docx");

        Comment comment = (Comment) doc.getChild(NodeType.COMMENT, 0, true);
        comment.removeReply(comment.getReplies().get(0));

        Calendar calendar = new GregorianCalendar(2017, Calendar.SEPTEMBER, 25);
        calendar.set(Calendar.HOUR, 12);
        calendar.set(Calendar.MINUTE, 15);
        calendar.set(Calendar.SECOND, 0);

        comment.addReply("John Doe", "JD", calendar.getTime(), "New reply");

        doc.save(getArtifactsDir() + "WorkingWithComments.AddRemoveCommentReply.docx");
        //ExEnd:AddRemoveCommentReply
    }

    @Test
    public void processComments() throws Exception
    {
        //ExStart:ProcessComments
        Document doc = new Document(getMyDir() + "Comments.docx");

        // Extract the information about the comments of all the authors.
        for (String comment : extractComments(doc))
            System.out.println(comment);

        // Remove comments by the "pm" author.
        removeComments(doc, "pm");
        System.out.println("Comments from \"pm\" are removed!");

        // Extract the information about the comments of the "ks" author.
        for (String comment : extractComments(doc, "ks"))
            System.out.println(comment);

        // Read the comment's reply and resolve them.
        commentResolvedAndReplies(doc);

        // Remove all comments.
        removeComments(doc);
        System.out.println("All comments are removed!");

        doc.save(getArtifactsDir() + "WorkingWithComments.ProcessComments.docx");
        //ExEnd:ProcessComments
    }

    //ExStart:ExtractComments
    private ArrayList<String> extractComments(Document doc) throws Exception
    {
        ArrayList<String> collectedComments = new ArrayList<String>();
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        for (Comment comment : (Iterable<Comment>) comments)
        {
            collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " +
                                  comment.toString(SaveFormat.TEXT));
        }

        return collectedComments;
    }
    //ExEnd:ExtractComments

    //ExStart:ExtractCommentsByAuthor
    private ArrayList<String> extractComments(Document doc, String authorName) throws Exception
    {
        ArrayList<String> collectedComments = new ArrayList<String>();
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        for (Comment comment : (Iterable<Comment>) comments)
        {
            if (comment.getAuthor().equals(authorName))
                collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " +
                                      comment.toString(SaveFormat.TEXT));
        }

        return collectedComments;
    }
    //ExEnd:ExtractCommentsByAuthor

    //ExStart:RemoveComments
    private void removeComments(Document doc)
    {
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        comments.clear();
    }
    //ExEnd:RemoveComments

    //ExStart:RemoveCommentsByAuthor
    private void removeComments(Document doc, String authorName)
    {
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        // Look through all comments and remove those written by the authorName.
        for (int i = comments.getCount() - 1; i >= 0; i--)
        {
            Comment comment = (Comment) comments.get(i);
            if (comment.getAuthor().equals(authorName))
                comment.remove();
        }
    }
    //ExEnd:RemoveCommentsByAuthor

    //ExStart:CommentResolvedandReplies
    private void commentResolvedAndReplies(Document doc)
    {
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);

        Comment parentComment = (Comment) comments.get(0);
        for (Comment childComment : parentComment.getReplies())
        {
            // Get comment parent and status.
            System.out.println(childComment.getAncestor().getId());
            System.out.println(childComment.getDone());

            // And update comment Done mark.
            childComment.setDone(true);
        }
    }
    //ExEnd:CommentResolvedandReplies
}

