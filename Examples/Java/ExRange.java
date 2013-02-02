//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import org.testng.Assert;
import java.util.regex.Pattern;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Run;
import com.aspose.words.Paragraph;
import java.awt.Color;
import com.aspose.words.Underline;


public class ExRange extends ExBase
{
    @Test
    public void deleteSelection() throws Exception
    {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete a section from a Word document.
        // Open Word document.
        Document doc = new Document(getMyDir() + "Range.DeleteSection.doc");

        // The document contains two sections. Each section has a paragraph of text.
        System.out.println(doc.getText());

        // Delete the first section from the document.
        doc.getSections().get(0).getRange().delete();

        // Check the first section was deleted by looking at the text of the whole document again.
        System.out.println(doc.getText());
        //ExEnd

        Assert.assertEquals(doc.getText(), "Hello2\f");
    }

    @Test
    public void replaceSimple() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(String,String,Boolean,Boolean)
        //ExSummary:Simple find and replace operation.
        // Open the document.
        Document doc = new Document(getMyDir() + "Range.ReplaceSimple.doc");

        // Check the document contains what we are about to test.
        System.out.println(doc.getFirstSection().getBody().getParagraphs().get(0).getText());

        // Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", false, false);

        // Save the modified document.
        doc.save(getMyDir() + "Range.ReplaceSimple Out.doc");
        //ExEnd

        Assert.assertEquals(doc.getText(), "Hello James Bond,\r\f");
    }

    @Test
    public void replaceWithInsertHtmlCaller() throws Exception
    {
        replaceWithInsertHtml();
    }

    //ExStart
    //ExFor:Range.Replace(Regex,IReplacingCallback,Boolean)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:DocumentBuilder.InsertHtml
    //ExSummary:Replaces text specified with regular expression with HTML.
    public void replaceWithInsertHtml() throws Exception
    {
        // Open the document.
        Document doc = new Document(getMyDir() + "Range.ReplaceWithInsertHtml.doc");

        doc.getRange().replace(Pattern.compile("<CustomerName>"), new ReplaceWithHtmlEvaluator(), false);

        // Save the modified document.
        doc.save(getMyDir() + "Range.ReplaceWithInsertHtml Out.doc");

        Assert.assertEquals(doc.getText(), "Hello James Bond,\r\f");  //ExSkip
    }

    private class ReplaceWithHtmlEvaluator implements IReplacingCallback
    {
        /**
         * NOTE: This is a simplistic method that will only work well when the match
         * starts at the beginning of a run.
         */
        public int replacing(ReplacingArgs e) throws Exception
        {
            DocumentBuilder builder = new DocumentBuilder((Document)e.getMatchNode().getDocument());
            builder.moveTo(e.getMatchNode());
            // Replace '<CustomerName>' text with a red bold name.
            builder.insertHtml("<b><font color='red'>James Bond</font></b>");

            e.setReplacement("");
            return ReplaceAction.REPLACE;
        }
    }
    //ExEnd

    @Test
    public void rangesGetText() throws Exception
    {
        //ExStart
        //ExFor:Range
        //ExFor:Range.Text
        //ExId:RangesGetText
        //ExSummary:Shows how to get plain, unformatted text of a range.
        Document doc = new Document(getMyDir() + "Document.doc");
        String text = doc.getRange().getText();
        //ExEnd
    }

    @Test
    public void replaceWithString() throws Exception
    {
        //ExStart
        //ExFor:Range
        //ExId:RangesReplaceString
        //ExSummary:Shows how to replace all occurrences of word "sad" to "bad".
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.getRange().replace("sad", "bad", false, true);
        //ExEnd
        doc.save(getMyDir() + "ReplaceWithString Out.doc");
    }

    @Test
    public void replaceWithRegex() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(Regex, String)
        //ExId:RangesReplaceRegex
        //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad");
        //ExEnd
        doc.save(getMyDir() + "ReplaceWithRegex Out.doc");
    }

    @Test
    public void replaceWithEvaluatorCaller() throws Exception
    {
        replaceWithEvaluator();
    }

    //ExStart
    //ExFor:Range
    //ExFor:ReplacingArgs.Match
    //ExId:RangesReplaceWithReplaceEvaluator
    //ExSummary:Shows how to replace with a custom evaluator.
    public void replaceWithEvaluator() throws Exception
    {
        Document doc = new Document(getMyDir() + "Range.ReplaceWithEvaluator.doc");
        doc.getRange().replace(Pattern.compile("[s|m]ad"), new MyReplaceEvaluator(), true);
        doc.save(getMyDir() + "Range.ReplaceWithEvaluator Out.doc");
    }

    private class MyReplaceEvaluator implements IReplacingCallback
    {
        /**
         * This is called during a replace operation each time a match is found.
         * This method appends a number to the match string and returns it as a replacement string.
         */
        public int replacing(ReplacingArgs e) throws Exception
        {
            e.setReplacement(e.getMatch().group() + Integer.toString(mMatchNumber));
            mMatchNumber++;
            return ReplaceAction.REPLACE;
        }

        private int mMatchNumber;
    }
    //ExEnd

    @Test
    public void rangesDeleteText() throws Exception
    {
        //ExStart
        //ExId:RangesDeleteText
        //ExSummary:Shows how to delete all characters of a range.
        Document doc = new Document(getMyDir() + "Document.doc");
        doc.getSections().get(0).getRange().delete();
        //ExEnd
    }

    /**
     * RK This works, but the logic is so complicated that I don't want to show it to users.
     */
    @Test
    public void changeTextToHyperlinks() throws Exception
    {
        Document doc = new Document(getMyDir() + "Range.ChangeTextToHyperlinks.doc");

        // Create regular expression for URL search
        // Group 1 - protocol
        // Group 2 - domain
        Pattern regexUrl = Pattern.compile("(\\w+):\\/\\/([\\w.]+\\/?)\\S*(?x)");

        // Run replacement, using regular expression and evaluator.
        doc.getRange().replace(regexUrl, new ChangeTextToHyperlinksEvaluator(doc), false);

        // Save updated document.
        doc.save(getMyDir() + "Range.ChangeTextToHyperlinks Out.docx");
    }

    private class ChangeTextToHyperlinksEvaluator implements IReplacingCallback
    {
        ChangeTextToHyperlinksEvaluator(Document doc) throws Exception
        {
            mBuilder = new DocumentBuilder(doc);
        }

        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs e) throws Exception
        {
            // This is the run node that contains the found text. Note that the run might contain other
            // text apart from the URL. All the complexity below is just to handle that. I don't think there
            // is a simpler way at the moment.
            Run run = (Run)e.getMatchNode();

            Paragraph para = run.getParentParagraph();

            String url = e.getMatch().group();

            // We are using \xbf (inverted question mark) symbol for temporary purposes.
            // Any symbol will do that is non-special and is guaranteed not to be presented in the document.
            // The purpose is to split the matched run into two and insert a hyperlink field between them.
            para.getRange().replace(url, "\u00bf", true, true);

            Run subRun = (Run)run.deepClone(false);
            int pos = run.getText().indexOf("\u00bf");
            subRun.setText(subRun.getText().substring(0, pos));
            run.setText(run.getText().substring(pos + 1, run.getText().length()));

            para.getChildNodes().insert(para.getChildNodes().indexOf(run), subRun);

            mBuilder.moveTo(run);

            // Specify font formatting for the hyperlink.
            mBuilder.getFont().setColor(Color.BLUE);
            mBuilder.getFont().setUnderline(Underline.SINGLE);

            // Insert the hyperlink.
            mBuilder.insertHyperlink(url, url, false);

            // Clear hyperlink formatting.
            mBuilder.getFont().clearFormatting();

            // Let's remove run if it is empty.
            if (run.getText().equals(""))
                run.remove();

            // No replace action is necessary - we have already done what we intended to do.
            return ReplaceAction.SKIP;
        }

        private DocumentBuilder mBuilder;
    }
}

