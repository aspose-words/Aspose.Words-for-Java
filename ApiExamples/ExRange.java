//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;
import org.testng.Assert;

import java.text.MessageFormat;
import java.util.regex.Pattern;
import java.awt.Color;

public class ExRange extends ApiExampleBase
{

    @Test
    public void replaceSimple() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.MatchCase
        //ExFor:FindReplaceOptions.FindWholeWordsOnly
        //ExSummary:Simple find and replace operation.
        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello _CustomerName_,");

        // Check the document contains what we are about to test.
        System.out.println(doc.getFirstSection().getBody().getParagraphs().get(0).getText());

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        // Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", options);

        // Save the modified document.
        doc.save(getMyDir() + "\\Artifacts\\Range.ReplaceSimple.docx");
        //ExEnd

        Assert.assertEquals(doc.getText(), "Hello James Bond,\r\f");
    }

    @Test
    public void replaceWithString() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("This one is sad.");
        builder.writeln("That one is mad.");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(true);

        doc.getRange().replace("sad", "bad", options);

        doc.save(getMyDir() + "\\Artifacts\\ReplaceWithString.docx");
    }

    @Test
    public void replaceWithRegex() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
        Document doc = new Document(getMyDir() + "Document.doc");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad", options);
        //ExEnd
        doc.save(getMyDir() + "\\Artifacts\\ReplaceWithRegex.docx");
    }

    @Test
    public void replaceWithInsertHtmlCaller() throws Exception
    {
        replaceWithInsertHtml();
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExSummary:Replaces text specified with regular expression with HTML.
    public void replaceWithInsertHtml() throws Exception
    {
        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello <CustomerName>,");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplaceWithHtmlEvaluator(options));

        doc.getRange().replace(Pattern.compile(" <CustomerName>,"), "", options);

        // Save the modified document.
        doc.save(getMyDir() + "\\Artifacts\\Range.ReplaceWithInsertHtml.doc");

        Assert.assertEquals(doc.getText(), "James Bond, Hello\r\f"); //ExSkip
    }

    private class ReplaceWithHtmlEvaluator implements IReplacingCallback
    {
        ReplaceWithHtmlEvaluator(FindReplaceOptions options)
        {
            mOptions = options;
        }

        /**
         * NOTE: This is a simplistic method that will only work well when the match
         * starts at the beginning of a run.
         */
        public int replacing(ReplacingArgs e) throws Exception
        {
            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo(e.getMatchNode());

            // Replace '<CustomerName>' text with a red bold name.
            builder.insertHtml("<b><font color='red'>James Bond, </font></b>");
            e.setReplacement("");

            return ReplaceAction.REPLACE;
        }

        private /*final*/ FindReplaceOptions mOptions;
    }
    //ExEnd

    @Test
    public void replaceNumbersAsHex() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.write("There are few numbers that should be converted to HEX and highlighted: 123, 456, 789 and 17379.");

        FindReplaceOptions options = new FindReplaceOptions();

        // Highlight newly inserted content.
        options.getApplyFont().setHighlightColor(new Color(255, 140, 0));
        options.setReplacingCallback(new NumberHexer());

        int count = doc.getRange().replace(Pattern.compile("[0-9]+"), "", options);
    }

    // Customer defined callback.
    private class NumberHexer implements IReplacingCallback
    {
        public /*ReplaceAction*/int replacing(ReplacingArgs args)
        {
            // Parse numbers.
            String numberStr = args.getMatch().group();
            numberStr = numberStr.trim();
            // Java throws NumberFormatException both for overflow and bad format.
            int number = Integer.parseInt(numberStr);

            // And write it as HEX.
            args.setReplacement(String.format("0x{0:X}", number));

            return ReplaceAction.REPLACE;
        }
    }

    @Test
    public void deleteSelection() throws Exception
    {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete all characters of a range.
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
}

