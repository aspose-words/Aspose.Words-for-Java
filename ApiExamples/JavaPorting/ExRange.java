// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.msConsole;
import com.aspose.words.FindReplaceOptions;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.ms.System.Convert;
import com.aspose.ms.System.msString;


@Test
public class ExRange extends ApiExampleBase
{
    //>>>>>>>> #region  Replace 

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
        msConsole.writeLine(doc.getFirstSection().getBody().getParagraphs().get(0).getText());

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        // Replace the text in the document.
        doc.getRange().replace("_CustomerName_", "James Bond", options);

        // Save the modified document.
        doc.save(getArtifactsDir() + "Range.ReplaceSimple.docx");
        //ExEnd

        msAssert.areEqual("Hello James Bond,\r\f", doc.getText());
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

        doc.save(getArtifactsDir() + "ReplaceWithString.docx");
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

        doc.getRange().replaceInternal(new Regex("[s|m]ad"), "bad", options);
        //ExEnd

        doc.save(getArtifactsDir() + "ReplaceWithRegex.docx");
    }

    // Note: Need more info from dev.
    @Test
    public void replaceWithoutPreserveMetaCharacters() throws Exception
    {
        final String TEXT = "some text";
        final String REPLACE_WITH_TEXT = "&ldquo;";

        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write(TEXT);

        FindReplaceOptions options = new FindReplaceOptions();
        options.setPreserveMetaCharacters(false);

        doc.getRange().replace(TEXT, REPLACE_WITH_TEXT, options);

        msAssert.areEqual("\u000bdquo;\f", doc.getText());
    }

    @Test
    public void findAndReplaceWithPreserveMetaCharacters() throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.PreserveMetaCharacters
        //ExSummary:Shows how to preserved meta-characters that beginning with "&".
        Document doc = new Document(getMyDir() + "Range.FindAndReplaceWithPreserveMetaCharacters.docx");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(true);
        options.setPreserveMetaCharacters(true);

        doc.getRange().replace("sad", "&ldquo; some text &rdquo;", options);
        //ExEnd

        doc.save(getArtifactsDir() + "Range.FindAndReplaceWithMetacharacters.docx");
    }

    @Test
    public void replaceWithInsertHtml() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExFor:ReplacingArgs.Replacement
        //ExFor:IReplacingCallback
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExFor:DocumentBuilder.InsertHtml(String)
        //ExSummary:Replaces text specified with regular expression with HTML.
        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello <CustomerName>,");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplaceWithHtmlEvaluator(options));

        doc.getRange().replaceInternal(new Regex(" <CustomerName>,"), "", options);

        // Save the modified document.
        doc.save(getArtifactsDir() + "Range.ReplaceWithInsertHtml.doc");

        msAssert.areEqual("James Bond, Hello\r\f", doc.getText()); //ExSkip
    }

    private static class ReplaceWithHtmlEvaluator implements IReplacingCallback
    {
        ReplaceWithHtmlEvaluator(FindReplaceOptions options)
        {
            mOptions = options;
        }

        /// <summary>
        /// NOTE: This is a simplistic method that will only work well when the match
        /// starts at the beginning of a run.
        /// </summary>
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs args) throws Exception
        {
            DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
            builder.moveTo(args.getMatchNode());

            // Replace '<CustomerName>' text with a red bold name.
            builder.insertHtml("<b><font color='red'>James Bond, </font></b>");
            args.setReplacement("");

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
        builder.write(
            "There are few numbers that should be converted to HEX and highlighted: 123, 456, 789 and 17379.");

        FindReplaceOptions options = new FindReplaceOptions();

        // Highlight newly inserted content.
        options.getApplyFont().setHighlightColor(msColor.getDarkOrange());
        options.setReplacingCallback(new NumberHexer());

        int count = doc.getRange().replaceInternal(new Regex("[0-9]+"), "", options);
    }

    // Customer defined callback.
    private static class NumberHexer implements IReplacingCallback
    {
        public /*ReplaceAction*/int replacing(ReplacingArgs args)
        {
            // Parse numbers.
            int number = Convert.toInt32(args.getMatchInternal().getValue());

            // And write it as HEX.
            args.setReplacement(msString.format("0x{0:X}", number));

            return ReplaceAction.REPLACE;
        }
    }

    //<<<<<<<< #endregion 

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
        msConsole.writeLine(doc.getText());

        // Delete the first section from the document.
        doc.getSections().get(0).getRange().delete();

        // Check the first section was deleted by looking at the text of the whole document again.
        msConsole.writeLine(doc.getText());
        //ExEnd

        msAssert.areEqual("Hello2\f", doc.getText());
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
