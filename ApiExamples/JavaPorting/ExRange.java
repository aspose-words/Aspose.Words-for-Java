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
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.msConsole;
import com.aspose.words.FindReplaceOptions;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.ms.System.msString;
import com.aspose.ms.System.Text.RegularExpressions.Regex;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.FindReplaceDirection;
import com.aspose.ms.System.Convert;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.BreakType;


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
        // Open the document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello _CustomerName_,");

        // Check the document contains what we are about to test
        msConsole.writeLine(doc.getFirstSection().getBody().getParagraphs().get(0).getText());

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        // Replace the text in the document
        doc.getRange().replace("_CustomerName_", "James Bond", options);

        // Save the modified document
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

        doc.save(getArtifactsDir() + "Range.ReplaceWithString.docx");
    }

    @Test
    public void replaceWithRegex() throws Exception
    {
        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("sad mad bad");

        msAssert.areEqual("sad mad bad", msString.trim(doc.getText()));

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        doc.getRange().replaceInternal(new Regex("[s|m]ad"), "bad", options);

        msAssert.areEqual("bad bad bad", msString.trim(doc.getText()));
        //ExEnd
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
        //ExSummary:Shows how to preserved meta-characters that begin with "&".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("one");
        builder.writeln("two");
        builder.writeln("three");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(true);
        options.setPreserveMetaCharacters(true);

        doc.getRange().replace("two", "&ldquo; four &rdquo;", options);
        //ExEnd

        doc.save(getArtifactsDir() + "Range.FindAndReplaceWithMetacharacters.docx");
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExSummary:Replaces text specified with regular expression with HTML.
    @Test //ExSkip
    public void replaceWithInsertHtml() throws Exception
    {
        // Open the document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello <CustomerName>,");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplaceWithHtmlEvaluator(options));

        doc.getRange().replaceInternal(new Regex(" <CustomerName>,"), "", options);

        // Save the modified document
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

            // Replace '<CustomerName>' text with a red bold name
            builder.insertHtml("<b><font color='red'>James Bond, </font></b>");
            args.setReplacement("");

            return ReplaceAction.REPLACE;
        }

        private /*final*/ FindReplaceOptions mOptions;
    }
    //ExEnd

    //ExStart
    //ExFor:FindReplaceOptions.ApplyFont
    //ExFor:FindReplaceOptions.Direction
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:ReplacingArgs.GroupIndex
    //ExFor:ReplacingArgs.GroupName
    //ExFor:ReplacingArgs.Match
    //ExFor:ReplacingArgs.MatchOffset
    //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    @Test //ExSkip
    public void replaceNumbersAsHex() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.write(
            "There are few numbers that should be converted to HEX and highlighted: 123, 456, 789 and 17379.");

        FindReplaceOptions options = new FindReplaceOptions();

        // Highlight newly inserted content with a color
        options.getApplyFont().setHighlightColor(msColor.getLightGray());

        // Apply an IReplacingCallback to make the replacement to convert integers into hex equivalents
        // and also to count replacements in the order they take place
        options.setReplacingCallback(new NumberHexer());

        // By default, text is searched for replacements front to back, but we can change it to go the other way
        options.setDirection(FindReplaceDirection.BACKWARD);

        int count = doc.getRange().replaceInternal(new Regex("[0-9]+"), "", options);
        msAssert.areEqual(4, count);

        doc.save(getArtifactsDir() + "Range.ReplaceNumbersAsHex.docx");
    }

    /// <summary>
    /// Replaces arabic numbers with hexadecimal equivalents and appends the number of each replacement.
    /// </summary>
    private static class NumberHexer implements IReplacingCallback
    {
        public /*ReplaceAction*/int replacing(ReplacingArgs args)
        {
            mCurrentReplacementNumber++;
            
            // Parse numbers
            int number = Convert.toInt32(args.getMatchInternal().getValue());

            // And write it as HEX
            args.setReplacement("0x{number:X} (replacement #{mCurrentReplacementNumber})");

            msConsole.writeLine($"Match #{mCurrentReplacementNumber}");
            msConsole.writeLine($"\tOriginal value:\t{args.Match.Value}");
            msConsole.writeLine($"\tReplacement:\t{args.Replacement}");
            msConsole.writeLine($"\tOffset in parent {args.MatchNode.NodeType} node:\t{args.MatchOffset}");

            msConsole.writeLine(msString.isNullOrEmpty(args.GroupName)
                ? $"\tGroup index:\t{args.GroupIndex}"
                : $"\tGroup name:\t{args.GroupName}");

            return ReplaceAction.REPLACE;
        }

        private int mCurrentReplacementNumber;
    }
    //ExEnd

    //<<<<<<<< #endregion 

    @Test
    public void applyParagraphFormat() throws Exception
    {
        //ExStart
        //ExFor:FindReplaceOptions.ApplyParagraphFormat
        //ExSummary:Shows how to affect the format of paragraphs with successful replacements.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Every paragraph that ends with a full stop like this one will be right aligned.");
        builder.writeln("This one will not!");
        builder.writeln("And this one will.");
        
        FindReplaceOptions options = new FindReplaceOptions();
        options.getApplyParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        int count = doc.getRange().replace(".&p", "!&p", options);
        msAssert.areEqual(2, count);

        doc.save(getArtifactsDir() + "Range.ApplyParagraphFormat.docx");
        //ExEnd
    }

    @Test
    public void deleteSelection() throws Exception
    {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete all characters of a range.
        // Insert two sections into a blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1. ");
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.write("Section 2.");

        // Verify the whole text of the document
        msAssert.areEqual("Section 1. \fSection 2.", msString.trim(doc.getText()));

        // Delete the first section from the document
        doc.getSections().get(0).getRange().delete();

        // Check the first section was deleted by looking at the text of the whole document again
        msAssert.areEqual("Section 2.", msString.trim(doc.getText()));
        //ExEnd
    }

    @Test
    public void rangesGetText() throws Exception
    {
        //ExStart
        //ExFor:Range
        //ExFor:Range.Text
        //ExSummary:Shows how to get plain, unformatted text of a range.
        Document doc = new Document(getMyDir() + "Document.docx");
        String text = doc.getRange().getText();
        //ExEnd
    }
}
