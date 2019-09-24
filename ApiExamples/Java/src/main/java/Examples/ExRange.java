package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;
import java.util.regex.Pattern;

public class ExRange extends ApiExampleBase {
    @Test
    public void replaceSimple() throws Exception {
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
        doc.save(getArtifactsDir() + "Range.ReplaceSimple.docx");
        //ExEnd

        Assert.assertEquals(doc.getText(), "Hello James Bond,\r\f");
    }

    @Test
    public void replaceWithString() throws Exception {
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
    public void replaceWithRegex() throws Exception {
        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace all occurrences of words "sad" or "mad" to "bad".
        Document doc = new Document(getMyDir() + "Document.doc");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad", options);
        //ExEnd
        doc.save(getArtifactsDir() + "ReplaceWithRegex.docx");
    }

    @Test
    public void replaceWithoutPreserveMetaCharacters() throws Exception {
        final String text = "some text";
        final String replaceWithText = "&ldquo;";

        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write(text);

        FindReplaceOptions options = new FindReplaceOptions();
        options.setPreserveMetaCharacters(false);

        doc.getRange().replace(text, replaceWithText, options);

        Assert.assertEquals("\u000bdquo;\f", doc.getText());
    }

    @Test
    public void findAndReplaceWithPreserveMetaCharacters() throws Exception {
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

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:DocumentBuilder.InsertHtml(String)
    //ExSummary:Replaces text specified with regular expression with HTML.
    @Test //ExSkip
    public void replaceWithInsertHtml() throws Exception {
        // Open the document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello <CustomerName>,");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplaceWithHtmlEvaluator(options));

        doc.getRange().replace(Pattern.compile(" <CustomerName>,"), "", options);

        // Save the modified document.
        doc.save(getArtifactsDir() + "Range.ReplaceWithInsertHtml.doc");

        Assert.assertEquals(doc.getText(), "James Bond, Hello\r\f"); //ExSkip
    }

    private class ReplaceWithHtmlEvaluator implements IReplacingCallback {
        ReplaceWithHtmlEvaluator(final FindReplaceOptions options) {
            mOptions = options;
        }

        /**
         * NOTE: This is a simplistic method that will only work well when the match
         * starts at the beginning of a run.
         */
        public int replacing(final ReplacingArgs e) throws Exception {
            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo(e.getMatchNode());

            // Replace '<CustomerName>' text with a red bold name.
            builder.insertHtml("<b><font color='red'>James Bond, </font></b>");
            e.setReplacement("");

            return ReplaceAction.REPLACE;
        }

        private FindReplaceOptions mOptions;
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
    public void replaceNumbersAsHex() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.write("There are few numbers that should be converted to HEX and highlighted: 123, 456, 789 and 17379.");

        FindReplaceOptions options = new FindReplaceOptions();
        // Highlight newly inserted content with a color
        options.getApplyFont().setHighlightColor(new Color(255, 140, 0));
        // Apply an IReplacingCallback to make the replacement to convert integers into hex equivalents
        // and also to count replacements in the order they take place
        options.setReplacingCallback(new NumberHexer());
        // By default, text is searched for replacements front to back, but we can change it to go the other way
        options.setDirection(FindReplaceDirection.BACKWARD);

        int count = doc.getRange().replace(Pattern.compile("[0-9]+"), "", options);
        Assert.assertEquals(count, 4);

        doc.save(getArtifactsDir() + "Range.ReplaceNumbersAsHex.docx");
    }

    /// <summary>
    /// Replaces arabic numbers with hexadecimal equivalents and appends the number of each replacement
    /// </summary>
    private static class NumberHexer implements IReplacingCallback {
        public int replacing(ReplacingArgs args) {
            mCurrentReplacementNumber++;

            // Parse numbers
            String numberStr = args.getMatch().group();
            numberStr = numberStr.trim();
            // Java throws NumberFormatException both for overflow and bad format
            int number = Integer.parseInt(numberStr);

            // And write it as HEX.
            args.setReplacement(MessageFormat.format("0x{0} (replacement #{1})", Integer.toHexString(number), mCurrentReplacementNumber));

            System.out.println(MessageFormat.format("Match #{0}", mCurrentReplacementNumber));
            System.out.println(MessageFormat.format("\tOriginal value:\t{0}", args.getMatch().group()));
            System.out.println(MessageFormat.format("\tReplacement:\t{0}", args.getReplacement()));
            System.out.println(MessageFormat.format("\tOffset in parent {0} node:\t{1}", args.getMatchNode().getNodeType(), args.getMatchOffset()));

            return ReplaceAction.REPLACE;
        }

        private int mCurrentReplacementNumber;
    }
    //ExEnd

    @Test
    public void applyParagraphFormat() throws Exception {
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
        Assert.assertEquals(count, 2);

        doc.save(getArtifactsDir() + "Range.ApplyParagraphFormat.docx");
        //ExEnd
    }

    @Test
    public void deleteSelection() throws Exception {
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
    public void rangesGetText() throws Exception {
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

