package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.regex.Pattern;

public class ExRange extends ApiExampleBase {
    @Test
    public void replace() throws Exception {
        //ExStart
        //ExFor:Range.Replace(String, String)
        //ExSummary:Shows how to perform a find-and-replace text operation on the contents of a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Greetings, _FullName_!");

        // Perform a find-and-replace operation on our document's contents and verify the number of replacements that took place.
        int replacementCount = doc.getRange().replace("_FullName_", "John Doe");

        Assert.assertEquals(1, replacementCount);
        Assert.assertEquals("Greetings, John Doe!", doc.getText().trim());
        //ExEnd
    }

    @Test(dataProvider = "replaceMatchCaseDataProvider")
    public void replaceMatchCase(boolean matchCase) throws Exception {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.MatchCase
        //ExSummary:Shows how to toggle case sensitivity when performing a find-and-replace operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Ruby bought a ruby necklace.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "MatchCase" flag to "true" to apply case sensitivity while finding strings to replace.
        // Set the "MatchCase" flag to "false" to ignore character case while searching for text to replace.
        options.setMatchCase(matchCase);

        doc.getRange().replace("Ruby", "Jade", options);

        Assert.assertEquals(matchCase ? "Jade bought a ruby necklace." : "Jade bought a Jade necklace.",
                doc.getText().trim());
        //ExEnd
    }

    @DataProvider(name = "replaceMatchCaseDataProvider")
    public static Object[][] replaceMatchCaseDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "replaceFindWholeWordsOnlyDataProvider")
    public void replaceFindWholeWordsOnly(boolean findWholeWordsOnly) throws Exception {
        //ExStart
        //ExFor:Range.Replace(String, String, FindReplaceOptions)
        //ExFor:FindReplaceOptions
        //ExFor:FindReplaceOptions.FindWholeWordsOnly
        //ExSummary:Shows how to toggle standalone word-only find-and-replace operations. 
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Jackson will meet you in Jacksonville.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "FindWholeWordsOnly" flag to "true" to replace the found text if it is not a part of another word.
        // Set the "FindWholeWordsOnly" flag to "false" to replace all text regardless of its surroundings.
        options.setFindWholeWordsOnly(findWholeWordsOnly);

        doc.getRange().replace("Jackson", "Louis", options);

        Assert.assertEquals(
                findWholeWordsOnly ? "Louis will meet you in Jacksonville." : "Louis will meet you in Louisville.",
                doc.getText().trim());
        //ExEnd
    }

    @DataProvider(name = "replaceFindWholeWordsOnlyDataProvider")
    public static Object[][] replaceFindWholeWordsOnlyDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "ignoreDeletedDataProvider")
    public void ignoreDeleted(boolean ignoreTextInsideDeleteRevisions) throws Exception {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreDeleted
        //ExSummary:Shows how to include or ignore text inside delete revisions during a find-and-replace operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.writeln("Hello again!");

        // Start tracking revisions and remove the second paragraph, which will create a delete revision.
        // That paragraph will persist in the document until we accept the delete revision.
        doc.startTrackRevisions("John Doe", new Date());
        doc.getFirstSection().getBody().getParagraphs().get(1).remove();
        doc.stopTrackRevisions();

        Assert.assertTrue(doc.getFirstSection().getBody().getParagraphs().get(1).isDeleteRevision());

        // We can use a "FindReplaceOptions" object to modify the find and replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "IgnoreDeleted" flag to "true" to get the find-and-replace
        // operation to ignore paragraphs that are delete revisions.
        // Set the "IgnoreDeleted" flag to "false" to get the find-and-replace
        // operation to also search for text inside delete revisions.
        options.setIgnoreDeleted(ignoreTextInsideDeleteRevisions);

        doc.getRange().replace("Hello", "Greetings", options);

        Assert.assertEquals(
                ignoreTextInsideDeleteRevisions
                        ? "Greetings world!\rHello again!"
                        : "Greetings world!\rGreetings again!", doc.getText().trim());
        //ExEnd
    }

    @DataProvider(name = "ignoreDeletedDataProvider")
    public static Object[][] ignoreDeletedDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(dataProvider = "ignoreInsertedDataProvider")
    public void ignoreInserted(boolean ignoreTextInsideInsertRevisions) throws Exception {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreInserted
        //ExSummary:Shows how to include or ignore text inside insert revisions during a find-and-replace operation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");

        // Start tracking revisions and insert a paragraph. That paragraph will be an insert revision.
        doc.startTrackRevisions("John Doe", new Date());
        builder.writeln("Hello again!");
        doc.stopTrackRevisions();

        Assert.assertTrue(doc.getFirstSection().getBody().getParagraphs().get(1).isInsertRevision());

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "IgnoreInserted" flag to "true" to get the find-and-replace
        // operation to ignore paragraphs that are insert revisions.
        // Set the "IgnoreInserted" flag to "false" to get the find-and-replace
        // operation to also search for text inside insert revisions.
        options.setIgnoreInserted(ignoreTextInsideInsertRevisions);

        doc.getRange().replace("Hello", "Greetings", options);

        Assert.assertEquals(
                ignoreTextInsideInsertRevisions
                        ? "Greetings world!\rHello again!"
                        : "Greetings world!\rGreetings again!", doc.getText().trim());
        //ExEnd
    }

    @DataProvider(name = "ignoreInsertedDataProvider")
    public static Object[][] ignoreInsertedDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test(dataProvider = "ignoreFieldsDataProvider")
    public void ignoreFields(boolean ignoreTextInsideFields) throws Exception {
        //ExStart
        //ExFor:FindReplaceOptions.IgnoreFields
        //ExSummary:Shows how to ignore text inside fields.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello world!");
        builder.insertField("QUOTE", "Hello again!");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "IgnoreFields" flag to "true" to get the find-and-replace
        // operation to ignore text inside fields.
        // Set the "IgnoreFields" flag to "false" to get the find-and-replace
        // operation to also search for text inside fields.
        options.setIgnoreFields(ignoreTextInsideFields);

        doc.getRange().replace("Hello", "Greetings", options);

        Assert.assertEquals(
                ignoreTextInsideFields
                        ? "Greetings world!\r\u0013QUOTE\u0014Hello again!"
                        : "Greetings world!\r\u0013QUOTE\u0014Greetings again!", doc.getText().trim());
        //ExEnd
    }

    @DataProvider(name = "ignoreFieldsDataProvider")
    public static Object[][] ignoreFieldsDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    @Test
    public void updateFieldsInRange() throws Exception {
        //ExStart
        //ExFor:Range.UpdateFields
        //ExSummary:Shows how to update all the fields in a range.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertField(" DOCPROPERTY Category");
        builder.insertBreak(BreakType.SECTION_BREAK_EVEN_PAGE);
        builder.insertField(" DOCPROPERTY Category");

        // The above DOCPROPERTY fields will display the value of this built-in document property.
        doc.getBuiltInDocumentProperties().setCategory("MyCategory");

        // If we update the value of a document property, we will need to update all the DOCPROPERTY fields to display it.
        Assert.assertEquals("", doc.getRange().getFields().get(0).getResult());
        Assert.assertEquals("", doc.getRange().getFields().get(1).getResult());

        // Update all the fields that are in the range of the first section.
        doc.getFirstSection().getRange().updateFields();

        Assert.assertEquals("MyCategory", doc.getRange().getFields().get(0).getResult());
        Assert.assertEquals("", doc.getRange().getFields().get(1).getResult());
        //ExEnd
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

        doc.save(getArtifactsDir() + "Range.ReplaceWithString.docx");
    }

    @Test
    public void replaceWithRegex() throws Exception {
        //ExStart
        //ExFor:Range.Replace(Regex, String)
        //ExSummary:Shows how to replace all occurrences of a regular expression pattern with other text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("I decided to get the curtains in gray, ideal for the grey-accented room.");

        doc.getRange().replace(Pattern.compile("gr(a|e)y"), "lavender");

        Assert.assertEquals("I decided to get the curtains in lavender, ideal for the lavender-accented room.", doc.getText().trim());
        //ExEnd
    }

    //ExStart
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:ReplacingArgs.Replacement
    //ExFor:IReplacingCallback
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExSummary:Shows how to replace all occurrences of a regular expression pattern with another string, while tracking all such replacements.
    @Test //ExSkip
    public void replaceWithCallback() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Our new location in New York City is opening tomorrow. " +
                "Hope to see all our NYC-based customers at the opening!");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set a callback that tracks any replacements that the "Replace" method will make.
        TextFindAndReplacementLogger logger = new TextFindAndReplacementLogger();
        options.setReplacingCallback(logger);

        doc.getRange().replace(Pattern.compile("New York City|NYC"), "Washington", options);

        Assert.assertEquals("Our new location in (Old value:\"New York City\") Washington is opening tomorrow. " +
                "Hope to see all our (Old value:\"NYC\") Washington-based customers at the opening!", doc.getText().trim());

        Assert.assertEquals("\"New York City\" converted to \"Washington\" 20 characters into a 21 node." +
                "\"NYC\" converted to \"Washington\" 42 characters into a 21 node.", logger.getLog().trim());
    }

    /// <summary>
    /// Maintains a log of every text replacement done by a find-and-replace operation
    /// and notes the original matched text's value.
    /// </summary>
    private static class TextFindAndReplacementLogger implements IReplacingCallback {
        public int replacing(ReplacingArgs args) {
            mLog.append(MessageFormat.format("\"{0}\" converted to \"{1}\" {2} characters into a {3} node.", args.getMatch().group(0), args.getReplacement(), args.getMatchOffset(), args.getMatchNode().getNodeType()));

            args.setReplacement(MessageFormat.format("(Old value:\"{0}\") {1}", args.getMatch().group(0), args.getReplacement()));
            return ReplaceAction.REPLACE;
        }

        public String getLog() {
            return mLog.toString();
        }

        private final StringBuilder mLog = new StringBuilder();
    }
    //ExEnd

    //ExStart
    //ExFor:FindReplaceOptions.ApplyFont
    //ExFor:FindReplaceOptions.ReplacingCallback
    //ExFor:ReplacingArgs.GroupIndex
    //ExFor:ReplacingArgs.GroupName
    //ExFor:ReplacingArgs.Match
    //ExFor:ReplacingArgs.MatchOffset
    //ExSummary:Shows how to apply a different font to new content via FindReplaceOptions.
    @Test //ExSkip
    public void convertNumbersToHexadecimal() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\n" +
                "123, 456, 789 and 17379.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "HighlightColor" property to a background color that we want to apply to the operation's resulting text.
        options.getApplyFont().setHighlightColor(Color.GRAY);

        NumberHexer numberHexer = new NumberHexer();
        options.setReplacingCallback(numberHexer);

        int replacementCount = doc.getRange().replace(Pattern.compile("[0-9]+"), "", options);

        System.out.println(numberHexer.getLog());

        Assert.assertEquals(4, replacementCount);
        Assert.assertEquals("Numbers that the find-and-replace operation will convert to hexadecimal and highlight:\r" +
                "0x123, 0x456, 0x789 and 0x17,379.", doc.getText().trim());
    }

    /// <summary>
    /// Replaces numeric find-and-replacement matches with their hexadecimal equivalents.
    /// Maintains a log of every replacement.
    /// </summary>
    private static class NumberHexer implements IReplacingCallback {
        public int replacing(ReplacingArgs args) {
            mCurrentReplacementNumber++;

            int number = Integer.parseInt(args.getMatch().group(0));

            args.setReplacement(MessageFormat.format("0x{0}", number));

            mLog.append(MessageFormat.format("Match #{0}", mCurrentReplacementNumber));
            mLog.append(MessageFormat.format("\tOriginal value:\t{0}", args.getMatch().group(0)));
            mLog.append(MessageFormat.format("\tReplacement:\t{0}", args.getReplacement()));
            mLog.append(MessageFormat.format("\tOffset in parent {0} node:\t{1}", args.getMatchNode().getNodeType(), args.getMatchOffset()));

            return ReplaceAction.REPLACE;
        }

        public String getLog() {
            return mLog.toString();
        }

        private int mCurrentReplacementNumber;
        private final StringBuilder mLog = new StringBuilder();
    }
    //ExEnd

    @Test
    public void applyParagraphFormat() throws Exception {
        //ExStart
        //ExFor:FindReplaceOptions.ApplyParagraphFormat
        //ExFor:Range.Replace(String, String)
        //ExSummary:Shows how to add formatting to paragraphs in which a find-and-replace operation has found matches.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Every paragraph that ends with a full stop like this one will be right aligned.");
        builder.writeln("This one will not!");
        builder.write("This one also will.");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(ParagraphAlignment.LEFT, paragraphs.get(0).getParagraphFormat().getAlignment());
        Assert.assertEquals(ParagraphAlignment.LEFT, paragraphs.get(1).getParagraphFormat().getAlignment());
        Assert.assertEquals(ParagraphAlignment.LEFT, paragraphs.get(2).getParagraphFormat().getAlignment());

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "Alignment" property to "ParagraphAlignment.Right" to right-align every paragraph
        // that contains a match that the find-and-replace operation finds.
        options.getApplyParagraphFormat().setAlignment(ParagraphAlignment.RIGHT);

        // Replace every full stop that is right before a paragraph break with an exclamation point.
        int count = doc.getRange().replace(".&p", "!&p", options);

        Assert.assertEquals(2, count);
        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(0).getParagraphFormat().getAlignment());
        Assert.assertEquals(ParagraphAlignment.LEFT, paragraphs.get(1).getParagraphFormat().getAlignment());
        Assert.assertEquals(ParagraphAlignment.RIGHT, paragraphs.get(2).getParagraphFormat().getAlignment());
        Assert.assertEquals("Every paragraph that ends with a full stop like this one will be right aligned!\r" +
                "This one will not!\r" +
                "This one also will!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void deleteSelection() throws Exception {
        //ExStart
        //ExFor:Node.Range
        //ExFor:Range.Delete
        //ExSummary:Shows how to delete all the nodes from a range.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add text to the first section in the document, and then add another section.
        builder.write("Section 1. ");
        builder.insertBreak(BreakType.SECTION_BREAK_CONTINUOUS);
        builder.write("Section 2.");

        Assert.assertEquals("Section 1. \fSection 2.", doc.getText().trim());

        // Remove the first section entirely by removing all the nodes
        // within its range, including the section itself.
        doc.getSections().get(0).getRange().delete();

        Assert.assertEquals(1, doc.getSections().getCount());
        Assert.assertEquals("Section 2.", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void rangesGetText() throws Exception {
        //ExStart
        //ExFor:Range
        //ExFor:Range.Text
        //ExSummary:Shows how to get the text contents of all the nodes that a range covers.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        Assert.assertEquals("Hello world!", doc.getRange().getText().trim());
        //ExEnd
    }

    //ExStart
    //ExFor:FindReplaceOptions.UseLegacyOrder
    //ExSummary:Shows how to change the searching order of nodes when performing a find-and-replace text operation.
    @Test(dataProvider = "useLegacyOrderDataProvider") //ExSkip
    public void useLegacyOrder(boolean useLegacyOrder) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three runs which we can search for using a regex pattern.
        // Place one of those runs inside a text box.
        builder.writeln("[tag 1]");
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 50.0);
        builder.writeln("[tag 2]");
        builder.moveTo(textBox.getFirstParagraph());
        builder.write("[tag 3]");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Assign a custom callback to the "ReplacingCallback" property.
        TextReplacementTracker callback = new TextReplacementTracker();
        options.setReplacingCallback(callback);

        // If we set the "UseLegacyOrder" property to "true", the
        // find-and-replace operation will go through all the runs outside of a text box
        // before going through the ones inside a text box.
        // If we set the "UseLegacyOrder" property to "false", the
        // find-and-replace operation will go over all the runs in a range in sequential order.
        options.setUseLegacyOrder(useLegacyOrder);

        doc.getRange().replace("\\[tag d*\\]", "", options);
    }

    @DataProvider(name = "useLegacyOrderDataProvider")
    public static Object[][] useLegacyOrderDataProvider() {
        return new Object[][]
                {
                        {true},
                        {false},
                };
    }

    /// <summary>
    /// Records the order of all matches that occur during a find-and-replace operation.
    /// </summary>
    private static class TextReplacementTracker implements IReplacingCallback {
        public int replacing(ReplacingArgs e) {
            mMatches.add(e.getMatch().group(1));
            return ReplaceAction.REPLACE;
        }

        public ArrayList<String> getMatches() {
            return mMatches;
        }

        private ArrayList<String> mMatches;
    }
    //ExEnd

    @Test(dataProvider = "useSubstitutionsDataProvider")
    public void useSubstitutions(boolean useSubstitutions) throws Exception {
        //ExStart
        //ExFor:FindReplaceOptions.UseSubstitutions
        //ExSummary:Shows how to replace the text with substitutions.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("John sold a car to Paul.");
        builder.writeln("Jane sold a house to Joe.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Set the "UseSubstitutions" property to "true" to get
        // the find-and-replace operation to recognize substitution elements.
        // Set the "UseSubstitutions" property to "false" to ignore substitution elements.
        options.setUseSubstitutions(useSubstitutions);

        doc.getRange().replace(Pattern.compile("([A-z]+) sold a ([A-z]+) to ([A-z]+)"), "$3 bought a $2 from $1", options);

        Assert.assertEquals(
                useSubstitutions
                        ? "Paul bought a car from John.\rJoe bought a house from Jane."
                        : "$3 bought a $2 from $1.\r$3 bought a $2 from $1.", doc.getText().trim());
        //ExEnd
    }

    @DataProvider(name = "useSubstitutionsDataProvider")
    public static Object[][] useSubstitutionsDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    //ExStart
    //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
    //ExFor:IReplacingCallback
    //ExFor:ReplaceAction
    //ExFor:IReplacingCallback.Replacing
    //ExFor:ReplacingArgs
    //ExFor:ReplacingArgs.MatchNode
    //ExSummary:Shows how to insert an entire document's contents as a replacement of a match in a find-and-replace operation.
    @Test //ExSkip
    public void insertDocumentAtReplace() throws Exception {
        Document mainDoc = new Document(getMyDir() + "Document insertion destination.docx");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new InsertDocumentAtReplaceHandler());

        mainDoc.getRange().replace(Pattern.compile("\\[MY_DOCUMENT\\]"), "", options);
        mainDoc.save(getArtifactsDir() + "InsertDocument.InsertDocumentAtReplace.docx");

        testInsertDocumentAtReplace(new Document(getArtifactsDir() + "InsertDocument.InsertDocumentAtReplace.docx")); //ExSkip
    }

    private static class InsertDocumentAtReplaceHandler implements IReplacingCallback {
        public int replacing(ReplacingArgs args) throws Exception {
            Document subDoc = new Document(getMyDir() + "Document.docx");

            // Insert a document after the paragraph containing the matched text.
            Paragraph para = (Paragraph) args.getMatchNode().getParentNode();
            insertDocument(para, subDoc);

            // Remove the paragraph with the matched text.
            para.remove();

            return ReplaceAction.SKIP;
        }
    }

    /// <summary>
    /// Inserts all the nodes of another document after a paragraph or table.
    /// </summary>
    private static void insertDocument(Node insertionDestination, Document docToInsert) {
        if (((insertionDestination.getNodeType()) == (NodeType.PARAGRAPH)) || ((insertionDestination.getNodeType()) == (NodeType.TABLE))) {
            CompositeNode dstStory = insertionDestination.getParentNode();

            NodeImporter importer =
                    new NodeImporter(docToInsert, insertionDestination.getDocument(), ImportFormatMode.KEEP_SOURCE_FORMATTING);

            for (Section srcSection : docToInsert.getSections())
                for (Node srcNode : srcSection.getBody()) {
                    // Skip the node if it is the last empty paragraph in a section.
                    if (((srcNode.getNodeType()) == (NodeType.PARAGRAPH))) {
                        Paragraph para = (Paragraph) srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }

                    Node newNode = importer.importNode(srcNode, true);

                    dstStory.insertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
        } else {
            throw new IllegalArgumentException("The destination node must be either a paragraph or table.");
        }
    }
    //ExEnd

    private void testInsertDocumentAtReplace(Document doc) {
        Assert.assertEquals("1) At text that can be identified by regex:\rHello World!\r" +
                "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                "3) At a bookmark:", doc.getFirstSection().getBody().getText().trim());
    }

    //ExStart
    //ExFor:FindReplaceOptions.Direction
    //ExFor:FindReplaceDirection
    //ExSummary:Shows how to determine which direction a find-and-replace operation traverses the document in.
    @Test(dataProvider = "directionDataProvider") //ExSkip
    public void direction(int findReplaceDirection) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three runs which we can search for using a regex pattern.
        // Place one of those runs inside a text box.
        builder.writeln("Match 1.");
        builder.writeln("Match 2.");
        builder.writeln("Match 3.");
        builder.writeln("Match 4.");

        // We can use a "FindReplaceOptions" object to modify the find-and-replace process.
        FindReplaceOptions options = new FindReplaceOptions();

        // Assign a custom callback to the "ReplacingCallback" property.
        TextReplacementRecorder callback = new TextReplacementRecorder();
        options.setReplacingCallback(callback);

        // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
        // operation to start from the end of the range, and traverse back to the beginning.
        // Set the "Direction" property to "FindReplaceDirection.Backward" to get the find-and-replace
        // operation to start from the beginning of the range, and traverse to the end.
        options.setDirection(findReplaceDirection);

        doc.getRange().replace(Pattern.compile("Match \\d*"), "Replacement", options);

        Assert.assertEquals("Replacement.\r" +
                "Replacement.\r" +
                "Replacement.\r" +
                "Replacement.", doc.getText().trim());

        switch (findReplaceDirection) {
            case FindReplaceDirection.FORWARD:
                Assert.assertEquals(new String[]{"Match 1", "Match 2", "Match 3", "Match 4"}, callback.getMatches().toArray());
                break;
            case FindReplaceDirection.BACKWARD:
                Assert.assertEquals(new String[]{"Match 4", "Match 3", "Match 2", "Match 1"}, callback.getMatches().toArray());
                break;
        }
    }

    @DataProvider(name = "directionDataProvider")
    public static Object[][] directionDataProvider() {
        return new Object[][]
                {
                        {FindReplaceDirection.BACKWARD},
                        {FindReplaceDirection.FORWARD},
                };
    }

    /// <summary>
    /// Records all matches that occur during a find-and-replace operation in the order that they take place.
    /// </summary>
    private static class TextReplacementRecorder implements IReplacingCallback {
        public int replacing(ReplacingArgs e) {
            mMatches.add(e.getMatch().group(0));
            return ReplaceAction.REPLACE;
        }

        public ArrayList<String> getMatches() {
            return mMatches;
        }

        private final ArrayList<String> mMatches = new ArrayList<String>();
    }
    //ExEnd
}
