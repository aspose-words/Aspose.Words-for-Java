package DocsExamples.Programming_with_documents.Contents_managment;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Shape;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.regex.Pattern;

@Test
public class FindAndReplace extends DocsExamplesBase
{
    @Test
    public void simpleFindReplace() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello _CustomerName_,");
        System.out.println("Original document text: " + doc.getRange().getText());

        doc.getRange().replace("_CustomerName_", "James Bond", new FindReplaceOptions(FindReplaceDirection.FORWARD));

        System.out.println("Document text after replace: " + doc.getRange().getText());

        // Save the modified document
        doc.save(getArtifactsDir() + "FindAndReplace.SimpleFindReplace.docx");
    }

    @Test
    public void findAndHighlight() throws Exception
    {
        //ExStart:FindAndHighlight
        Document doc = new Document(getMyDir() + "Find and highlight.docx");

        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setReplacingCallback(new ReplaceEvaluatorFindAndHighlight()); options.setDirection(FindReplaceDirection.BACKWARD);
        }

        Pattern regex = Pattern.compile("your document");
        doc.getRange().replace(regex, "", options);

        doc.save(getArtifactsDir() + "FindAndReplace.FindAndHighlight.docx");
        //ExEnd:FindAndHighlight
    }

    //ExStart:ReplaceEvaluatorFindAndHighlight
    private static class ReplaceEvaluatorFindAndHighlight implements IReplacingCallback
    {
        /// <summary>
        /// This method is called by the Aspose.Words find and replace engine for each match.
        /// This method highlights the match string, even if it spans multiple runs.
        /// </summary>
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs e)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.getMatchNode();

            // The first (and may be the only) run can contain text before the match, 
            // in this case it is necessary to split the run.
            if (e.getMatchOffset() > 0)
                currentNode = splitRun((Run) currentNode, e.getMatchOffset());

            // This array is used to store all nodes of the match for further highlighting.
            ArrayList<Run> runs = new ArrayList<Run>();

            // Find all runs that contain parts of the match string.
            int remainingLength = e.getMatch().group().length();
            while (
                remainingLength > 0 &&
                currentNode != null &&
                currentNode.getText().length() <= remainingLength)
            {
                runs.add((Run) currentNode);
                remainingLength -= currentNode.getText().length();

                // Select the next Run node.
                // Have to loop because there could be other nodes such as BookmarkStart etc.
                do
                {
                    currentNode = currentNode.getNextSibling();
                } while (currentNode != null && currentNode.getNodeType() != NodeType.RUN);
            }

            // Split the last run that contains the match if there is any text left.
            if (currentNode != null && remainingLength > 0)
            {
                splitRun((Run) currentNode, remainingLength);
                runs.add((Run) currentNode);
            }

            // Now highlight all runs in the sequence.
            for (Run run : runs)
                run.getFont().setHighlightColor(Color.YELLOW);

            // Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.SKIP;
        }
    }
    //ExEnd:ReplaceEvaluatorFindAndHighlight

    //ExStart:SplitRun
    /// <summary>
    /// Splits text of the specified run into two runs.
    /// Inserts the new run just after the specified run.
    /// </summary>
    private static Run splitRun(Run run, int position)
    {
        Run afterRun = (Run) run.deepClone(true);
        afterRun.setText(run.getText().substring(position));

        run.setText(run.getText().substring((0), (0) + (position)));
        run.getParentNode().insertAfter(afterRun, run);
        
        return afterRun;
    }
    //ExEnd:SplitRun

    @Test
    public void metaCharactersInSearchPattern() throws Exception
    {
        /* meta-characters
            &p - paragraph break
            &b - section break
            &m - page break
            &l - manual line break
            */

        //ExStart:MetaCharactersInSearchPattern
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("This is Line 1");
        builder.writeln("This is Line 2");

        doc.getRange().replace("This is Line 1&pThis is Line 2", "This is replaced line");

        builder.moveToDocumentEnd();
        builder.write("This is Line 1");
        builder.insertBreak(BreakType.PAGE_BREAK);
        builder.writeln("This is Line 2");

        doc.getRange().replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.");

        doc.save(getArtifactsDir() + "FindAndReplace.MetaCharactersInSearchPattern.docx");
        //ExEnd:MetaCharactersInSearchPattern
    }

    @Test
    public void replaceTextContainingMetaCharacters() throws Exception
    {
        //ExStart:ReplaceTextContainingMetaCharacters
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getFont().setName("Arial");
        builder.writeln("First section");
        builder.writeln("  1st paragraph");
        builder.writeln("  2nd paragraph");
        builder.writeln("{insert-section}");
        builder.writeln("Second section");
        builder.writeln("  1st paragraph");

        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
        findReplaceOptions.getApplyParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        // Double each paragraph break after word "section", add kind of underline and make it centered.
        int count = doc.getRange().replace("section&p", "section&p----------------------&p", findReplaceOptions);

        // Insert section break instead of custom text tag.
        count = doc.getRange().replace("{insert-section}", "&b", findReplaceOptions);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceTextContainingMetaCharacters.docx");
        //ExEnd:ReplaceTextContainingMetaCharacters
    }

    @Test
    public void ignoreTextInsideFields() throws Exception
    {
        //ExStart:IgnoreTextInsideFields
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert field with text inside.
        builder.insertField("INCLUDETEXT", "Text in field");
        
        FindReplaceOptions options = new FindReplaceOptions(); { options.setIgnoreFields(true); }
        
        Pattern regex = Pattern.compile("e");
        doc.getRange().replace(regex, "*", options);
        
        System.out.println(doc.getText());

        options.setIgnoreFields(false);
        doc.getRange().replace(regex, "*", options);
        
        System.out.println(doc.getText());
        //ExEnd:IgnoreTextInsideFields
    }

    @Test
    public void ignoreTextInsideDeleteRevisions() throws Exception
    {
        //ExStart:IgnoreTextInsideDeleteRevisions
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert non-revised text.
        builder.writeln("Deleted");
        builder.write("Text");

        // Remove first paragraph with tracking revisions.
        doc.startTrackRevisions("author", new Date());
        doc.getFirstSection().getBody().getFirstParagraph().remove();
        doc.stopTrackRevisions();

        FindReplaceOptions options = new FindReplaceOptions(); { options.setIgnoreDeleted(true); }

        Pattern regex = Pattern.compile("e");
        doc.getRange().replace(regex, "*", options);

        System.out.println(doc.getText());

        options.setIgnoreDeleted(false);
        doc.getRange().replace(regex, "*", options);

        System.out.println(doc.getText());
        //ExEnd:IgnoreTextInsideDeleteRevisions
    }

    @Test
    public void ignoreTextInsideInsertRevisions() throws Exception
    {
        //ExStart:IgnoreTextInsideInsertRevisions
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text with tracking revisions.
        doc.startTrackRevisions("author", new Date());
        builder.writeln("Inserted");
        doc.stopTrackRevisions();

        // Insert non-revised text.
        builder.write("Text");

        FindReplaceOptions options = new FindReplaceOptions(); { options.setIgnoreInserted(true); }

        Pattern regex = Pattern.compile("e");
        doc.getRange().replace(regex, "*", options);
        
        System.out.println(doc.getText());

        options.setIgnoreInserted(false);
        doc.getRange().replace(regex, "*", options);
        
        System.out.println(doc.getText());
        //ExEnd:IgnoreTextInsideInsertRevisions
    }

    @Test
    public void replaceHtmlTextWithMetaCharacters() throws Exception
    {
        //ExStart:ReplaceHtmlTextWithMetaCharacters
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("{PLACEHOLDER}");

        FindReplaceOptions findReplaceOptions = new FindReplaceOptions(); { findReplaceOptions.setReplacingCallback(new FindAndInsertHtml()); }

        doc.getRange().replace("{PLACEHOLDER}", "<p>&ldquo;Some Text&rdquo;</p>", findReplaceOptions);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceHtmlTextWithMetaCharacters.docx");
        //ExEnd:ReplaceHtmlTextWithMetaCharacters
    }

    //ExStart:ReplaceHtmlFindAndInsertHtml
    public final static class FindAndInsertHtml implements IReplacingCallback
    {
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs e) throws Exception
        {
            Node currentNode = e.getMatchNode();

            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo(currentNode);
            builder.insertHtml(e.getReplacement());

            currentNode.remove();

            return ReplaceAction.SKIP;
        }
    }
    //ExEnd:ReplaceHtmlFindAndInsertHtml

    @Test
    public void replaceTextInFooter() throws Exception
    {
        //ExStart:ReplaceTextInFooter
        Document doc = new Document(getMyDir() + "Footer.docx");

        HeaderFooterCollection headersFooters = doc.getFirstSection().getHeadersFooters();
        HeaderFooter footer = headersFooters.getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);

        FindReplaceOptions options = new FindReplaceOptions(); { options.setMatchCase(false); options.setFindWholeWordsOnly(false); }

        footer.getRange().replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2020 by Aspose Pty Ltd.", options);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceTextInFooter.docx");
        //ExEnd:ReplaceTextInFooter
    }

    @Test
    //ExStart:ShowChangesForHeaderAndFooterOrders
    public void showChangesForHeaderAndFooterOrders() throws Exception
    {
        ReplaceLog logger = new ReplaceLog();
        
        Document doc = new Document(getMyDir() + "Footer.docx");
        Section firstPageSection = doc.getFirstSection();
        
        FindReplaceOptions options = new FindReplaceOptions(); { options.setReplacingCallback(logger); }

        doc.getRange().replace(Pattern.compile("(header|footer)"), "", options);
        
        doc.save(getArtifactsDir() + "FindAndReplace.ShowChangesForHeaderAndFooterOrders.docx");

        logger.clearText();

        firstPageSection.getPageSetup().setDifferentFirstPageHeaderFooter(false);

        doc.getRange().replace(Pattern.compile("(header|footer)"), "", options);
    }

    private static class ReplaceLog implements IReplacingCallback
    {
        public int replacing(ReplacingArgs args)
        {
            mTextBuilder.append(args.getMatchNode().getText());
            return ReplaceAction.SKIP;
        }

        void clearText()
        {
            mTextBuilder.setLength(0);
        }

        private StringBuilder mTextBuilder = new StringBuilder();
    }
    //ExEnd:ShowChangesForHeaderAndFooterOrders

    @Test
    public void replaceTextWithField() throws Exception
    {
        Document doc = new Document(getMyDir() + "Replace text with fields.docx");

        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setReplacingCallback(new ReplaceTextWithFieldHandler(FieldType.FIELD_MERGE_FIELD));
        }

        doc.getRange().replace(Pattern.compile("PlaceHolder(\\d+)"), "", options);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceTextWithField.docx");
    }


    public static class ReplaceTextWithFieldHandler implements IReplacingCallback
    {
        public ReplaceTextWithFieldHandler(int type)
        {
            mFieldType = type;
        }

        public int replacing(ReplacingArgs args) throws Exception {
            ArrayList<Run> runs = findAndSplitMatchRuns(args);

            DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
            builder.moveTo(runs.get(runs.size() - 1));

            // Calculate the field's name from the FieldType enumeration by removing
            // the first instance of "Field" from the text. This works for almost all of the field types.
            String fieldName = FieldType.toString(mFieldType).toUpperCase().substring(5);

            // Insert the field into the document using the specified field type and the matched text as the field name.
            // If the fields you are inserting do not require this extra parameter, it can be removed from the string below.
            builder.insertField(MessageFormat.format("{0} {1}", fieldName, args.getMatch().group(0)));

            for (Run run : runs)
                run.remove();

            return ReplaceAction.SKIP;
        }

        /// <summary>
        /// Finds and splits the match runs and returns them in an List.
        /// </summary>
        public ArrayList<Run> findAndSplitMatchRuns(ReplacingArgs args)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = args.getMatchNode();

            // The first (and may be the only) run can contain text before the match, 
            // In this case it is necessary to split the run.
            if (args.getMatchOffset() > 0)
                currentNode = splitRun((Run) currentNode, args.getMatchOffset());

            // This array is used to store all nodes of the match for further removing.
            ArrayList<Run> runs = new ArrayList<Run>();

            // Find all runs that contain parts of the match string.
            int remainingLength = args.getMatch().group().length();
            while (
                remainingLength > 0 &&
                currentNode != null &&
                currentNode.getText().length() <= remainingLength)
            {
                runs.add((Run) currentNode);
                remainingLength -= currentNode.getText().length();

                do
                {
                    currentNode = currentNode.getNextSibling();
                } while (currentNode != null && currentNode.getNodeType() != NodeType.RUN);
            }

            // Split the last run that contains the match if there is any text left.
            if (currentNode != null && remainingLength > 0)
            {
                splitRun((Run) currentNode, remainingLength);
                runs.add((Run) currentNode);
            }

            return runs;
        }

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private Run splitRun(Run run, int position)
        {
            Run afterRun = (Run) run.deepClone(true);
            
            afterRun.setText(run.getText().substring(position));
            run.setText(run.getText().substring((0), (0) + (position)));
            
            run.getParentNode().insertAfter(afterRun, run);
            
            return afterRun;
        }

        private /*final*/ /*FieldType*/int mFieldType;
    }

    @Test
    public void replaceWithEvaluator() throws Exception
    {
        //ExStart:ReplaceWithEvaluator
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("sad mad bad");

        FindReplaceOptions options = new FindReplaceOptions(); { options.setReplacingCallback(new MyReplaceEvaluator()); }

        doc.getRange().replace(Pattern.compile("[s|m]ad"), "", options);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceWithEvaluator.docx");
        //ExEnd:ReplaceWithEvaluator
    }

    //ExStart:MyReplaceEvaluator
    private static class MyReplaceEvaluator implements IReplacingCallback
    {
        /// <summary>
        /// This is called during a replace operation each time a match is found.
        /// This method appends a number to the match string and returns it as a replacement string.
        /// </summary>
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs e)
        {
            e.setReplacement(e.getMatch() + Integer.toString(mMatchNumber));
            mMatchNumber++;
            
            return ReplaceAction.REPLACE;
        }

        private int mMatchNumber;
    }
    //ExEnd:MyReplaceEvaluator

    @Test
    //ExStart:ReplaceWithHtml
    public void replaceWithHtml() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello <CustomerName>,");

        FindReplaceOptions options = new FindReplaceOptions();
        options.setReplacingCallback(new ReplaceWithHtmlEvaluator(options));

        doc.getRange().replace(Pattern.compile(" <CustomerName>,"), "", options);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceWithHtml.docx");
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
    //ExEnd:ReplaceWithHtml

    @Test
    public void replaceWithRegex() throws Exception
    {
        //ExStart:ReplaceWithRegex
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("sad mad bad");

        FindReplaceOptions options = new FindReplaceOptions();

        doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad", options);

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceWithRegex.docx");
        //ExEnd:ReplaceWithRegex
    }
    
    @Test
    public void recognizeAndSubstitutionsWithinReplacementPatterns() throws Exception
    {
        //ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Jason give money to Paul.");

        Pattern regex = Pattern.compile("([A-z]+) give money to ([A-z]+)");

        FindReplaceOptions options = new FindReplaceOptions(); { options.setUseSubstitutions(true); }

        doc.getRange().replace(regex, "$2 take money from $1", options);
        //ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
    }

    @Test
    public void replaceWithString() throws Exception
    {
        //ExStart:ReplaceWithString
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.writeln("sad mad bad");

        doc.getRange().replace("sad", "bad", new FindReplaceOptions(FindReplaceDirection.FORWARD));

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceWithString.docx");
        //ExEnd:ReplaceWithString
    }

    @Test
    //ExStart:UsingLegacyOrder
    public void usingLegacyOrder() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("[tag 1]");
        Shape textBox = builder.insertShape(ShapeType.TEXT_BOX, 100.0, 50.0);
        builder.writeln("[tag 3]");

        builder.moveTo(textBox.getFirstParagraph());
        builder.write("[tag 2]");

        FindReplaceOptions options = new FindReplaceOptions();
        {
            options.setReplacingCallback(new ReplacingCallback()); options.setUseLegacyOrder(true);
        }

        doc.getRange().replace(Pattern.compile("\\[(.*?)\\]"), "", options);

        doc.save(getArtifactsDir() + "FindAndReplace.UsingLegacyOrder.docx");
    }

    private static class ReplacingCallback implements IReplacingCallback
    {
        public /*ReplaceAction*/int /*IReplacingCallback.*/replacing(ReplacingArgs e)
        {
            System.out.println(e.getMatch().group());
            return ReplaceAction.REPLACE;
        }
    }
    //ExEnd:UsingLegacyOrder

    @Test
    public void replaceTextInTable() throws Exception
    {
        //ExStart:ReplaceText
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

        table.getRange().replace("Carrots", "Eggs", new FindReplaceOptions(FindReplaceDirection.FORWARD));
        table.getLastRow().getLastCell().getRange().replace("50", "20", new FindReplaceOptions(FindReplaceDirection.FORWARD));

        doc.save(getArtifactsDir() + "FindAndReplace.ReplaceTextInTable.docx");
        //ExEnd:ReplaceText
    }
}

