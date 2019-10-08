package com.aspose.words.examples.programming_documents.tableofcontents;

import com.aspose.words.*;

import java.util.regex.Pattern;

public class FindAndInsertATCField {

    public static void main(String[] args) throws Exception {

        //ExStart:FindAndInsertATCField
        Document doc = new Document();

        FindReplaceOptions opts = new FindReplaceOptions();
        opts.setDirection(FindReplaceDirection.BACKWARD);
        opts.setReplacingCallback(new InsertTCFieldHandler("Chapter 1", "\\l 1"));

        // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
        doc.getRange().replace(Pattern.compile("The Beginning"), "", opts);
        //ExEnd:FindAndInsertATCField
    }
}

//ExStart:InsertTCFieldHandler
class InsertTCFieldHandler implements IReplacingCallback {
    // Store the text and switches to be used for the TC fields.
    private String mFieldText;
    private String mFieldSwitches;

    /**
     * The switches to use for each TC field. Can be an empty string or null.
     */
    public InsertTCFieldHandler(String switches) throws Exception {
        this(null, switches);
    }

    /**
     * The display text and the switches to use for each TC field. Display text
     * Can be an empty string or null.
     */
    public InsertTCFieldHandler(String text, String switches) throws Exception {
        mFieldText = text;
        mFieldSwitches = switches;
    }

    public int replacing(ReplacingArgs args) throws Exception {
        // Create a builder to insert the field.
        DocumentBuilder builder = new DocumentBuilder((Document) args.getMatchNode().getDocument());
        // Move to the first node of the match.
        builder.moveTo(args.getMatchNode());

        // If the user specified text to be used in the field as display text then use that, otherwise use the
        // match string as the display text.
        String insertText;

        if (!(mFieldText == null || "".equals(mFieldText)))
            insertText = mFieldText;
        else
            insertText = args.getMatch().group();

        // Insert the TC field before this node using the specified string as the display text and user defined switches.
        builder.insertField(java.text.MessageFormat.format("TC \"{0}\" {1}", insertText, mFieldSwitches));

        // We have done what we want so skip replacement.
        return ReplaceAction.SKIP;
    }

}
//ExEnd:InsertTCFieldHandler
