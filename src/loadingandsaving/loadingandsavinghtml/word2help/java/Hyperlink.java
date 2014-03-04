/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.word2help.java;

import com.aspose.words.*;

import java.util.regex.Matcher;
import java.util.regex.Pattern;


/**
 * This "facade" class makes it easier to work with a hyperlink field in a Word document.
 *
 * A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words
 * consists of several nodes and it might be difficult to work with all those nodes directly.
 * This is a simple implementation and will work only if the hyperlink code and name
 * each consist of one Run only.
 *
 * [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
 *
 * The field code contains a string in one of these formats:
 * HYPERLINK "url"
 * HYPERLINK \l "bookmark name"
 *
 * The field result contains text that is displayed to the user.
 */
public class Hyperlink
{
    public Hyperlink(FieldStart fieldStart) throws Exception
    {
        if (fieldStart == null)
            throw new IllegalArgumentException("fieldStart");
        if (fieldStart.getFieldType() != FieldType.FIELD_HYPERLINK)
            throw new IllegalArgumentException("Field start type must be FieldHyperlink.");

        mFieldStart = fieldStart;

        // Find field separator node.
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null)
            throw new Exception("Cannot find field separator.");

        // Find field end node. Normally field end will always be found, but in the example document
        // there happens to be a paragraph break included in the hyperlink and this puts the field end
        // in the next paragraph. It will be much more complicated to handle fields which span several
        // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);

        // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
        String fieldCode = getTextSameParent(mFieldStart.getNextSibling(), mFieldSeparator);

        Matcher match = G_REGEX.matcher(fieldCode.trim());

        if(match.find())
        {
            mIsLocal = match.group(1) != null;
            mTarget = match.group(2);
        }
    }

    /*
     * Gets or sets the display name of the hyperlink.
     */
    public String getName() throws Exception
    {
        return getTextSameParent(mFieldSeparator, mFieldEnd);
    }

    public void setName(String value) throws Exception
    {
        // Hyperlink display name is stored in the field result which is a Run
        // node between field separator and field end.
        Run fieldResult = (Run)mFieldSeparator.getNextSibling();
        fieldResult.setText(value);

        // But sometimes the field result can consist of more than one run, delete these runs.
        removeSameParent(fieldResult.getNextSibling(), mFieldEnd);
    }

    /*
     * Gets or sets the target url or bookmark name of the hyperlink.
     */
    public String getTarget() throws Exception
    {
        return mTarget;
    }

    public void setTarget(String value) throws Exception
    {
        mTarget = value;
        updateFieldCode();
    }

    /*
     * True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
     */
    public boolean isLocal() throws Exception
    {
        return mIsLocal;
    }

    public void setLocal(boolean value) throws Exception
    {
        mIsLocal = value;
        updateFieldCode();
    }

    /**
     * Updates the field code.
     */
    private void updateFieldCode() throws Exception
    {
        // Field code is stored in a Run node between field start and field separator.
        Run fieldCode = (Run)mFieldStart.getNextSibling();
        fieldCode.setText(java.text.MessageFormat.format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget));

        // But sometimes the field code can consist of more than one run, delete these runs.
        removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
    }

    /**
     * Goes through siblings starting from the start node until it finds a node of the specified type or null.
     */
    private static Node findNextSibling(Node start, int nodeType) throws Exception
    {
        for (Node node = start; node != null; node = node.getNextSibling())
        {
            if (node.getNodeType() == nodeType)
                return node;
        }
        return null;
    }

    /*
     * Retrieves text from start up to but not including the end node.
     */
    private static String getTextSameParent(Node start, Node end) throws Exception
    {
        if ((end != null) && (start.getParentNode() != end.getParentNode()))
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        StringBuilder builder = new StringBuilder();
        for (Node child = start; child != end; child = child.getNextSibling())
            builder.append(child.getText());
        return builder.toString();
    }

    /*
     * Removes nodes from start up to but not including the end node.
     * Start and end are assumed to have the same parent.
     */
    private static void removeSameParent(Node start, Node end) throws Exception
    {
        if ((end != null) && (start.getParentNode() != end.getParentNode()))
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        Node curChild = start;
        while (curChild != end)
        {
            Node nextChild = curChild.getNextSibling();
            curChild.remove();
            curChild = nextChild;
        }
    }

    private final Node mFieldStart;
    private final Node mFieldSeparator;
    private final Node mFieldEnd;
    private String mTarget;
    private boolean mIsLocal;

    private static final Pattern G_REGEX = Pattern.compile(
            "\\S+" +            // One or more non spaces HYPERLINK or other word in other languages
                    "\\s+" +            // One or more spaces
                    "(?:\"\"\\s+)?" +   // Non capturing optional "" and one or more spaces, found in one of the customers files.
                    "(\\\\l\\s+)?" +    // Optional \l flag followed by one or more spaces
                    "\"" +              // One apostrophe
                    "([^\"]+)" +        // One or more chars except apostrophe (hyperlink target)
                    "\""                // One closing apostrophe
    );
}