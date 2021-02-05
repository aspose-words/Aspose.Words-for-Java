package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Test //ExSkip
public class ExReplaceHyperlinks extends ApiExampleBase {
    @Test //ExSkip
    public void fields() throws Exception {
        Document doc = new Document(getMyDir() + "Hyperlinks.docx");

        // Hyperlinks in a Word documents are fields. To begin looking for hyperlinks, we must first find all the fields.
        // Use the "SelectNodes" method to find all the fields in the document via an XPath.
        NodeList fieldStarts = doc.selectNodes("//FieldStart");
        for (FieldStart fieldStart : (Iterable<FieldStart>) fieldStarts) {
            if (fieldStart.getFieldType() == FieldType.FIELD_HYPERLINK) {
                Hyperlink hyperlink = new Hyperlink(fieldStart);

                // Hyperlinks that link to bookmarks do not have URLs.
                if (hyperlink.isLocal()) continue;

                // Give each URL hyperlink a new URL and name.
                hyperlink.setTarget(NEW_URL);
                hyperlink.setName(NEW_NAME);
            }
        }

        doc.save(getArtifactsDir() + "ReplaceHyperlinks.Fields.docx");
    }

    private static final String NEW_URL = "http://www.aspose.com";
    private static final String NEW_NAME = "Aspose - The .NET & Java Component Publisher";
}


/**
 * This "facade" class makes it easier to work with a hyperlink field in a Word document.
 * <p>
 * HYPERLINK fields contain and display hyperlinks in the document body. A field in Aspose.Words
 * consists of several nodes, and it might be difficult to work with all those nodes directly.
 * This implementation will work only if the hyperlink code and name each consist of only one Run node.
 * <p>
 * The node structure for fields is as follows:
 * <p>
 * [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
 * <p>
 * Below are two example field codes of HYPERLINK fields:
 * HYPERLINK "url"
 * HYPERLINK \l "bookmark name"
 * <p>
 * A field's "Result" property contains text that the field displays in the document body to the user.
 */
class Hyperlink {
    Hyperlink(final FieldStart fieldStart) throws Exception {
        if (fieldStart == null) {
            throw new IllegalArgumentException("fieldStart");
        }

        if (fieldStart.getFieldType() != FieldType.FIELD_HYPERLINK) {
            throw new IllegalArgumentException("Field start type must be FieldHyperlink.");
        }

        mFieldStart = fieldStart;

        // Find the field separator node.
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null) {
            throw new IllegalStateException("Cannot find field separator.");
        }

        // Normally, we can always find the field's end node, but the example document 
        // contains a paragraph break inside a hyperlink, which puts the field end 
        // in the next paragraph. It will be much more complicated to handle fields which span several 
        // paragraphs correctly. In this case allowing field end to be null is enough.
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);

        // Field code looks something like "HYPERLINK "http:\\www.myurl.com"", but it can consist of several runs.
        String fieldCode = getTextSameParent(mFieldStart.getNextSibling(), mFieldSeparator);
        Matcher matcher = G_REGEX.matcher(fieldCode.trim());
        matcher.find();

        // The hyperlink is local if \l is present in the field code.
        mIsLocal = (matcher.group(1) != null) && (matcher.group(1).length() > 0);
        mTarget = matcher.group(2);
    }

    /**
     * Gets or sets the display name of the hyperlink.
     */
    String getName() throws Exception {
        return getTextSameParent(mFieldSeparator, mFieldEnd);
    }

    void setName(final String value) throws Exception {
        // Hyperlink display name is stored in the field result, which is a Run 
        // node between field separator and field end.
        Run fieldResult = (Run) mFieldSeparator.getNextSibling();
        fieldResult.setText(value);

        // If the field result consists of more than one run, delete these runs.
        removeSameParent(fieldResult.getNextSibling(), mFieldEnd);
    }

    /**
     * Gets or sets the target URL or bookmark name of the hyperlink.
     */
    String getTarget() {
        return mTarget;
    }

    void setTarget(final String value) throws Exception {
        mTarget = value;
        updateFieldCode();
    }

    /**
     * True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
     */
    boolean isLocal() {
        return mIsLocal;
    }

    void isLocal(final boolean value) throws Exception {
        mIsLocal = value;
        updateFieldCode();
    }

    private void updateFieldCode() throws Exception {
        // A field's field code is in a Run node between the field's start node and field separator.
        Run fieldCode = (Run) mFieldStart.getNextSibling();
        fieldCode.setText(java.text.MessageFormat.format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget));

        // If the field code consists of more than one run, delete these runs.
        removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
    }

    /**
     * Goes through siblings starting from the start node until it finds a node of the specified type or null.
     */
    private static Node findNextSibling(final Node startNode, final int nodeType) {
        for (Node node = startNode; node != null; node = node.getNextSibling()) {
            if (node.getNodeType() == nodeType) return node;
        }
        return null;
    }

    /**
     * Retrieves text from start up to but not including the end node.
     */
    private static String getTextSameParent(final Node startNode, final Node endNode) {
        if ((endNode != null) && (startNode.getParentNode() != endNode.getParentNode())) {
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");
        }

        StringBuilder builder = new StringBuilder();
        for (Node child = startNode; !child.equals(endNode); child = child.getNextSibling()) {
            builder.append(child.getText());
        }

        return builder.toString();
    }

    /**
     * Removes nodes from start up to but not including the end node.
     * Assumes that the start and end nodes have the same parent.
     */
    private static void removeSameParent(final Node startNode, final Node endNode) {
        if ((endNode != null) && (startNode.getParentNode() != endNode.getParentNode())) {
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");
        }

        Node curChild = startNode;
        while ((curChild != null) && (curChild != endNode)) {
            Node nextChild = curChild.getNextSibling();
            curChild.remove();
            curChild = nextChild;
        }
    }

    private final Node mFieldStart;
    private final Node mFieldSeparator;
    private final Node mFieldEnd;
    private boolean mIsLocal;
    private String mTarget;

    private static final Pattern G_REGEX = Pattern.compile(
            "\\S+" +             // One or more non spaces HYPERLINK or other word in other languages.
                    "\\s+" +             // One or more spaces.
                    "(?:\"\"\\s+)?" +    // Non-capturing optional "" and one or more spaces.
                    "(\\\\l\\s+)?" +     // Optional \l flag followed by one or more spaces.
                    "\"" +               // One apostrophe.
                    "([^\"]+)" +         // One or more characters, excluding the apostrophe (hyperlink target).
                    "\""                 // One closing apostrophe.
    );
}
//ExEnd