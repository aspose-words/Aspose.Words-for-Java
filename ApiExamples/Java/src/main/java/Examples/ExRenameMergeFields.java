package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/// <summary>
/// Shows how to rename merge fields in a Word document.
/// </summary>
public class ExRenameMergeFields extends ApiExampleBase {
    /// <summary>
    /// Finds all merge fields in a Word document and changes their names.
    /// </summary>
    @Test
    public void rename() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Dear ");
        builder.insertField("MERGEFIELD  FirstName ");
        builder.write(" ");
        builder.insertField("MERGEFIELD  LastName ");
        builder.writeln(",");
        builder.insertField("MERGEFIELD  CustomGreeting ");

        // Select all field start nodes so we can find the MERGEFIELDs.
        NodeCollection fieldStarts = doc.getChildNodes(NodeType.FIELD_START, true);
        for (FieldStart fieldStart : (Iterable<FieldStart>) fieldStarts) {
            if (fieldStart.getFieldType() == FieldType.FIELD_MERGE_FIELD) {
                MergeField mergeField = new MergeField(fieldStart);
                mergeField.setName(mergeField.getName() + "_Renamed");
            }
        }

        doc.save(getArtifactsDir() + "RenameMergeFields.Rename.docx");
    }
}

/**
 * Represents a facade object for a merge field in a Microsoft Word document.
 */
class MergeField {
    MergeField(final FieldStart fieldStart) {
        if (fieldStart.equals(null)) {
            throw new IllegalArgumentException("fieldStart");
        }

        if (fieldStart.getFieldType() != FieldType.FIELD_MERGE_FIELD) {
            throw new IllegalArgumentException("Field start type must be FieldMergeField.");
        }

        mFieldStart = fieldStart;

        // Find the field separator node.
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null) {
            throw new IllegalStateException("Cannot find field separator.");
        }

        // Find the field end node. Normally field end will always be found, but in the example document 
        // there happens to be a paragraph break included in the hyperlink and this puts the field end 
        // in the next paragraph. It will be much more complicated to handle fields which span several 
        // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);
    }

    /**
     * Gets or sets the name of the merge field.
     */
    String getName() {
        String fieldResult = getTextSameParent(mFieldSeparator.getNextSibling(), mFieldEnd);
        int startPos = fieldResult.indexOf("«");
        startPos = (startPos >= 0) ? startPos + 1 : 0;

        int endPos = fieldResult.indexOf("»");
        endPos = (endPos >= 0) ? endPos : fieldResult.length();

        return fieldResult.substring(startPos, endPos);
    }

    void setName(final String value) {
        // Merge field name is stored in the field result which is a Run 
        // node between field separator and field end.
        Run fieldResult = (Run) mFieldSeparator.getNextSibling();
        fieldResult.setText(java.text.MessageFormat.format("«{0}»", value));

        // But sometimes the field result can consist of more than one run, delete these runs.
        removeSameParent(fieldResult.getNextSibling(), mFieldEnd);

        updateFieldCode(value);
    }

    private void updateFieldCode(final String fieldName) {
        // Field code is stored in a Run node between field start and field separator.
        Run fieldCode = (Run) mFieldStart.getNextSibling();
        Matcher matcher = G_REGEX.matcher(fieldCode.getText());

        matcher.find();

        String newFieldCode = java.text.MessageFormat.format(" {0}{1} ", matcher.group(1), fieldName);
        fieldCode.setText(newFieldCode);

        // But sometimes the field code can consist of more than one run, delete these runs.
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
     * Start and end are assumed to have the same parent.
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

    private static final Pattern G_REGEX = Pattern.compile("\\s*(MERGEFIELD\\s|)(\\s|)(\\S+)\\s+");
}