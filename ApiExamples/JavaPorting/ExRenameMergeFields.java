// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.Run;
import com.aspose.ms.System.Text.RegularExpressions.Match;
import com.aspose.words.Node;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.ms.System.Text.RegularExpressions.Regex;


/// <summary>
/// Shows how to rename merge fields in a Word document.
/// </summary>
@Test
public class ExRenameMergeFields extends ApiExampleBase
{
    /// <summary>
    /// Finds all merge fields in a Word document and changes their names.
    /// </summary>
    @Test
    public void rename() throws Exception
    {
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
        for (FieldStart fieldStart : fieldStarts.<FieldStart>OfType() !!Autoporter error: Undefined expression type )
        {
            if (fieldStart.getFieldType() == FieldType.FIELD_MERGE_FIELD)
            {
                MergeField mergeField = new MergeField(fieldStart);
                mergeField.(mergeField.getName() + "_Renamed");
            }
        }

        doc.save(getArtifactsDir() + "RenameMergeFields.Rename.docx");
    }
}

/// <summary>
/// Represents a facade object for a merge field in a Microsoft Word document.
/// </summary>
class MergeField
{
    MergeField(FieldStart fieldStart)
    {
        if (fieldStart.getFieldType() != FieldType.FIELD_MERGE_FIELD)
            throw new IllegalArgumentException("Field start type must be FieldMergeField.");

        mFieldStart = fieldStart;

        // Find the field separator node.
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null)
            throw new IllegalStateException("Cannot find field separator.");

        // Find the field end node. Normally field end will always be found, but in the example document 
        // there happens to be a paragraph break included in the hyperlink and this puts the field end 
        // in the next paragraph. It will be much more complicated to handle fields which span several 
        // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);
    }

    /// <summary>
    /// Gets or sets the name of the merge field.
    /// </summary>
    String getName() { return mName; }

    private String mName; => GetTextSameParent(mFieldSeparator.NextSibling, mFieldEnd).Trim('«', '»');
        set
        {
            // Merge field name is stored in the field result which is a Run 
            // node between field separator and field end.
            Run fieldResult = (Run) mFieldSeparator.NextSibling;
            fieldResult.Text = $"«{value}»";

            // But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult.NextSibling, mFieldEnd);

            UpdateFieldCode(value);
        }
    }

    private void updateFieldCode(String fieldName)
    {
        // Field code is stored in a Run node between field start and field separator.
        Run fieldCode = (Run) mFieldStart.getNextSibling();
        Match match = G_REGEX.match(fieldCode.getText());

        String newFieldCode = $" {match.Groups["start"].Value}{fieldName} ";
        fieldCode.setText(newFieldCode);

        // But sometimes the field code can consist of more than one run, delete these runs.
        removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
    }

    /// <summary>
    /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
    /// </summary>
    private static Node findNextSibling(Node startNode, /*NodeType*/int nodeType)
    {
        for (Node node = startNode; node != null; node = node.getNextSibling())
        {
            if (node.getNodeType() == nodeType)
                return node;
        }

        return null;
    }

    /// <summary>
    /// Retrieves text from start up to but not including the end node.
    /// </summary>
    private static String getTextSameParent(Node startNode, Node endNode)
    {
        if (endNode != null && startNode.getParentNode() != endNode.getParentNode())
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        StringBuilder builder = new StringBuilder();
        for (Node child = startNode; !child.equals(endNode); child = child.getNextSibling())
            msStringBuilder.append(builder, child.getText());

        return builder.toString();
    }

    /// <summary>
    /// Removes nodes from start up to but not including the end node.
    /// Start and end are assumed to have the same parent.
    /// </summary>
    private static void removeSameParent(Node startNode, Node endNode)
    {
        if (endNode != null && startNode.getParentNode() != endNode.getParentNode())
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        Node curChild = startNode;
        while ((curChild != null) && (curChild != endNode))
        {
            Node nextChild = curChild.getNextSibling();
            curChild.remove();
            curChild = nextChild;
        }
    }

    private /*final*/ Node mFieldStart;
    private /*final*/ Node mFieldSeparator;
    private /*final*/ Node mFieldEnd;

    private static /*final*/ Regex G_REGEX = new Regex("\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+");
}
