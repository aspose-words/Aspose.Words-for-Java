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
package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.NodeList;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.NodeType;
import com.aspose.ms.System.Text.RegularExpressions.Match;
import com.aspose.words.Run;
import java.text.MessageFormat;
import com.aspose.words.Node;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.ms.System.Text.RegularExpressions.Regex;


@Test //ExSkip
public class ExReplaceHyperlinks extends ApiExampleBase
{
    @Test //ExSkip
    public void fields() throws Exception
    {
        Document doc = new Document(getMyDir() + "Hyperlinks.docx");

        // Hyperlinks in a Word documents are fields. To begin looking for hyperlinks, we must first find all the fields.
        // Use the "SelectNodes" method to find all the fields in the document via an XPath.
        NodeList fieldStarts = doc.selectNodes("//FieldStart");

        for (FieldStart fieldStart : fieldStarts.<FieldStart>OfType() !!Autoporter error: Undefined expression type )
        {
            if (fieldStart.getFieldType() == FieldType.FIELD_HYPERLINK)
            {
                Hyperlink hyperlink = new Hyperlink(fieldStart);

                // Hyperlinks that link to bookmarks do not have URLs.
                if (hyperlink.isLocal())
                    continue;

                // Give each URL hyperlink a new URL and name.
                hyperlink.(NEW_URL);
                hyperlink.(NEW_NAME);
            }
        }

        doc.save(getArtifactsDir() + "ReplaceHyperlinks.Fields.docx");
    }

    private static final String NEW_URL = "http://www.aspose.com";
    private static final String NEW_NAME = "Aspose - The .NET & Java Component Publisher";
}

/// <summary>
/// HYPERLINK fields contain and display hyperlinks in the document body. A field in Aspose.Words 
/// consists of several nodes, and it might be difficult to work with all those nodes directly. 
/// This implementation will work only if the hyperlink code and name each consist of only one Run node.
///
/// The node structure for fields is as follows:
/// 
/// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
/// 
/// Below are two example field codes of HYPERLINK fields:
/// HYPERLINK "url"
/// HYPERLINK \l "bookmark name"
/// 
/// A field's "Result" property contains text that the field displays in the document body to the user.
/// </summary>
class Hyperlink
{
    Hyperlink(FieldStart fieldStart)
    {
        if (fieldStart == null)
            throw new NullPointerException("fieldStart");
        if (fieldStart.getFieldType() != FieldType.FIELD_HYPERLINK)
            throw new IllegalArgumentException("Field start type must be FieldHyperlink.");

        mFieldStart = fieldStart;

        // Find the field separator node.
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null)
            throw new IllegalStateException("Cannot find field separator.");

        // Normally, we can always find the field's end node, but the example document 
        // contains a paragraph break inside a hyperlink, which puts the field end 
        // in the next paragraph. It will be much more complicated to handle fields which span several 
        // paragraphs correctly. In this case allowing field end to be null is enough.
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);

        // Field code looks something like "HYPERLINK "http:\\www.myurl.com"", but it can consist of several runs.
        String fieldCode = getTextSameParent(mFieldStart.getNextSibling(), mFieldSeparator);
        Match match = G_REGEX.match(fieldCode.trim());

        // The hyperlink is local if \l is present in the field code.
        mIsLocal = match.getGroups().get(1).getLength() > 0; 
        mTarget = match.getGroups().get(2).getValue();
    }

    /// <summary>
    /// Gets or sets the display name of the hyperlink.
    /// </summary>
    String getName() { return mName; }

    private String mName; => GetTextSameParent(mFieldSeparator, mFieldEnd); 
        set
        {
            // Hyperlink display name is stored in the field result, which is a Run 
            // node between field separator and field end.
            Run fieldResult = (Run) mFieldSeparator.NextSibling;
            fieldResult.Text = value;

            // If the field result consists of more than one run, delete these runs.
            RemoveSameParent(fieldResult.NextSibling, mFieldEnd);
        }
    }

    /// <summary>
    /// Gets or sets the target URL or bookmark name of the hyperlink.
    /// </summary>
    String getTarget() { return mTarget; }

    private String mTarget; => mTarget;
        set
        {
            mTarget = value;
            UpdateFieldCode();
        }
    }

    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
    /// </summary>
    boolean isLocal() { return mIsLocal; }

    private boolean mIsLocal; => mIsLocal; 
        set
        {
            mIsLocal = value;
            UpdateFieldCode();
        }
    }

    private void updateFieldCode()
    {
        // A field's field code is in a Run node between the field's start node and field separator.
        Run fieldCode = (Run) mFieldStart.getNextSibling();
        fieldCode.setText(MessageFormat.format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget));

        // If the field code consists of more than one run, delete these runs.
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
        if ((endNode != null) && (startNode.getParentNode() != endNode.getParentNode()))
            throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

        StringBuilder builder = new StringBuilder();
        for (Node child = startNode; !child.equals(endNode); child = child.getNextSibling())
            msStringBuilder.append(builder, child.getText());

        return builder.toString();
    }

    /// <summary>
    /// Removes nodes from start up to but not including the end node.
    /// Assumes that the start and end nodes have the same parent.
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
    private boolean mIsLocal;
    private String mTarget;

    private static /*final*/ Regex G_REGEX = new Regex(
        "\\S+" + // One or more non spaces HYPERLINK or other word in other languages.
        "\\s+" + // One or more spaces.
        "(?:\"\"\\s+)?" + // Non-capturing optional "" and one or more spaces.
        "(\\\\l\\s+)?" + // Optional \l flag followed by one or more spaces.
        "\"" + // One apostrophe.	
        "([^\"]+)" + // One or more characters, excluding the apostrophe (hyperlink target).
        "\"" // One closing apostrophe.
    );
}

//ExEnd