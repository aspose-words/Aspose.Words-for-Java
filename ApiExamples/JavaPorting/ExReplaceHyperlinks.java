// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Finds all hyperlinks in a Word document and changes their URL and display name.
package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.NodeList;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.NodeType;
import com.aspose.ms.System.Text.RegularExpressions.Match;
import com.aspose.ms.System.msString;
import com.aspose.words.Run;
import com.aspose.words.Node;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.ms.System.Text.RegularExpressions.Regex;


/// <summary>
/// Shows how to replace hyperlinks in a Word document.
/// </summary>
@Test //ExSkip
public class ExReplaceHyperlinks extends ApiExampleBase
{
    /// <summary>
    /// Finds all hyperlinks in a Word document and changes their URL and display name.
    /// </summary>
    @Test //ExSkip
    public void fields() throws Exception
    {
        // Specify your document name here
        Document doc = new Document(getMyDir() + "Hyperlinks.docx");

        // Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks
        NodeList fieldStarts = doc.selectNodes("//FieldStart");
        for (FieldStart fieldStart : fieldStarts.<FieldStart>OfType() !!Autoporter error: Undefined expression type )
        {
            if (((fieldStart.getFieldType()) == (FieldType.FIELD_HYPERLINK)))
            {
                // The field is a hyperlink field, use the "facade" class to help to deal with the field
                Hyperlink hyperlink = new Hyperlink(fieldStart);

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these
                if (hyperlink.isLocal())
                    continue;

                // The Hyperlink class allows to set the target URL and the display name 
                // of the link easily by setting the properties
                hyperlink.setTarget(NEW_URL);
                hyperlink.setName(NEW_NAME);
            }
        }

        doc.save(getArtifactsDir() + "ReplaceHyperlinks.Fields.docx");
    }

    private static final String NEW_URL = "http://www.aspose.com";
    private static final String NEW_NAME = "Aspose - The .NET & Java Component Publisher";
}

/// <summary>
/// This "facade" class makes it easier to work with a hyperlink field in a Word document. 
/// 
/// A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words 
/// consists of several nodes and it might be difficult to work with all those nodes directly. 
/// Note this is a simple implementation and will work only if the hyperlink code and name 
/// each consist of one Run only.
/// 
/// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
/// 
/// The field code contains a String in one of these formats:
/// HYPERLINK "url"
/// HYPERLINK \l "bookmark name"
/// 
/// The field result contains text that is displayed to the user.
/// </summary>
class Hyperlink
{
    Hyperlink(FieldStart fieldStart)
    {
        if (fieldStart == null)
            throw new NullPointerException("fieldStart");
        if (!((fieldStart.getFieldType()) == (FieldType.FIELD_HYPERLINK)))
            throw new IllegalArgumentException("Field start type must be FieldHyperlink.");

        mFieldStart = fieldStart;

        // Find the field separator node
        mFieldSeparator = findNextSibling(mFieldStart, NodeType.FIELD_SEPARATOR);
        if (mFieldSeparator == null)
            throw new IllegalStateException("Cannot find field separator.");

        // Find the field end node. Normally field end will always be found, but in the example document 
        // there happens to be a paragraph break included in the hyperlink and this puts the field end 
        // in the next paragraph. It will be much more complicated to handle fields which span several 
        // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes
        mFieldEnd = findNextSibling(mFieldSeparator, NodeType.FIELD_END);

        // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs
        String fieldCode = getTextSameParent(mFieldStart.getNextSibling(), mFieldSeparator);
        Match match = G_REGEX.match(msString.trim(fieldCode));
        mIsLocal = match.getGroups().get(1).getLength() > 0; //The link is local if \l is present in the field code
        mTarget = match.getGroups().get(2).getValue();
    }

    /// <summary>
    /// Gets or sets the display name of the hyperlink.
    /// </summary>
    String getName() { return getTextSameParent(mFieldSeparator, mFieldEnd); }
    void setName(String value)
    {
        // Hyperlink display name is stored in the field result which is a Run 
        // node between field separator and field end
        Run fieldResult = (Run) mFieldSeparator.getNextSibling();
        fieldResult.setText(value);

        // But sometimes the field result can consist of more than one run, delete these runs
        removeSameParent(fieldResult.getNextSibling(), mFieldEnd);
    }

    /// <summary>
    /// Gets or sets the target url or bookmark name of the hyperlink.
    /// </summary>
    String getTarget()
    {
        return mTarget;
    }
    void setTarget(String value)
    {
        mTarget = value;
        updateFieldCode();
    }

    /// <summary>
    /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a url.
    /// </summary>
    boolean isLocal() { return mIsLocal; }
    void isLocal(boolean value)
    {
        mIsLocal = value;
        updateFieldCode();
    }

    private void updateFieldCode()
    {
        // Field code is stored in a Run node between field start and field separator
        Run fieldCode = (Run) mFieldStart.getNextSibling();
        fieldCode.setText(msString.format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget));

        // But sometimes the field code can consist of more than one run, delete these runs
        removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
    }

    /// <summary>
    /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
    /// </summary>
    private static Node findNextSibling(Node startNode, /*NodeType*/int nodeType)
    {
        for (Node node = startNode; node != null; node = node.getNextSibling())
        {
            if (((node.getNodeType()) == (nodeType)))
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
    private boolean mIsLocal;
    private String mTarget;

    private static /*final*/ Regex G_REGEX = new Regex(
        "\\S+" + // one or more non spaces HYPERLINK or other word in other languages
        "\\s+" + // one or more spaces
        "(?:\"\"\\s+)?" + // non capturing optional "" and one or more spaces, found in one of the customers files.
        "(\\\\l\\s+)?" + // optional \l flag followed by one or more spaces
        "\"" + // one apostrophe	
        "([^\"]+)" + // one or more chars except apostrophe (hyperlink target)
        "\"" // one closing apostrophe
    );
}

//ExEnd