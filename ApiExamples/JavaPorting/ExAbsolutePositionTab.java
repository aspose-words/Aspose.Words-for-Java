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
import org.testng.Assert;
import com.aspose.words.AbsolutePositionTab;
import com.aspose.words.NodeType;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.Run;
import com.aspose.ms.System.Text.msStringBuilder;


@Test
public class ExAbsolutePositionTab extends ApiExampleBase
{
    //ExStart
    //ExFor:AbsolutePositionTab
    //ExFor:AbsolutePositionTab.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitAbsolutePositionTab
    //ExSummary:Shows how to process absolute position tab characters with a document visitor.
    @Test //ExSkip
    public void documentToTxt() throws Exception
    {
        Document doc = new Document(getMyDir() + "Absolute position tab.docx");

        // Extract the text contents of our document by accepting this custom document visitor.
        DocTextExtractor myDocTextExtractor = new DocTextExtractor();
        doc.getFirstSection().getBody().accept(myDocTextExtractor);

        // The absolute position tab, which has no equivalent in string form, has been explicitly converted to a tab character.
        Assert.assertEquals("Before AbsolutePositionTab\tAfter AbsolutePositionTab", myDocTextExtractor.getText());

        // An AbsolutePositionTab can accept a DocumentVisitor by itself too.
        AbsolutePositionTab absPositionTab = (AbsolutePositionTab)doc.getFirstSection().getBody().getFirstParagraph().getChild(NodeType.SPECIAL_CHAR, 0, true);

        myDocTextExtractor = new DocTextExtractor();
        absPositionTab.accept(myDocTextExtractor);

        Assert.assertEquals("\t", myDocTextExtractor.getText());
    }

    /// <summary>
    /// Collects the text contents of all runs in the visited document. Replaces all absolute tab characters with ordinary tabs.
    /// </summary>
    public static class DocTextExtractor extends DocumentVisitor
    {
        public DocTextExtractor()
        {
            mBuilder = new StringBuilder();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitRun(Run run)
        {
            appendText(run.getText());
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an AbsolutePositionTab node is encountered in the document.
        /// </summary>
        public /*override*/ /*VisitorAction*/int visitAbsolutePositionTab(AbsolutePositionTab tab)
        {
            msStringBuilder.append(mBuilder, "\t");
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        private void appendText(String text)
        {
            msStringBuilder.append(mBuilder, text);
        }

        /// <summary>
        /// Plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText()
        {
            return mBuilder.toString();
        }

        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd
}
