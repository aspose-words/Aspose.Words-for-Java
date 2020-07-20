package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

@Test
public class ExAbsolutePositionTab extends ApiExampleBase {
    //ExStart
    //ExFor:AbsolutePositionTab
    //ExFor:AbsolutePositionTab.Accept(DocumentVisitor)
    //ExFor:DocumentVisitor.VisitAbsolutePositionTab
    //ExSummary:Shows how to work with absolute position tabs.
    @Test //ExSkip
    public void documentToTxt() throws Exception {
        // This document contains two sentences separated by an absolute position tab
        Document doc = new Document(getMyDir() + "Absolute position tab.docx");

        // An AbsolutePositionTab is a child node of a paragraph
        // AbsolutePositionTabs get picked up when looking for nodes of the SpecialChar type
        Paragraph para = doc.getFirstSection().getBody().getFirstParagraph();
        AbsolutePositionTab absPositionTab = (AbsolutePositionTab) para.getChild(NodeType.SPECIAL_CHAR, 0, true);

        // This implementation of the DocumentVisitor pattern converts the document to plain text
        DocToTxtWriter myDocToTxtWriter = new DocToTxtWriter();

        // We can run the DocumentVisitor over the whole first paragraph
        para.accept(myDocToTxtWriter);

        // A tab character is placed where the AbsolutePositionTab was found
        Assert.assertEquals(myDocToTxtWriter.getText(), "Before AbsolutePositionTab\tAfter AbsolutePositionTab");

        // An AbsolutePositionTab can accept a DocumentVisitor by itself too
        myDocToTxtWriter = new DocToTxtWriter();
        absPositionTab.accept(myDocToTxtWriter);

        Assert.assertEquals(myDocToTxtWriter.getText(), "\t");
    }

    /// <summary>
    /// Visitor implementation that simply collects the Runs and AbsolutePositionTabs of a document as plain text.
    /// </summary>
    public static class DocToTxtWriter extends DocumentVisitor {
        public DocToTxtWriter() {
            mBuilder = new StringBuilder();
        }

        /// <summary>
        /// Called when a Run node is encountered in the document.
        /// </summary>
        public int visitRun(final Run run) {
            appendText(run.getText());

            // Let the visitor continue visiting other nodes
            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Called when an AbsolutePositionTab node is encountered in the document.
        /// </summary>
        public int visitAbsolutePositionTab(final AbsolutePositionTab tab) {
            // We'll treat the AbsolutePositionTab as a regular tab in this case
            mBuilder.append("\t");

            return VisitorAction.CONTINUE;
        }

        /// <summary>
        /// Adds text to the current output. Honors the enabled/disabled output flag.
        /// </summary>
        public void appendText(final String text) {
            mBuilder.append(text);
        }

        /// <summary>
        /// Gets the plain text of the document that was accumulated by the visitor.
        /// </summary>
        public String getText() {
            return mBuilder.toString();
        }

        private StringBuilder mBuilder;
    }
    //ExEnd
}
