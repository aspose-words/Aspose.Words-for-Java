//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

public class ExVisitor extends ApiExampleBase
{
    @Test
    public void toTextCaller() throws Exception
    {
        toText();
    }

    //ExStart
    //ExFor:Document.Accept
    //ExFor:Body.Accept
    //ExFor:DocumentVisitor
    //ExFor:DocumentVisitor.VisitAbsolutePositionTab
    //ExFor:DocumentVisitor.VisitBookmarkStart 
    //ExFor:DocumentVisitor.VisitBookmarkEnd
    //ExFor:DocumentVisitor.VisitRun
    //ExFor:DocumentVisitor.VisitFieldStart
    //ExFor:DocumentVisitor.VisitFieldEnd
    //ExFor:DocumentVisitor.VisitFieldSeparator
    //ExFor:DocumentVisitor.VisitBodyStart
    //ExFor:DocumentVisitor.VisitBodyEnd
    //ExFor:DocumentVisitor.VisitParagraphEnd
    //ExFor:DocumentVisitor.VisitHeaderFooterStart
    //ExFor:VisitorAction
    //ExId:ExtractContentDocToTxtConverter
    //ExSummary:Shows how to use the Visitor pattern to add new operations to the Aspose.Words object model. In this case we create a simple document converter into a text format.
    public void toText() throws Exception
    {
        // Open the document we want to convert.
        Document doc = new Document(getMyDir() + "DocumentVisitor.Destination.docx");

        // Create an object that inherits from the DocumentVisitor class.
        MyDocToTxtWriter myConverter = new MyDocToTxtWriter();

        // This is the well known Visitor pattern. Get the model to accept a visitor.
        // The model will iterate through itself by calling the corresponding methods
        // on the visitor object (this is called visiting).
        //
        // Note that every node in the object model has the Accept method so the visiting
        // can be executed not only for the whole document, but for any node in the document.
        doc.accept(myConverter);

        // Once the visiting is complete, we can retrieve the result of the operation,
        // that in this example, has accumulated in the visitor.
        System.out.println(myConverter.getText());
    }

    /**
     * Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
     */
    public class MyDocToTxtWriter extends DocumentVisitor
    {
        public MyDocToTxtWriter()
        {
            mIsSkipText = false;
            mBuilder = new StringBuilder();
        }

        /**
         * Gets the plain text of the document that was accumulated by the visitor.
         */
        public String getText()
        {
            return mBuilder.toString();
        }

        /**
         * Called when a Run node is encountered in the document.
         */
        public int visitRun(Run run) throws Exception
        {
            appendText(run.getText());

            // Let the visitor continue visiting other nodes.
            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldStart node is encountered in the document.
         */
        public int visitFieldStart(FieldStart fieldStart)
        {
            // In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
            // after a field start character. We want to skip field codes and output field
            // result only, therefore we use a flag to suspend the output while inside a field code.
            //
            // Note this is a very simplistic implementation and will not work very well
            // if you have nested fields in a document.
            mIsSkipText = true;

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldSeparator node is encountered in the document.
         */
        public int visitFieldSeparator(FieldSeparator fieldSeparator)
        {
            // Once reached a field separator node, we enable the output because we are
            // now entering the field result nodes.
            mIsSkipText = false;

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a FieldEnd node is encountered in the document.
         */
        public int visitFieldEnd(FieldEnd fieldEnd)
        {
            // Make sure we enable the output when reached a field end because some fields
            // do not have field separator and do not have field result.
            mIsSkipText = false;

            return VisitorAction.CONTINUE;
        }

        /**
         * Called when visiting of a Paragraph node is ended in the document.
         */
        public int visitParagraphEnd(Paragraph paragraph) throws Exception
        {
            // When outputting to plain text we output Cr+Lf characters.
            appendText(ControlChar.CR_LF);

            return VisitorAction.CONTINUE;
        }

        public int visitBodyStart(Body body)
        {
            // We can detect beginning and end of all composite nodes such as Section, Body,
            // Table, Paragraph etc and provide custom handling for them.
            mBuilder.append("*** Body Started ***\r\n");

            return VisitorAction.CONTINUE;
        }

        public int visitBodyEnd(Body body)
        {
            mBuilder.append("*** Body Ended ***\r\n");
            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a HeaderFooter node is encountered in the document.
         */
        public int visitHeaderFooterStart(HeaderFooter headerFooter)
        {
            // Returning this value from a visitor method causes visiting of this
            // node to stop and move on to visiting the next sibling node.
            // The net effect in this example is that the text of headers and footers
            // is not included in the resulting output.
            return VisitorAction.SKIP_THIS_NODE;
        }

        /**
         * Called when an AbsolutePositionTab is encountered in the document.
         */
        public int visitAbsolutePositionTab(AbsolutePositionTab tab)
        {
            mBuilder.append("\t");
            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a BookmarkStart is encountered in the document.
         */
        public int visitBookmarkStart(BookmarkStart bookmarkStart)
        {
            mBuilder.append("[");
            return VisitorAction.CONTINUE;
        }

        /**
         * Called when a BookmarkEnd is encountered in the document.
         */
        public int visitBookmarkEnd(BookmarkEnd bookmarkEnd)
        {
            mBuilder.append("]");
            return VisitorAction.CONTINUE;
        }

        /**
         * Adds text to the current output. Honours the enabled/disabled output flag.
         */
        private void appendText(String text)
        {
            if (!mIsSkipText) mBuilder.append(text);
        }

        private final StringBuilder mBuilder;
        private boolean mIsSkipText;
    }
    //ExEnd
}

