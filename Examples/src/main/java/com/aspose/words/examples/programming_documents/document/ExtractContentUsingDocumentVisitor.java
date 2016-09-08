package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Body;
import com.aspose.words.ControlChar;
import com.aspose.words.Document;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.FieldEnd;
import com.aspose.words.FieldSeparator;
import com.aspose.words.FieldStart;
import com.aspose.words.HeaderFooter;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.VisitorAction;
import com.aspose.words.examples.Utils;

public class ExtractContentUsingDocumentVisitor {

	public static void main(String[] args) throws Exception {
		
		String dataDir = Utils.getSharedDataDir(ExtractContentUsingDocumentVisitor.class) + "ExtractedSelectedContentBetweenNodes/";
		
		// Open the document we want to convert.
	    Document doc = new Document(dataDir + "Visitor.ToText.doc");

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
}

/**
 * Simple implementation of saving a document in the plain text format.
 * Implemented as a Visitor.
 */
class MyDocToTxtWriter extends DocumentVisitor {
	
	private final StringBuilder mBuilder;
	private boolean mIsSkipText;
	
	public MyDocToTxtWriter() throws Exception {
		mIsSkipText = false;
		mBuilder = new StringBuilder();
	}

	/**
	 * Gets the plain text of the document that was accumulated by the visitor.
	 */
	public String getText() throws Exception {
		return mBuilder.toString();
	}

	/**
	 * Called when a Run node is encountered in the document.
	 */
	public int visitRun(Run run) throws Exception {
		appendText(run.getText());

		// Let the visitor continue visiting other nodes.
		return VisitorAction.CONTINUE;
	}

	/**
	 * Called when a FieldStart node is encountered in the document.
	 */
	public int visitFieldStart(FieldStart fieldStart) throws Exception {
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
	public int visitFieldSeparator(FieldSeparator fieldSeparator) throws Exception {
		// Once reached a field separator node, we enable the output because we are
		// now entering the field result nodes.
		mIsSkipText = false;

		return VisitorAction.CONTINUE;
	}

	/**
	 * Called when a FieldEnd node is encountered in the document.
	 */
	public int visitFieldEnd(FieldEnd fieldEnd) throws Exception {
		// Make sure we enable the output when reached a field end because some fields
		// do not have field separator and do not have field result.
		mIsSkipText = false;

		return VisitorAction.CONTINUE;
	}

	/**
	 * Called when visiting of a Paragraph node is ended in the document.
	 */
	public int visitParagraphEnd(Paragraph paragraph) throws Exception {
		// When outputting to plain text we output Cr+Lf characters.
		appendText(ControlChar.CR_LF);

		return VisitorAction.CONTINUE;
	}

	public int visitBodyStart(Body body) throws Exception {
		// We can detect beginning and end of all composite nodes such as Section, Body,
		// Table, Paragraph etc and provide custom handling for them.
		mBuilder.append("*** Body Started ***\r\n");

		return VisitorAction.CONTINUE;
	}

	public int visitBodyEnd(Body body) throws Exception {
		mBuilder.append("*** Body Ended ***\r\n");
		return VisitorAction.CONTINUE;
	}

	/**
	 * Called when a HeaderFooter node is encountered in the document.
	 */
	public int visitHeaderFooterStart(HeaderFooter headerFooter) throws Exception {
		// Returning this value from a visitor method causes visiting of this
		// node to stop and move on to visiting the next sibling node.
		// The net effect in this example is that the text of headers and footers
		// is not included in the resulting output.
		return VisitorAction.SKIP_THIS_NODE;
	}

	/**
	 * Adds text to the current output. Honours the enabled/disabled output
	 * flag.
	 */
	private void appendText(String text) throws Exception {
		if (!mIsSkipText)
			mBuilder.append(text);
	}
}
