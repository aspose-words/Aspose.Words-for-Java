package com.aspose.words.examples.programming_documents.find_replace;

import java.awt.Color;
import java.util.Calendar;
import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.HeaderFooter;
import com.aspose.words.HeaderFooterCollection;
import com.aspose.words.HeaderFooterType;
import com.aspose.words.IReplacingCallback;
import com.aspose.words.Node;
import com.aspose.words.ParagraphAlignment;
import com.aspose.words.ReplaceAction;
import com.aspose.words.ReplacingArgs;
import com.aspose.words.examples.Utils;

public class FindAndReplace {

	public static final String dataDir = Utils.getSharedDataDir(FindAndReplace.class) + "FindAndReplace/";
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		SimpleStringReplacement();
		UsingRegularExpression();
		ReplaceTextContaingMetaCharacters();
		ReplaceTextInHeader();
		ReplaceTextInFooter();
		CustomizeFindAndReplaceOperation();
		ReplaceWithHtml();
		TestLineCounter();
	}

	public static void SimpleStringReplacement() throws Exception {
		//ExStart: SimpleStringReplacement
		// Load a Word DOCX document by creating an instance of the Document class.
		Document doc = new Document(); 
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.writeln("Hello _CustomerName_,");

		// Specify the search string and replace string using the Replace method.
		doc.getRange().replace("_CustomerName_", "James Bond", new FindReplaceOptions());

		// Save the result.
		doc.save(dataDir + "Range.ReplaceSimple.docx");
		//ExEnd: SimpleStringReplacement
	}
	
	public static void UsingRegularExpression() throws Exception {
		//ExStart: UsingRegularExpression
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.writeln("sad mad bad");
		
		if(doc.getText().trim() == "sad mad bad")
		{
			System.out.println("Strings are equal!");
		}

		// Replaces all occurrences of the words "sad" or "mad" to "bad".
		FindReplaceOptions options = new FindReplaceOptions();
		doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad", options);

		// Save the Word document.
		doc.save(dataDir + "Range.ReplaceWithRegex.docx");
		//ExEnd: UsingRegularExpression
	}
	
	public static void ReplaceTextContaingMetaCharacters() throws Exception {
		//ExStart: ReplaceTextContaingMetaCharacters
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.getFont().setName("Arial");
		builder.writeln("First section");
		builder.writeln("  1st paragraph");
		builder.writeln("  2nd paragraph");
		builder.writeln("{insert-section}");
		builder.writeln("Second section");
		builder.writeln("  1st paragraph");

		FindReplaceOptions options = new FindReplaceOptions();
		options.getApplyParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

		// Double each paragraph break after word "section", add kind of underline and make it centered.
		int count = doc.getRange().replace("section&p", "section&p----------------------&p", options);

		// Insert section break instead of custom text tag.
		count = doc.getRange().replace("{insert-section}", "&b", options);
		doc.save(dataDir + "ReplaceTextContaingMetaCharacters_out.docx");
		//ExEnd: ReplaceTextContaingMetaCharacters
	}
	
	private static void ReplaceTextInHeader() throws Exception {
        // ExStart:ReplaceTextInHeader
		// Open the template document, containing obsolete copyright information in the footer.
        Document doc = new Document(dataDir + "HeaderFooter.ReplaceText.doc");

		// Access header of the Word document.
		HeaderFooterCollection headersFooters = doc.getFirstSection().getHeadersFooters();
		HeaderFooter header = headersFooters.get(HeaderFooterType.HEADER_PRIMARY);

		// Set options.
		FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

		// Replace text in the header of the Word document.
		header.getRange().replace("Aspose.Words", "Remove", options);

		// Save the Word document.
		doc.save(dataDir + "HeaderReplace.docx");
        // ExEnd:ReplaceTextInHeader
    }
	
	private static void ReplaceTextInFooter() throws Exception {
		// Open the template document, containing obsolete copyright information in the footer.
        Document doc = new Document(dataDir + "HeaderFooter.ReplaceText.doc");

        // Set options.
        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);
        
		// Access header of the Word document.
        // ExStart:ReplaceTextInFooter
		HeaderFooterCollection headersFooters = doc.getFirstSection().getHeadersFooters();
        HeaderFooter footer = headersFooters.get(HeaderFooterType.FOOTER_PRIMARY);

        // Replace text in the footer of the Word document.
        int year = Calendar.getInstance().get(Calendar.YEAR);
        footer.getRange().replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) " + year + " by Aspose Pty Ltd.", options);
        // ExEnd:ReplaceTextInFooter
        
		// Save the Word document.
		doc.save(dataDir + "FooterReplace.docx");
    }
	
	private static void CustomizeFindAndReplaceOperation() throws Exception {
		Document doc = new Document(dataDir + "HeaderFooter.ReplaceText.doc");
		// ExStart:CustomizeFindAndReplaceOperation
		// Highlight word "the" with yellow color.
		FindReplaceOptions options = new FindReplaceOptions();
		options.getApplyFont().setHighlightColor(Color.YELLOW);

		// Replace highlighted text.
		doc.getRange().replace("the", "the", options);
		// ExEnd:CustomizeFindAndReplaceOperation
	}
	
	//ExStart: ReplaceWithHtml
	public static void ReplaceWithHtml() throws Exception { 
		Document doc = new Document();
		DocumentBuilder builder = new  DocumentBuilder(doc);
		builder.writeln("Hello <CustomerName>,"); 
		FindReplaceOptions options = new FindReplaceOptions();
		options.setReplacingCallback(new ReplaceWithHtmlEvaluator());
		
		doc.getRange().replace(Pattern.compile(" <CustomerName>,"), "", options);
		//doc.getRange().replace(" <CustomerName>,", html, options);
		
		// Save the modified document. 
		doc.save(dataDir + "Range.ReplaceWithInsertHtml.doc"); 
		System.out.println("\nText replaced with meta characters successfully.\nFile saved at " + dataDir);
	}

	static class ReplaceWithHtmlEvaluator implements IReplacingCallback {
        public int replacing(ReplacingArgs e) throws Exception {
        
        	// This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.getMatchNode();
            // create Document Buidler and insert MergeField
            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo(currentNode);
            // Replace '<CustomerName>' text with a red bold name.
            builder.insertHtml("<b><font color='red'>James Bond, </font></b>");e.getReplacement();
            currentNode.remove();
            //Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.SKIP;
        }
	}
	//ExEnd: ReplaceWithHtml
	
	//ExStart: NumberHighlightCallback
	// Replace and Highlight Numbers.
	static class NumberHighlightCallback implements IReplacingCallback {
		public int replacing (ReplacingArgs args) throws Exception {
			Node currentNode = args.getMatchNode();
			// Let replacement to be the same text.
			args.setReplacement(currentNode.getText());
			int val = currentNode.hashCode();
			
			// Apply either red or green color depending on the number value sign.
			FindReplaceOptions options = new FindReplaceOptions();
			if(val > 0)
			{
				options.getApplyFont().setColor(Color.GREEN);
			}
			else
			{
				options.getApplyFont().setColor(Color.RED);
			}
			
			return ReplaceAction.REPLACE;
		}
	}
	//ExEnd: NumberHighlightCallback
	
	//ExStart: TestLineCounter
	public static void TestLineCounter() throws Exception	{
		// Create a document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		// Add lines of text.
		builder.writeln("This is first line");
		builder.writeln("Second line");
		builder.writeln("And last line");
		
		// Prepend each line with line number.
		FindReplaceOptions opt = new FindReplaceOptions();
		opt.setReplacingCallback(new LineCounterCallback());
		doc.getRange().replace(Pattern.compile("[^&p]*&p"), "", opt);
		
		doc.save(dataDir + "TestLineCounter.docx");
	}

	static class LineCounterCallback implements IReplacingCallback
	{
		private int mCounter = 1;
		public int replacing(ReplacingArgs args) throws Exception {
			Node currentNode = args.getMatchNode();
			System.out.println(currentNode.getText());

			args.setReplacement(mCounter++ +"."+ currentNode.getText());
			return ReplaceAction.REPLACE;
		}
	}
	//ExEnd: TestLineCounter
}
