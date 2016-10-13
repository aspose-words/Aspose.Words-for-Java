package com.aspose.words.examples.programming_documents.styles;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.ApplyFormatting.ApplyBordersAndShading;
import java.util.ArrayList;

/**
 * Shows how to find paragraphs and runs formatted with a specific style.
 */
public class ExtractContentBasedOnStyles {

	private static final String dataDir = Utils.getSharedDataDir(ApplyBordersAndShading.class) + "Styles/";

	public static void main(String[] args) throws Exception {
		// Open the document.
		Document doc = new Document(dataDir + "TestFile.doc");

		// Define style names as they are specified in the Word document.
		final String PARA_STYLE = "Heading 1";
		final String RUN_STYLE = "Intense Emphasis";

		// Collect paragraphs with defined styles.
		// Show the number of collected paragraphs and display the text of this paragraphs.
		ArrayList<Paragraph> paragraphs = paragraphsByStyleName(doc, PARA_STYLE);
		System.out.println(java.text.MessageFormat.format("Paragraphs with \"{0}\" styles ({1}):", PARA_STYLE, paragraphs.size()));
		for (Paragraph paragraph : paragraphs)
			System.out.print(paragraph.toString(SaveFormat.TEXT));

		// Collect runs with defined styles.
		// Show the number of collected runs and display the text of this runs.
		ArrayList<Run> runs = runsByStyleName(doc, RUN_STYLE);
		System.out.println(java.text.MessageFormat.format("\nRuns with \"{0}\" styles ({1}):", RUN_STYLE, runs.size()));
		for (Run run : runs)
			System.out.println(run.getRange().getText());
	}

	public static ArrayList<Paragraph> paragraphsByStyleName(Document doc, String styleName) throws Exception {
		// Create an array to collect paragraphs of the specified style.
		ArrayList<Paragraph> paragraphsWithStyle = new ArrayList();
		// Get all paragraphs from the document.
		NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);
		// Look through all paragraphs to find those with the specified style.
		for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs) {
			if (paragraph.getParagraphFormat().getStyle().getName().equals(styleName))
				paragraphsWithStyle.add(paragraph);
		}
		return paragraphsWithStyle;
	}

	public static ArrayList<Run> runsByStyleName(Document doc, String styleName) throws Exception {
		// Create an array to collect runs of the specified style.
		ArrayList<Run> runsWithStyle = new ArrayList();
		// Get all runs from the document.
		NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);
		// Look through all runs to find those with the specified style.
		for (Run run : (Iterable<Run>) runs) {
			if (run.getFont().getStyle().getName().equals(styleName))
				runsWithStyle.add(run);
		}
		return runsWithStyle;
	}
}