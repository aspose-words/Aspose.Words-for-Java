/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithstyles.extractcontentbasedonstyles.java;

import java.util.ArrayList;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;


/**
 * Shows how to find paragraphs and runs formatted with a specific style.
 */
@SuppressWarnings("unchecked")
public class ExtractContentBasedOnStyles
{
    public static void main(String[] args) throws Exception
    {
            // The path to the documents directory.
        String dataDir = "src/programmingwithdocuments/workingwithstyles/extractcontentbasedonstyles/data/";

        //ExStart
        //ExId:ExtractContentBasedOnStyles_Main
        //ExSummary:Run queries and display results.
        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        // Define style names as they are specified in the Word document.
        final String PARA_STYLE = "Heading 1";
        final String RUN_STYLE = "Intense Emphasis";

        // Collect paragraphs with defined styles.
        // Show the number of collected paragraphs and display the text of this paragraphs.
        ArrayList paragraphs = paragraphsByStyleName(doc, PARA_STYLE);
        System.out.println(java.text.MessageFormat.format("Paragraphs with \"{0}\" styles ({1}):", PARA_STYLE, paragraphs.size()));
        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs)
            System.out.print(paragraph.toString(SaveFormat.TEXT));

        // Collect runs with defined styles.
        // Show the number of collected runs and display the text of this runs.
        ArrayList runs = runsByStyleName(doc, RUN_STYLE);
        System.out.println(java.text.MessageFormat.format("\nRuns with \"{0}\" styles ({1}):", RUN_STYLE, runs.size()));
        for (Run run : (Iterable<Run>) runs)
            System.out.println(run.getRange().getText());
        //ExEnd
    }

    //ExStart
    //ExId:ExtractContentBasedOnStyles_Paragraphs
    //ExSummary:Find all paragraphs formatted with the specified style.
    public static ArrayList paragraphsByStyleName(Document doc, String styleName) throws Exception
    {
        // Create an array to collect paragraphs of the specified style.
        ArrayList paragraphsWithStyle = new ArrayList();
        // Get all paragraphs from the document.
        NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH, true);
        // Look through all paragraphs to find those with the specified style.
        for (Paragraph paragraph : (Iterable<Paragraph>) paragraphs)
        {
            if (paragraph.getParagraphFormat().getStyle().getName().equals(styleName))
                paragraphsWithStyle.add(paragraph);
        }
        return paragraphsWithStyle;
    }
    //ExEnd

    //ExStart
    //ExId:ExtractContentBasedOnStyles_Runs
    //ExSummary:Find all runs formatted with the specified style.
    public static ArrayList runsByStyleName(Document doc, String styleName) throws Exception
    {
        // Create an array to collect runs of the specified style.
        ArrayList runsWithStyle = new ArrayList();
        // Get all runs from the document.
        NodeCollection runs = doc.getChildNodes(NodeType.RUN, true);
        // Look through all runs to find those with the specified style.
        for (Run run : (Iterable<Run>) runs)
        {
            if (run.getFont().getStyle().getName().equals(styleName))
                runsWithStyle.add(run);
        }
        return runsWithStyle;
    }
    //ExEnd
}