package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by Home on 10/13/2017.
 */
public class WorkingWithFootnote {


    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getSharedDataDir(WorkingWithFootnote.class) + "Document/";
        SetFootNoteColumns(dataDir);
        SetFootnoteAndEndNotePosition(dataDir);
        SetEndnoteOptions(dataDir);
    }

    private static void SetFootNoteColumns(String dataDir) throws Exception {
        // ExStart:SetFootNoteColumns
        Document doc = new Document(dataDir + "TestFile.docx");

        //Specify the number of columns with which the footnotes area is formatted.
        doc.getFootnoteOptions().setColumns(3);
        dataDir = dataDir + "TestFile_Out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:SetFootNoteColumns
        System.out.println("\nFootnote number of columns set successfully.\nFile saved at " + dataDir);
    }

    private static void SetFootnoteAndEndNotePosition(String dataDir) throws Exception {
        // ExStart:SetFootnoteAndEndNotePosition
        Document doc = new Document(dataDir + "TestFile.docx");

        //Set footnote and endnode position.
        doc.getFootnoteOptions().setPosition(FootnotePosition.BENEATH_TEXT);
        doc.getEndnoteOptions().setPosition(EndnotePosition.END_OF_SECTION);
        dataDir = dataDir + "TestFile_Out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:SetFootnoteAndEndNotePosition
        System.out.println("\nFootnote number of columns set successfully.\nFile saved at " + dataDir);
    }

    private static void SetEndnoteOptions(String dataDir) throws Exception {
        // ExStart:SetEndnoteOptions
        Document doc = new Document(dataDir + "TestFile.docx");

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("Some text");

        builder.insertFootnote(FootnoteType.ENDNOTE, "Endnote text.");

        EndnoteOptions option = doc.getEndnoteOptions();
        option.setRestartRule(FootnoteNumberingRule.RESTART_PAGE);
        option.setPosition(EndnotePosition.END_OF_SECTION);

        dataDir = dataDir + "TestFile_Out.doc";

        // Save the document to disk.
        doc.save(dataDir);
        // ExEnd:SetEndnoteOptions
        System.out.println("\nEootnote is inserted at the end of section successfully.\nFile saved at " + dataDir);
    }
}
