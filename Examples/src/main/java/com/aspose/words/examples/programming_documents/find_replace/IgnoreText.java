package com.aspose.words.examples.programming_documents.find_replace;

import java.util.Date;
import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FindReplaceOptions;

public class IgnoreText {

	public static void main(String[] args) throws Exception {
		
		IgnoreTextInsideFields();
        IgnoreTextInsideDeleteRevisions();
        IgnoreTextInsideInsertRevisions();
	}
	
	public static void IgnoreTextInsideFields() throws Exception {
		// ExStart:IgnoreTextInsideFields
        // Create document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert field with text inside.
        builder.insertField("INCLUDETEXT", "Text in field");

        Pattern regex = Pattern.compile("e", Pattern.CASE_INSENSITIVE);
        
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace 'e' in document ignoring text inside field.
        options.setIgnoreFields(true);
        doc.getRange().replace(regex, "*", options);
        System.out.println(doc.getText()); // The output is: \u0013INCLUDETEXT\u0014Text in field\u0015\f

        // Replace 'e' in document NOT ignoring text inside field.
        options.setIgnoreFields(false);
        doc.getRange().replace(regex, "*", options);
        System.out.println(doc.getText()); // The output is: \u0013INCLUDETEXT\u0014T*xt in fi*ld\u0015\f
        // ExEnd:IgnoreTextInsideFields
    }

	private static void IgnoreTextInsideDeleteRevisions() throws Exception
    {
        // ExStart:IgnoreTextInsideDeleteRevisions
        // Create new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert non-revised text.
        builder.writeln("Deleted");
        builder.write("Text");

        // Remove first paragraph with tracking revisions.
        doc.startTrackRevisions("author", new Date());
        doc.getFirstSection().getBody().getFirstParagraph().remove();
        doc.stopTrackRevisions();

        Pattern regex = Pattern.compile("e", Pattern.CASE_INSENSITIVE);
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace 'e' in document ignoring deleted text.
        options.setIgnoreDeleted(true);
        doc.getRange().replace(regex, "*", options);
        System.out.println(doc.getText()); // The output is: Deleted\rT*xt\f

        // Replace 'e' in document NOT ignoring deleted text.
        options.setIgnoreDeleted(false);
        doc.getRange().replace(regex, "*", options);
        System.out.println(doc.getText()); // The output is: D*l*t*d\rT*xt\f
        // ExEnd:IgnoreTextInsideDeleteRevisions
    }

	private static void IgnoreTextInsideInsertRevisions() throws Exception {
        // ExStart:IgnoreTextInsideInsertRevisions
        // Create new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert text with tracking revisions.
        doc.startTrackRevisions("author", new Date());
        builder.writeln("Inserted");
        doc.stopTrackRevisions();

        // Insert non-revised text.
        builder.write("Text");

        Pattern regex = Pattern.compile("e", Pattern.CASE_INSENSITIVE);
        FindReplaceOptions options = new FindReplaceOptions();

        // Replace 'e' in document ignoring inserted text.
        options.setIgnoreInserted(true);
        doc.getRange().replace(regex, "*", options);
        System.out.println(doc.getText()); // The output is: Inserted\rT*xt\f

        // Replace 'e' in document NOT ignoring inserted text.
        options.setIgnoreInserted(false);
        doc.getRange().replace(regex, "*", options);
        System.out.println(doc.getText()); // The output is: Ins*rt*d\rT*xt\f
        // ExEnd:IgnoreTextInsideInsertRevisions
    }
}