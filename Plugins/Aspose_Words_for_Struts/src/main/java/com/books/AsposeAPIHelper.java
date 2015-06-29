package com.books;

import java.util.List;
import java.util.Map;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;

import com.aspose.words.CellVerticalAlignment;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.HeightRule;
import com.aspose.words.ParagraphAlignment;

/**
 * 
 * @author Adeel
 *
 */
public class AsposeAPIHelper {
	/**
	 * Creates word document from list of book provided from grid. 
	 * 
	 * @param  out the current scope OutputStream.
	 * @param  books books list as map containing attributes.
	 * @param  context the App ServletContext
	 * @see    com.aspose.words.Document
	 */
	public static void createAsposeWordDoc(ServletOutputStream out,
			List<Map> books, ServletContext context) throws Exception {

		try {

			com.aspose.words.Document doc = new com.aspose.words.Document();

			// DocumentBuilder provides members to easily add content to a
			// document.
			DocumentBuilder builder = new DocumentBuilder(doc);

			Font font = builder.getFont();

			font.setSize(16);

			font.setColor(java.awt.Color.BLUE);

			font.setName("Arial");

			builder.insertParagraph();
			// Write a new paragraph in the document with the text

			builder.insertParagraph();
			builder.writeln("Books List");
			builder.insertParagraph();
			// Save the document in DOCX format. The format to save as is
			// inferred from the extension of the file name.
			// Aspose.Words supports saving any document in many more formats.

			builder.startTable();
			builder.insertCell();

			// Set height and define the height rule for the header row.
			builder.getRowFormat().setHeight(40.0);
			builder.getRowFormat().setHeightRule(HeightRule.AT_LEAST);

			// Some special features for the header row.
			builder.getCellFormat()
					.getShading()
					.setBackgroundPatternColor(
							new java.awt.Color(198, 217, 241));
			builder.getParagraphFormat()
					.setAlignment(ParagraphAlignment.CENTER);
			builder.getFont().setSize(16);
			builder.getFont().setName("Arial");
			builder.getFont().setBold(true);

			builder.getCellFormat().setWidth(100.0);
			builder.write("Book Id");
			builder.insertCell();
			builder.write("Book Name");
			builder.insertCell();
			builder.write("AuthorName");
			builder.insertCell();
			builder.write("Book Cost");
			builder.endRow();
			// Set features for the other rows and cells.
			builder.getCellFormat().getShading()
					.setBackgroundPatternColor(java.awt.Color.WHITE);
			builder.getCellFormat().setWidth(100.0);
			builder.getCellFormat().setVerticalAlignment(
					CellVerticalAlignment.CENTER);

			// Reset height and define a different height rule for table body
			builder.getRowFormat().setHeight(30.0);
			builder.getRowFormat().setHeightRule(HeightRule.AUTO);

			for (Map book : books) {
				String bookId = book.get("BookId").toString();
				String bookName = book.get("BookName").toString();
				String bookAuthorName = book.get("AuthorName").toString();
				String bookCost = book.get("BookCost").toString();
				builder.insertCell();
				// Reset font formatting.
				builder.getFont().setSize(12);
				builder.getFont().setBold(false);
				builder.write(bookId);
				builder.insertCell();
				builder.write(bookName);
				builder.insertCell();
				builder.write(bookAuthorName);
				builder.insertCell();
				builder.write(bookCost);
				builder.endRow();
			}
			builder.endTable();
			builder.insertParagraph();
			builder.insertParagraph();

			// Save the document

			doc.save(out, com.aspose.words.SaveFormat.DOC);

		} catch (Exception e) {
			throw new Exception(
					"Aspose: Unable to export to ms word format.. some error occured",
					e);

		}
	}
}
