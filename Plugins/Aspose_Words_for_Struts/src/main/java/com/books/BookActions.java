package com.books;

import org.apache.struts.action.ActionForm;
import org.apache.struts.action.ActionMapping;
import org.apache.struts.action.ActionForward;
import org.apache.struts.actions.DispatchAction;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import java.util.List;
import java.util.Map;

/**
 * 
 * @author Adeel
 *
 */

public class BookActions extends DispatchAction {
	public ActionForward AddBook(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response) {
		System.out.println("Add Book Page");
		return mapping.findForward("addBook");
	}

	public ActionForward EditBook(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response) {
		System.out.println("Edit Book Page");
		int bookId = Integer.parseInt(request.getParameter("bookId"));

		Books b = Books.getInstance();
		Map bookDet = b.searchBook(bookId);

		// Used form bean class methods to fill the form input elements with
		// selected book values.
		BookForm bf = (BookForm) form;
		bf.setBookName(bookDet.get("BookName").toString());
		bf.setAuthorName(bookDet.get("AuthorName").toString());
		bf.setBookCost((Integer) bookDet.get("BookCost"));
		bf.setBookId((Integer) bookDet.get("BookId"));
		return mapping.findForward("editBook");
	}

	public ActionForward SaveBook(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response) {
		System.out.println("Save Book");
		// Used form bean class methods to get the value of form input elements.
		BookForm bf = (BookForm) form;
		String bookName = bf.getBookName();
		String authorName = bf.getAuthorName();
		int bookCost = bf.getBookCost();

		Books b = Books.getInstance();
		b.storeBook(bookName, authorName, bookCost);
		return new ActionForward("/showbooks.do", true);
	}

	public ActionForward UpdateBook(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response) {
		System.out.println("Update Book");
		BookForm bf = (BookForm) form;
		String bookName = bf.getBookName();
		String authorName = bf.getAuthorName();
		int bookCost = bf.getBookCost();
		int bookId = bf.getBookId();

		Books b = Books.getInstance();
		b.updateBook(bookId, bookName, authorName, bookCost);
		return new ActionForward("/showbooks.do", true);
	}

	public ActionForward DeleteBook(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response) {
		System.out.println("Delete Book");
		int bookId = Integer.parseInt(request.getParameter("bookId"));
		Books b = Books.getInstance();
		b.deleteBook(bookId);
		return new ActionForward("/showbooks.do", true);
	}

	/**
	 * Returns word file that can then be downloaded locally. 
	 * @see         AsposeAPIHelper
	 */
	public ActionForward ExportToWord(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response) {
		System.out.println("Aspose export document");

		Books b = Books.getInstance();

		List<Map> books = b.getBookList();
		response.setContentType("application/msword");
		response.setHeader("Content-Disposition",
				"attachment;filename=AsposeExportBooksList.doc");
		for (Map book : books) {
			try {
				AsposeAPIHelper.createAsposeWordDoc(response.getOutputStream(),
						books, request.getServletContext());
			} catch (Exception e) {
				e.printStackTrace();

			}

		}

		return null;
	}

}