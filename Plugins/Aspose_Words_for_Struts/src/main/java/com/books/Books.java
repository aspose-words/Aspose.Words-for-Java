package com.books;

import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Set;

class Books {

	int bookIdCount = 1000;
	Map<Integer, StoreBook> bookMap = new HashMap<Integer, StoreBook>();
	private static Books books = null;

	private Books() {
	}

	public static Books getInstance() {
		if (books == null) {
			books = new Books();
			books.storeBook("Mastering Java", "John Zakowsi", 200);
			books.storeBook(
					"Struts in Action",
					"Cedric Dumoulin, David Winterfeldt, George Franciscus, and Ted Husted",
					500);

		}
		return books;
	}

	public void storeBook(String bookName, String authorName, int bookCost) {
		StoreBook sb = new StoreBook();
		bookIdCount++;
		sb.addBook(bookIdCount, bookName, authorName, bookCost);
		bookMap.put(bookIdCount, sb);
	}

	public void updateBook(int bookId, String bookName, String authorName,
			int bookCost) {
		StoreBook sb = bookMap.get(bookId);
		sb.updateBook(bookName, authorName, bookCost);
	}

	public Map searchBook(int bookId) {
		return bookMap.get(bookId).getBooks();
	}

	public void deleteBook(int bookId) {
		bookMap.remove(bookId);
	}

	// Inner Class used to persist the app data ie) book details.
	class StoreBook {

		private String bookName;
		private String authorName;
		private int bookCost;
		private int bookId;

		StoreBook() {
		}

		public void addBook(int bookId, String bookName, String authorName,
				int bookCost) {
			this.bookId = bookId;
			this.bookName = bookName;
			this.authorName = authorName;
			this.bookCost = bookCost;
		}

		public void updateBook(String bookName, String authorName, int bookCost) {
			this.bookName = bookName;
			this.authorName = authorName;
			this.bookCost = bookCost;
		}

		public Map getBooks() {
			Map books = new HashMap();
			books.put("BookId", this.bookId);
			books.put("BookName", this.bookName);
			books.put("AuthorName", this.authorName);
			books.put("BookCost", this.bookCost);
			return books;
		}
	}

	public List getBookList() {
		List booksList = new ArrayList();
		Set s = bookMap.keySet();
		Iterator itr = s.iterator();
		while (itr.hasNext()) {
			booksList.add(bookMap.get((Integer) itr.next()).getBooks());
		}
		return booksList;
	}
}