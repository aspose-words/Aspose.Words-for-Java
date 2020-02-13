package com.aspose.words.examples.programming_documents.bookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class ShowHideBookmarks {

	public static void main(String[] args) throws Exception {
		//ExStart:ShowHideBookmarks_call
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ShowHideBookmarks.class);
        
		Document doc = new Document(dataDir + "Bookmark.doc");
		showHideBookmarkedContent(doc,"MyBookmark",false);
		doc.save(dataDir + "Updated_Document.docx");
		
		//ExEnd:ShowHideBookmarks_call
		System.out.println("\n Updated Dcoument saved successfully.\nFile saved at " + dataDir);
	}
	
	//ExStart:ShowHideBookmarks
	public static void showHideBookmarkedContent(Document doc, String bookmarkName, boolean showHide) throws Exception {
	    DocumentBuilder builder = new DocumentBuilder(doc);
	    Bookmark bm = doc.getRange().getBookmarks().get(bookmarkName);

	    builder.moveToDocumentEnd();
	    // {IF "{MERGEFIELD bookmark}" = "true" "" ""}
	    Field field = builder.insertField("IF \"", null);
	    builder.moveTo(field.getStart().getNextSibling().getNextSibling());
	    builder.insertField("MERGEFIELD " + bookmarkName + "", null);
	    builder.write("\" = \"true\" ");
	    builder.write("\"");
	    builder.write("\"");
	    builder.write(" \"\"");

	    Node currentNode = field.getStart();
	    boolean flag = true;
	    while (currentNode != null && flag) {
	        if (currentNode.getNodeType() == NodeType.RUN)
	            if (currentNode.toString(SaveFormat.TEXT).trim().equals("\""))
	                flag = false;

	        Node nextNode = currentNode.getNextSibling();

	        bm.getBookmarkStart().getParentNode().insertBefore(currentNode, bm.getBookmarkStart());
	        currentNode = nextNode;
	    }

	    Node endNode = bm.getBookmarkEnd();
	    flag = true;
	    while (currentNode != null && flag) {
	        if (currentNode.getNodeType() == NodeType.FIELD_END)
	            flag = false;

	        Node nextNode = currentNode.getNextSibling();

	        bm.getBookmarkEnd().getParentNode().insertAfter(currentNode, endNode);
	        endNode = currentNode;
	        currentNode = nextNode;
	    }

	    doc.getMailMerge().execute(new String[]{bookmarkName}, new Object[]{showHide});
	    
	    //In case, you do not want to use MailMerge then you may use the following lines of code.
	    //builder.moveToMergeField(bookmarkName);
	    //builder.write(showHide ? "true" : "false");
	}
	//ExEnd:ShowHideBookmarks
}
