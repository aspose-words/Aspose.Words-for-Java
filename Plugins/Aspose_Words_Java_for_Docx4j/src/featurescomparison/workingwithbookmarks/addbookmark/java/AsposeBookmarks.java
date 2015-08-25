package featurescomparison.workingwithbookmarks.addbookmark.java;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;

public class AsposeBookmarks
{
	// See more @ http://www.aspose.com/docs/display/wordsjava/Bookmarks+in+Aspose.Words
	
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithbookmarks/addbookmark/data/";
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		
		builder.startBookmark("AsposeBookmark");
		builder.writeln("Text inside a bookmark.");
		builder.endBookmark("AsposeBookmark");
		
		// By index.
		Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);
		
		// By name.
		Bookmark bookmark2 = doc.getRange().getBookmarks().get("AsposeBookmark");
		
		doc.save(dataPath + "Aspose_Bookmark.doc", SaveFormat.DOC);
		System.out.println("Done.");
	}
}