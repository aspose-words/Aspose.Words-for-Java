package featurescomparison.workingwithranges.deleteranges.java;

import com.aspose.words.Document;

public class AsposeDeleteRange
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithranges/deleteranges/data/";
		
		Document doc = new Document(dataPath + "document.doc");
		doc.getSections().get(0).getRange().delete();
		
		String text = doc.getRange().getText();
		
		System.out.println("Range: " + text);
	}
}
