package featurescomparison.workingwithranges.accessranges.java;

import com.aspose.words.Document;
import com.aspose.words.Range;

public class AsposeRanges
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithranges/accessranges/data/";

		Document doc = new Document(dataPath + "document.doc");
		Range range = doc.getRange();
		
		String text = range.getText();
		System.out.println("Range: " + text);
	}
}
