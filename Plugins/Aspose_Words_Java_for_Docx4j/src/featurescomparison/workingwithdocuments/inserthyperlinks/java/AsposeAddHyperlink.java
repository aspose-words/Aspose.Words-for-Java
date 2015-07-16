package featurescomparison.workingwithdocuments.inserthyperlinks.java;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;

public class AsposeAddHyperlink 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/inserthyperlinks/data/";
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		 
		builder.write("Please make sure to visit ");
		// Insert the link.
		builder.insertHyperlink("Aspose Website", "http://www.aspose.com", false);
		
		doc.save(dataPath + "AsposeAddHyperlinks.doc");
		System.out.println("Done.");
	}
}
