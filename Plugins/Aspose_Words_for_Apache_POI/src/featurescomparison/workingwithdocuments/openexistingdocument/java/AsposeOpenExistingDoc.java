package featurescomparison.workingwithdocuments.openexistingdocument.java;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

public class AsposeOpenExistingDoc
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/openexistingdocument/data/";
		
		Document doc = new Document(dataPath + "document.doc");
		
		// Save the document in DOCX format.
		// Aspose.Words supports saving any document in many more formats.
		doc.save(dataPath + "Aspose_ExistingDoc_Out.docx",SaveFormat.DOCX);
		
        System.out.println("Process Completed Successfully");
	}
}
