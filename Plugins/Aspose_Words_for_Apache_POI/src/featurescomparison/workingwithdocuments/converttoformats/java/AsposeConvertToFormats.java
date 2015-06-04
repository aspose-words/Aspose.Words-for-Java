package featurescomparison.workingwithdocuments.converttoformats.java;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

public class AsposeConvertToFormats
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/converttoformats/data/";

		// Load the document from disk.
        Document doc = new Document(dataPath + "document.doc");
        
        doc.save(dataPath + "html/Aspose_DocToHTML_Out.html",SaveFormat.HTML); //Save the document in HTML format.
        doc.save(dataPath + "Aspose_DocToPDF_Out.pdf",SaveFormat.PDF); //Save the document in PDF format.
        doc.save(dataPath + "Aspose_DocToTxt_Out.txt",SaveFormat.TEXT); //Save the document in TXT format.
        doc.save(dataPath + "Aspose_DocToJPG_Out.jpg",SaveFormat.JPEG); //Save the document in JPEG format.
        
        System.out.println("Aspose - Doc file converted in specified formats");
	}
}
