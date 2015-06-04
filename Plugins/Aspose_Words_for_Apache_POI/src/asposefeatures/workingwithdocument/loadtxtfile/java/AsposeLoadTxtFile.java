package asposefeatures.workingwithdocument.loadtxtfile.java;

import com.aspose.words.Document;

public class AsposeLoadTxtFile
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithdocument/loadtxtfile/data/";
		
        // The encoding of the text file is automatically detected.
        Document doc = new Document(dataPath + "LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        doc.save(dataPath + "AsposeLoadTxt_Out.docx");
        
		System.out.println("Process Completed Successfully");
	}
}
