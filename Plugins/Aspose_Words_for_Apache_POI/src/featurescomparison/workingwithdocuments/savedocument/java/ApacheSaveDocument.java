package featurescomparison.workingwithdocuments.savedocument.java;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ApacheSaveDocument
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/savedocument/data/";

		XWPFDocument document = new XWPFDocument();
		XWPFParagraph tmpParagraph = document.createParagraph();
		
		XWPFRun tmpRun = tmpParagraph.createRun();
		tmpRun.setText("Apache Sample Content for Word file.");
		
		FileOutputStream fos = new FileOutputStream(dataPath + "Apache_SaveDoc_Out.doc");
		document.write(fos);
		fos.close();
		
        System.out.println("Process Completed Successfully");
	}
}
