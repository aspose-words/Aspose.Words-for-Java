package featurescomparison.workingwithdocuments.createnewdocument.java;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ApacheNewDocument
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/createnewdocument/data/";
		
		XWPFDocument document = new XWPFDocument();
		XWPFParagraph tmpParagraph = document.createParagraph();
		XWPFRun tmpRun = tmpParagraph.createRun();
		tmpRun.setText("Apache Sample Content for Word file.");
		tmpRun.setFontSize(18);
		
		FileOutputStream fos = new FileOutputStream(dataPath + "Apache_newWordDoc_Out.doc");
		document.write(fos);
		fos.close();
	}
}
