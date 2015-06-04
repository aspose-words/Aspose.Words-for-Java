package featurescomparison.workingwithdocuments.getdocumentproperties.java;

import java.io.FileInputStream;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;

public class ApacheDocumentProperties
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/getdocumentproperties/data/";
		
		HWPFDocument doc = new HWPFDocument(new FileInputStream(
				dataPath + "document.doc"));
		SummaryInformation summaryInfo = doc.getSummaryInformation();
		System.out.println(summaryInfo.getApplicationName());
		System.out.println(summaryInfo.getAuthor());
		System.out.println(summaryInfo.getComments());
		System.out.println(summaryInfo.getCharCount());
		System.out.println(summaryInfo.getEditTime());
		System.out.println(summaryInfo.getKeywords());
		System.out.println(summaryInfo.getLastAuthor());
		System.out.println(summaryInfo.getPageCount());
		System.out.println(summaryInfo.getRevNumber());
		System.out.println(summaryInfo.getSecurity());
		System.out.println(summaryInfo.getSubject());
		System.out.println(summaryInfo.getTemplate());
	}
}
