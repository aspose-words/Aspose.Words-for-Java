package featurescomparison.workingwithranges.insertbeforeandafterranges.java;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;

public class ApacheInsert
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithranges/insertbeforeandafterranges/data/";
		
		HWPFDocument doc = new HWPFDocument(new FileInputStream(
				dataPath + "document.doc"));

		doc.getRange().getSection(0).insertBefore("Apache Inserted THIS Text before the below section");
		
		String text = doc.getRange().text();
		
		System.out.println("Range: " + text);
	}
}
