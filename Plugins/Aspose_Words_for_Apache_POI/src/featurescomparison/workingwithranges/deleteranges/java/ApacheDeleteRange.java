package featurescomparison.workingwithranges.deleteranges.java;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;

public class ApacheDeleteRange
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithranges/deleteranges/data/";
		
		HWPFDocument doc = new HWPFDocument(new FileInputStream(
				dataPath + "document.doc"));

		doc.getRange().getSection(0).delete();
		
		String text = doc.getRange().text();
		
		System.out.println("Range: " + text);
	}
}
