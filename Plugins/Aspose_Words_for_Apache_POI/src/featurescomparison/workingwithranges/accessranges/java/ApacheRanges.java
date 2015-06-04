package featurescomparison.workingwithranges.accessranges.java;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

public class ApacheRanges
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithranges/accessranges/data/";
		
		HWPFDocument doc = new HWPFDocument(new FileInputStream(
				dataPath + "document.doc"));

		Range range = doc.getRange();
		String text = range.text();
		
		System.out.println("Range: " + text);
	}
}
