package featurescomparison.workingwithdocuments.openexistingdocument.java;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hwpf.HWPFDocument;

public class ApacheOpenExistingDoc
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/openexistingdocument/data/";

		HWPFDocument doc = new HWPFDocument(new FileInputStream(
				dataPath + "document.doc"));
		
		// write the file
        FileOutputStream out = new FileOutputStream(dataPath + "Apache_ExistingDoc_Out.doc");
        doc.write(out);
        out.close();
        
        System.out.println("Process Completed Successfully");
	}
}