package featurescomparison.workingwithheadersandfooters.workingwithheader.java;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ApacheHeaders
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithheadersandfooters/workingwithheader/data/";

		POIFSFileSystem fs = null;

		fs = new POIFSFileSystem(new FileInputStream(dataPath + "AsposeHeader.doc"));
		HWPFDocument doc = new HWPFDocument(fs);

		int pageNumber = 1;

		HeaderStories headerStore = new HeaderStories(doc);
		String header = headerStore.getHeader(pageNumber);

		System.out.println("Header Is: " + header);
	}
}
