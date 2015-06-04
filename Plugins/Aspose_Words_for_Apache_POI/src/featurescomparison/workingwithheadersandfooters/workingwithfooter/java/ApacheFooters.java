package featurescomparison.workingwithheadersandfooters.workingwithfooter.java;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ApacheFooters
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithheadersandfooters/workingwithfooter/data/";

		POIFSFileSystem fs = null;

		fs = new POIFSFileSystem(new FileInputStream(dataPath + "AsposeFooter.doc"));
		HWPFDocument doc = new HWPFDocument(fs);

		int pageNumber = 1;

		HeaderStories headerStore = new HeaderStories(doc);
		String header = headerStore.getFooter(pageNumber);

		System.out.println("Header Is: " + header);
	}
}
