package com.aspose.words.examples.featurescomparison.headerfooter;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.aspose.words.examples.Utils;

public class ApacheFooters
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheFooters.class);

        POIFSFileSystem fs = null;

        fs = new POIFSFileSystem(new FileInputStream(dataDir + "AsposeFooter.doc"));
        HWPFDocument doc = new HWPFDocument(fs);

        int pageNumber = 1;

        HeaderStories headerStore = new HeaderStories(doc);
        String header = headerStore.getFooter(pageNumber);

        System.out.println("Footer Is: " + header);
    }
}
