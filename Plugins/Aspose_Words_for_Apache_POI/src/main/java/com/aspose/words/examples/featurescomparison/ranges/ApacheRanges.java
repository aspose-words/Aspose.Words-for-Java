package com.aspose.words.examples.featurescomparison.ranges;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import com.aspose.words.examples.Utils;

public class ApacheRanges
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheRanges.class);

        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                        dataDir + "document.doc"));

        Range range = doc.getRange();
        String text = range.text();

        System.out.println("Range: " + text);
    }
}
