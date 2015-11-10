package com.aspose.words.examples.featurescomparison.ranges;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;

import com.aspose.words.examples.Utils;

public class ApacheDeleteRange
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheDeleteRange.class);

        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                        dataDir + "document.doc"));

        doc.getRange().getSection(0).delete();

        String text = doc.getRange().text();

        System.out.println("Range: " + text);
    }
}
