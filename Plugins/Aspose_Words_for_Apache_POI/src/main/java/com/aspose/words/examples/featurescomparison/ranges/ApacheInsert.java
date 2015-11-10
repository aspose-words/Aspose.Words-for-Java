package com.aspose.words.examples.featurescomparison.ranges;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;

import com.aspose.words.examples.Utils;

public class ApacheInsert
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ApacheInsert.class);

        HWPFDocument doc = new HWPFDocument(new FileInputStream(
                        dataDir + "document.doc"));

        doc.getRange().getSection(0).insertBefore("Apache Inserted THIS Text before the below section");

        String text = doc.getRange().text();

        System.out.println("Range: " + text);
    }
}
