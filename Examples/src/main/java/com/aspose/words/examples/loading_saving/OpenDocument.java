package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.FileFormatInfo;
import com.aspose.words.FileFormatUtil;
import com.aspose.words.LoadFormat;
import com.aspose.words.examples.Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class OpenDocument{
    public static void main(String[] args) throws Exception {

        // The path to the documents directory.
        String dataDir = Utils.getDataDir(OpenDocument.class);
        String filename = "Test.docx";

        Document doc = new Document(dataDir + filename);


    }

}
