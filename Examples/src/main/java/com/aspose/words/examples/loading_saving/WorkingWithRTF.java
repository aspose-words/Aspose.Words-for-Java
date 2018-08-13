package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.RtfLoadOptions;
import com.aspose.words.examples.Utils;

public class WorkingWithRTF {
    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(WorkingWithRTF.class);

        recognizeUtf8Text(dataDir);
    }

    public static void recognizeUtf8Text(String dataDir) throws Exception {
        //ExStart:RecognizeUtf8Text
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        loadOptions.setRecognizeUtf8Text(true);

        Document doc = new Document(dataDir + "Utf8Text.rtf", loadOptions);

        dataDir = dataDir + "RecognizeUtf8Text_out.rtf";
        doc.save(dataDir);
        //ExEnd:RecognizeUtf8Text
        System.out.println("\nUTF8 text has recognized successfully.\nFile saved at " + dataDir);
    }
}
