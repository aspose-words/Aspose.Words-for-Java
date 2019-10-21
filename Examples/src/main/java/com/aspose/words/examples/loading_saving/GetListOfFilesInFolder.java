package com.aspose.words.examples.loading_saving;

import com.aspose.words.examples.Utils;

import java.io.File;

public class GetListOfFilesInFolder {
    public static void main(String[] args) throws Exception {
        //ExStart:GetListOfFilesInFolder
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(GetListOfFilesInFolder.class);

        String supportedDir = dataDir + "OutSupported" + File.separator;
        String unknownDir = dataDir + "OutUnknown" + File.separator;
        String encryptedDir = dataDir + "OutEncrypted" + File.separator;
        String pre97Dir = dataDir + "OutPre97" + File.separator;

        File[] fileList = new File(dataDir).listFiles();

        // Loop through all found files.
        for (File file : fileList) {
            if (file.isDirectory())
                continue;

            // Extract and display the file name without the path.
            String nameOnly = file.getName();
            System.out.print(nameOnly);

        }
        //ExEnd:GetListOfFilesInFolder
    }
}
