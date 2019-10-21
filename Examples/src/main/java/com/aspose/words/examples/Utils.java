package com.aspose.words.examples;

import java.io.File;

public class Utils {

    public static String getSharedDataDir(Class c) {
        File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "main");
        dir = new File(dir, "resources");

        return dir.toString() + File.separator;
    }

    public static String getDataDir(Class c) {
        File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "main");
        dir = new File(dir, "resources");

        for (String s : c.getName().split("\\.")) {
            dir = new File(dir, s);
            if (dir.isDirectory() == false)
                dir.mkdir();
        }
        System.out.println("Using data directory: " + dir.toString());
        return dir.toString() + File.separator;
    }

    public static void applyALicense() throws Exception {
        String dataDir = getSharedDataDir(Utils.class) + "License/";
        com.aspose.words.License license = new com.aspose.words.License();
        license.setLicense(dataDir + "Aspose.Words.Java.lic");
    }

    public static String GetOutputFilePath(String inputFilePath) {
        String extension = "";
        int i = inputFilePath.lastIndexOf('.');
        if (i > 0) {
            extension = inputFilePath.substring(i + 1);
        }
        if (inputFilePath.contains(".")) {
            inputFilePath = inputFilePath.substring(0, inputFilePath.lastIndexOf("."));
        }

        return inputFilePath + "_out_." + extension;
    }

}
