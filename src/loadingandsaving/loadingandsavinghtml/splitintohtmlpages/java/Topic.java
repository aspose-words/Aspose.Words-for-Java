/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.splitintohtmlpages.java;

/**
 * A simple class to hold a topic title and HTML file name together.
 */
class Topic
{
    Topic(String title, String fileName) throws Exception
    {
        mTitle = title;
        mFileName = fileName;
    }

    String getTitle() throws Exception { return mTitle; }

    String getFileName() throws Exception { return mFileName; }

    private final String mTitle;
    private final String mFileName;
}