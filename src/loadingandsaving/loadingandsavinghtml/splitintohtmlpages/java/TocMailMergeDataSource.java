/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package loadingandsaving.loadingandsavinghtml.splitintohtmlpages.java;

import com.aspose.words.*;
import java.util.ArrayList;


/**
 * A custom data source for Aspose.Words mail merge.
 * Returns topic objects.
 */
class TocMailMergeDataSource implements IMailMergeDataSource
{
    TocMailMergeDataSource(ArrayList topics) throws Exception
    {
        mTopics = topics;
        // Initialize to BOF.
        mIndex = -1;
    }

    public boolean moveNext() throws Exception
    {
        if (mIndex < mTopics.size() - 1)
        {
            mIndex++;
            return true;
        }
        else
        {
            // Reached EOF, return false.
            return false;
        }
    }

    public boolean getValue(String fieldName, Object[] fieldValue) throws Exception
    {
        if ("TocEntry".equals(fieldName))
        {
            // The template document is supposed to have only one field called "TocEntry".
            fieldValue[0] = mTopics.get(mIndex);
            return true;
        }
        else
        {
            fieldValue[0] = null;
            return false;
        }
    }

    public String getTableName() throws Exception { return "TOC"; }

    public IMailMergeDataSource getChildDataSource(String tableName) throws Exception
    {
        return null;
    }

    private final ArrayList mTopics;
    private int mIndex;
}