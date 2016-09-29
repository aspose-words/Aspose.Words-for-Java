package com.aspose.words.examples.mail_merge;

import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.IMailMergeDataSourceRoot;

public class XmlMailMergeDataSet implements IMailMergeDataSourceRoot
{
    /**
     * Creates a new XmlMailMergeDataSet for the specified XML document. All regions in the document can be
     * merged at once using this class.
     *
     * @param xmlDoc The DOM object which contains the parsed XML data.
     */
    public XmlMailMergeDataSet(org.w3c.dom.Document xmlDoc)
    {
        mXmlDoc = xmlDoc;
    }

    public IMailMergeDataSource getDataSource(String tableName) throws Exception
    {
        return new XmlMailMergeDataTable(mXmlDoc, tableName);
    }

    private org.w3c.dom.Document mXmlDoc;
}