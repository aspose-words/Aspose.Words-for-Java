package com.aspose.words.examples.mail_merge;

import com.aspose.words.Document;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.IMailMergeDataSourceRoot;
import com.aspose.words.examples.Utils;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathFactory;
import java.util.HashMap;


public class MustacheTemplateSyntax
{

    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(MustacheTemplateSyntax.class);

        // Use DocumentBuilder from the javax.xml.parsers package and Document class from the org.w3c.dom package to read
        // the XML data file and store it in memory.
        javax.xml.parsers.DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();

        // Parse the XML data.
        org.w3c.dom.Document xmlData = db.parse(dataDir + "Orders.xml");

        // Open a template document.
        Document doc = new Document(dataDir + "ExecuteTemplate.doc");

        doc.getMailMerge().setUseNonMergeFields(true);

        // Note that this class also works with a single repeatable region (and any nested regions).
        // To merge multiple regions at the same time from a single XML data source, use the XmlMailMergeDataSet class.
        // e.g doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));
        doc.getMailMerge().executeWithRegions(new XmlMailMergeDataSet(xmlData));

        // Save the output document.
        doc.save(dataDir + "Output.docx");

        System.out.println("Mail merge performed successfully.");
    }


}

class XmlMailMergeDataSet implements IMailMergeDataSourceRoot
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

class XmlMailMergeDataTable implements IMailMergeDataSource
{
    /**
     * Creates a new XmlMailMergeDataSource for the specified XML document and table name.
     *
     * @param xmlDoc The DOM object which contains the parsed XML data.
     * @param tableName The name of the element in the data source where the data of the region is extracted from.
     */
    public XmlMailMergeDataTable(org.w3c.dom.Document xmlDoc, String tableName) throws Exception
    {
        this(xmlDoc.getDocumentElement(), tableName);
    }

    /**
     * Private constructor that is also called by GetChildDataSource.
     */
    private XmlMailMergeDataTable(Node rootNode, String tableName) throws Exception
    {
        mTableName = tableName;

        // Get the first element on this level matching the table name.
        mCurrentNode = (Node)retrieveExpression("./" + tableName).evaluate(rootNode, XPathConstants.NODE);
    }

    /**
     * The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
     */
    public String getTableName()
    {
        return mTableName;
    }

    /**
     * Aspose.Words calls this method to get a value for every data field.
     */
    public boolean getValue(String fieldName, Object[] fieldValue) throws Exception
    {
        // Attempt to retrieve the child node matching the field name by using XPath.
        Node value = (Node)retrieveExpression(fieldName).evaluate(mCurrentNode, XPathConstants.NODE);
        // We also look for the field name in attributes of the element node.
        Element nodeAsElement = (Element)mCurrentNode;

        if (value != null)
        {
            // Field exists in the data source as a child node, pass the value and return true.
            // This merges the data into the document.
            fieldValue[0] = value.getTextContent();
            return true;
        }
        else if (nodeAsElement.hasAttribute(fieldName))
        {
            // Field exists in the data source as an attribute of the current node, pass the value and return true.
            // This merges the data into the document.
            fieldValue[0] = nodeAsElement.getAttribute(fieldName);
            return true;
        }
        else
        {
            // Field does not exist in the data source, return false.
            // No value will be merged for this field and it is left over in the document.
            return false;
        }
    }

    /**
     * Moves to the next record in a collection. This method is a little different then the regular implementation as
     * we are walking over an XML document stored in a DOM.
     */
    public boolean moveNext()
    {
        if (!isEof())
        {
            // Don't move to the next node if this the first record to be merged.
            if (!mIsFirstRecord)
            {
                // Find the next node which is an element and matches the table name represented by this class.
                // This skips any text nodes and any elements which belong to a different table.
                do
                {
                    mCurrentNode = mCurrentNode.getNextSibling();
                }
                while ((mCurrentNode != null) && !(mCurrentNode.getNodeName().equals(mTableName) &&  (mCurrentNode.getNodeType() == Node.ELEMENT_NODE)));
            }
            else
            {
                mIsFirstRecord = false;
            }
        }

        return (!isEof());
    }

    /**
     * If the data source contains nested data this method will be called to retrieve the data for
     * the child table. In the XML data source nested data this should look like this:
     *
     * <Tables>
     *    <ParentTable>
     *       <Name>ParentName</Name>
     *       <ChildTable>
     *          <Text>Content</Text>
     *       </ChildTable>
     *    </ParentTable>
     * </Tables>
     */
    public IMailMergeDataSource getChildDataSource(String tableName) throws Exception
    {
        return new XmlMailMergeDataTable(mCurrentNode, tableName);
    }

    private boolean isEof()
    {
        return (mCurrentNode == null);
    }

    /**
     * Returns a cached version of a compiled XPathExpression if available, otherwise creates a new expression.
     */
    private XPathExpression retrieveExpression(String path) throws Exception
    {
        XPathExpression expression;

        if(mExpressionSet.containsKey(path))
        {
            expression = (XPathExpression)mExpressionSet.get(path);
        }
        else
        {
            expression = mXPath.compile(path);
            mExpressionSet.put(path, expression);
        }
        return expression;
    }

    /**
     * Instance variables.
     */
    private Node mCurrentNode;
    private boolean mIsFirstRecord = true;
    private final String mTableName;
    private final HashMap mExpressionSet = new HashMap();
    private final XPath mXPath = XPathFactory.newInstance().newXPath();
}