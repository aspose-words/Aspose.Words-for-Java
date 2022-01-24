package DocsExamples.Mail_Merge_and_Reporting;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.Document;
import com.aspose.words.IMailMergeDataSource;
import java.lang.Class;
import java.util.Iterator;
import com.aspose.words.ref.Ref;


class WorkingWithXmlData extends DocsExamplesBase
{
    @Test
    public void xmlMailMerge() throws Exception
    {
        //ExStart:XmlMailMerge
        DataSet customersDs = new DataSet();
        customersDs.readXml(getMyDir() + "Mail merge data - Customers.xml");

        Document doc = new Document(getMyDir() + "Mail merge destinations - Registration complete.docx");
        doc.getMailMerge().execute(customersDs.getTables().get("Customer"));

        doc.save(getArtifactsDir() + "WorkingWithXmlData.XmlMailMerge.docx");
        //ExEnd:XmlMailMerge
    }

    @Test
    public void nestedMailMerge() throws Exception
    {
        //ExStart:NestedMailMerge
        // The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
        DataSet pizzaDs = new DataSet();
        pizzaDs.readXml(getMyDir() + "Mail merge data - Orders.xml");
        
        Document doc = new Document(getMyDir() + "Mail merge destinations - Invoice.docx");

        // Trim trailing and leading whitespaces mail merge values.
        doc.getMailMerge().setTrimWhitespaces(false);

        doc.getMailMerge().executeWithRegions(pizzaDs);

        doc.save(getArtifactsDir() + "WorkingWithXmlData.NestedMailMerge.docx");
        //ExEnd:NestedMailMerge
    }

    @Test
    public void mustacheSyntaxUsingDataSet() throws Exception
    {
        //ExStart:MailMergeUsingMustacheSyntax
        DataSet ds = new DataSet();
        ds.readXml(getMyDir() + "Mail merge data - Vendors.xml");

        Document doc = new Document(getMyDir() + "Mail merge destinations - Vendor.docx");

        doc.getMailMerge().setUseNonMergeFields(true);

        doc.getMailMerge().executeWithRegions(ds);
        
        doc.save(getArtifactsDir() + "WorkingWithXmlData.MustacheSyntaxUsingDataSet.docx");
        //ExEnd:MailMergeUsingMustacheSyntax
    }

    @Test
    public void lINQtoXmlMailMerge() throws Exception
    {
        XElement orderXml = XElement.Load(getMyDir() + "Mail merge data - Purchase order.xml");

        // Query the purchase order XML file using LINQ to extract the order items into an object of an unknown type.
        //
        // Ensure you give the unknown type properties the same names as the MERGEFIELD fields in the document.
        //
        // To pass the actual values stored in the XML element or attribute to Aspose.Words,
        // we need to cast them to string. This prevents the XML tags from being inserted into the final document
        // when the XElement or XAttribute objects are passed to Aspose.Words.

        //ExStart:LINQtoXMLMailMergeorderItems
        var orderItems =
            from order : orderXml.Descendants("Item")
            select new
            {,
                PartNumber = (String) order.Attribute("PartNumber"),
                ProductName = (String) order.Element("ProductName"),
                Quantity = (String) order.Element("Quantity"),
                USPrice = (String) order.Element("USPrice"),
                Comment = (String) order.Element("Comment")
                ShipDate = (String) order.Element("ShipDate")
            };
        //ExEnd:LINQtoXMLMailMergeorderItems
        
        //ExStart:LINQToXMLQueryForDeliveryAddress
        var deliveryAddress =
            from delivery : orderXml.Elements("Address")
            where ((String) delivery.Attribute("Type") == "Shipping")
            select new
            {,
                Name = (String) delivery.Element("Name"),
                Country = (String) delivery.Element("Country"),
                Zip = (String) delivery.Element("Zip"),
                State = (String) delivery.Element("State"),
                City = (String) delivery.Element("City")
                Street = (String) delivery.Element("Street")
            };
        //ExEnd:LINQToXMLQueryForDeliveryAddress

        MyMailMergeDataSource orderItemsDataSource = new MyMailMergeDataSource(orderItems, "Items");
        MyMailMergeDataSource deliveryDataSource = new MyMailMergeDataSource(deliveryAddress);
        
        //ExStart:LINQToXMLMailMerge
        Document doc = new Document(getMyDir() + "Mail merge destinations - LINQ.docx");

        // Fill the document with data from our data sources using mail merge regions for populating the order items
        // table is required because it allows the region to be repeated in the document for each order item.
        doc.getMailMerge().executeWithRegions(orderItemsDataSource);

        doc.getMailMerge().execute(deliveryDataSource);

        doc.save(getArtifactsDir() + "WorkingWithXmlData.LINQtoXmlMailMerge.docx");
        //ExEnd:LINQToXMLMailMerge
    }

    /// <summary>
    /// Aspose.Words do not accept LINQ queries as input for mail merge directly
    /// but provide a generic mechanism that allows mail merges from any data source.
    /// 
    /// This class is a simple implementation of the Aspose.Words custom mail merge data source
    /// interface that accepts a LINQ query (any IEnumerable object).
    /// Aspose.Words call this class during the mail merge to retrieve the data.
    /// </summary>
    //ExStart:MyMailMergeDataSource 
    public static class MyMailMergeDataSource implements IMailMergeDataSource
    //ExEnd:MyMailMergeDataSource 
    {
        /// <summary>
        /// Creates a new instance of a custom mail merge data source.
        /// </summary>
        /// <param name="data">Data returned from a LINQ query.</param>
        //ExStart:MyMailMergeDataSourceConstructor 
        public MyMailMergeDataSource(Iterable data)
        {
            mEnumerator = data.iterator();
        }
        //ExEnd:MyMailMergeDataSourceConstructor

        /// <summary>
        /// Creates a new instance of a custom mail merge data source, for mail merge with regions.
        /// </summary>
        /// <param name="data">Data returned from a LINQ query.</param>
        /// <param name="tableName">The name of the data source is only used when you perform a mail merge with regions. 
        /// If you prefer to use the simple mail merge, then use the constructor with one parameter.</param>          
        //ExStart:MyMailMergeDataSourceConstructorWithDataTable
        public MyMailMergeDataSource(Iterable data, String tableName)
        {
            mEnumerator = data.iterator();
            mTableName = tableName;
        }
        //ExEnd:MyMailMergeDataSourceConstructorWithDataTable

        /// <summary>
        /// Aspose.Words call this method to get a value for every data field.
        /// 
        /// This is a simple "generic" implementation of a data source that can work over any IEnumerable collection.
        /// This implementation assumes that the merge field name in the document matches the public property's name
        /// on the object in the collection and uses reflection to get the property's value.
        /// </summary>
        //ExStart:MyMailMergeDataSourceGetValue
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            // Use reflection to get the property by name from the current object.
            Object obj = mEnumerator.next();
            Class currentRecordType = obj.getClass();

            PropertyInfo property = currentRecordType.GetProperty(fieldName);
            if (property != null)
            {
                fieldValue.set(property.GetValue(obj, null));
                return true;
            }
            fieldValue.set(null);

            return false;
        }
        //ExEnd:MyMailMergeDataSourceGetValue

        /// <summary>
        /// Moves to the next record in the collection.
        /// </summary>            
        //ExStart:MyMailMergeDataSourceMoveNext
        public boolean moveNext()
        {
            return mEnumerator.hasNext();
        }
        //ExEnd:MyMailMergeDataSourceMoveNext

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        //ExStart:MyMailMergeDataSourceTableName
        public String getTableName() { return mTableName; };

        private  String mTableName;
        //ExEnd:MyMailMergeDataSourceTableName

        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            return null;
        }

        private /*final*/ Iterator mEnumerator;
    }
}
