package DocsExamples.Mail_Merge_And_Reporting;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.ref.Ref;
import org.testng.annotations.Test;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;

@Test
public class WorkingWithXmlData extends DocsExamplesBase {
    @Test
    public void xmlMailMerge() throws Exception {
        //ExStart:XmlMailMerge
        //GistId:c4826e8042bd86942cd060b5f3bec3d0
        DataSet customersDs = new DataSet();
        customersDs.readXml(getMyDir() + "Mail merge data - Customers.xml");

        Document doc = new Document(getMyDir() + "Mail merge destinations - Registration complete.docx");
        doc.getMailMerge().execute(customersDs.getTables().get("Customer"));

        doc.save(getArtifactsDir() + "WorkingWithXmlData.XmlMailMerge.docx");
        //ExEnd:XmlMailMerge
    }

    @Test
    public void nestedMailMerge() throws Exception {
        //ExStart:NestedMailMerge
        //GistId:6ad68bd56dfc60c2162398d02d2fc1a5
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
    public void mustacheSyntaxUsingDataSet() throws Exception {
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
    public void linqToXmlMailMerge() throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        org.w3c.dom.Document xmlDoc = builder.parse(getMyDir() + "Mail merge data - Purchase order.xml");

        XPath xpath = XPathFactory.newInstance().newXPath();

        //ExStart:LinqToXmlMailMergeOrderItems
        NodeList itemNodes = (NodeList) xpath.compile("//Item").evaluate(xmlDoc, XPathConstants.NODESET);
        ArrayList<OrderItem> orderItems = new ArrayList();

        for (int i = 0; i < itemNodes.getLength(); i++) {
            Element item = (Element) itemNodes.item(i);
            OrderItem orderItem = new OrderItem();

            orderItem.setPartNumber(item.getAttribute("PartNumber"));
            orderItem.setProductName(getElementText(item, "ProductName"));
            orderItem.setQuantity(getElementText(item, "Quantity"));
            orderItem.setUSPrice(getElementText(item, "USPrice"));
            orderItem.setComment(getElementText(item, "Comment"));
            orderItem.setShipDate(getElementText(item, "ShipDate"));

            orderItems.add(orderItem);
        }
        //ExEnd:LinqToXmlMailMergeOrderItems

        //ExStart:LinqToXmlQueryForDeliveryAddress
        NodeList addressNodes = (NodeList) xpath.compile("//Address[@Type='Shipping']").evaluate(xmlDoc, XPathConstants.NODESET);
        ArrayList<DeliveryAddress> deliveryAddresses = new ArrayList();

        for (int i = 0; i < addressNodes.getLength(); i++) {
            Element address = (Element) addressNodes.item(i);
            DeliveryAddress deliveryAddress = new DeliveryAddress();

            deliveryAddress.setName(getElementText(address, "Name"));
            deliveryAddress.setCountry(getElementText(address, "Country"));
            deliveryAddress.setZip(getElementText(address, "Zip"));
            deliveryAddress.setState(getElementText(address, "State"));
            deliveryAddress.setCity(getElementText(address, "City"));
            deliveryAddress.setStreet(getElementText(address, "Street"));

            deliveryAddresses.add(deliveryAddress);
        }
        //ExEnd:LinqToXmlQueryForDeliveryAddress

        MyMailMergeDataSource orderItemsDataSource = new MyMailMergeDataSource(orderItems, "Items");
        MyMailMergeDataSource deliveryDataSource = new MyMailMergeDataSource(deliveryAddresses);

        //ExStart:LinqToXmlMailMerge
        Document doc = new Document(getMyDir() + "Mail merge destinations - LINQ.docx");

        // Fill the document with data from our data sources using mail merge regions for populating the order items
        // table is required because it allows the region to be repeated in the document for each order item.
        doc.getMailMerge().executeWithRegions(orderItemsDataSource);

        doc.getMailMerge().execute(deliveryDataSource);

        doc.save(getArtifactsDir() + "WorkingWithXmlData.LinqToXmlMailMerge.docx");
        //ExEnd:LinqToXmlMailMerge
    }

    private String getElementText(Element parent, String tagName) {
        NodeList nodeList = parent.getElementsByTagName(tagName);
        if (nodeList.getLength() > 0) {
            Node node = nodeList.item(0);
            return node.getTextContent();
        }
        return "";
    }

    public static class OrderItem {
        private String partNumber;
        private String productName;
        private String quantity;
        private String usPrice;
        private String comment;
        private String shipDate;

        public String getPartNumber() {
            return partNumber;
        }

        public void setPartNumber(String partNumber) {
            this.partNumber = partNumber;
        }

        public String getProductName() {
            return productName;
        }

        public void setProductName(String productName) {
            this.productName = productName;
        }

        public String getQuantity() {
            return quantity;
        }

        public void setQuantity(String quantity) {
            this.quantity = quantity;
        }

        public String getUSPrice() {
            return usPrice;
        }

        public void setUSPrice(String usPrice) {
            this.usPrice = usPrice;
        }

        public String getComment() {
            return comment;
        }

        public void setComment(String comment) {
            this.comment = comment;
        }

        public String getShipDate() {
            return shipDate;
        }

        public void setShipDate(String shipDate) {
            this.shipDate = shipDate;
        }
    }

    public static class DeliveryAddress {
        private String name;
        private String country;
        private String zip;
        private String state;
        private String city;
        private String street;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getCountry() {
            return country;
        }

        public void setCountry(String country) {
            this.country = country;
        }

        public String getZip() {
            return zip;
        }

        public void setZip(String zip) {
            this.zip = zip;
        }

        public String getState() {
            return state;
        }

        public void setState(String state) {
            this.state = state;
        }

        public String getCity() {
            return city;
        }

        public void setCity(String city) {
            this.city = city;
        }

        public String getStreet() {
            return street;
        }

        public void setStreet(String street) {
            this.street = street;
        }
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
        private Iterator<?> iterator;
        private Object current;
        private String tableName;
        private ArrayList<?> data;

        /// <summary>
        /// Creates a new instance of a custom mail merge data source.
        /// </summary>
        /// <param name="data">Data returned from a LINQ query.</param>
        //ExStart:MyMailMergeDataSourceConstructor
        public MyMailMergeDataSource(Collection<?> data) {
            this.data = new ArrayList<>(data);
            this.iterator = this.data.iterator();
            this.tableName = "";
        }
        //ExEnd:MyMailMergeDataSourceConstructor

        /// <summary>
        /// Creates a new instance of a custom mail merge data source, for mail merge with regions.
        /// </summary>
        /// <param name="data">Data returned from a LINQ query.</param>
        /// <param name="tableName">The name of the data source is only used when you perform a mail merge with regions. 
        /// If you prefer to use the simple mail merge, then use the constructor with one parameter.</param>          
        //ExStart:MyMailMergeDataSourceConstructorWithDataTable
        public MyMailMergeDataSource(Collection<?> data, String tableName) {
            this.data = new ArrayList<>(data);
            this.iterator = this.data.iterator();
            this.tableName = tableName != null ? tableName : "";
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
        @Override
        public boolean getValue(String fieldName, Ref<Object> fieldValue) {
            if (current == null) {
                fieldValue.set(null);
                return false;
            }

            try {
                Class<?> currentClass = current.getClass();

                // Try to get property using getter method (Java bean convention)
                String getterName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                Method getter = null;

                try {
                    getter = currentClass.getMethod(getterName);
                } catch (NoSuchMethodException e) {
                    // Try with "is" prefix for boolean properties
                    getterName = "is" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                    try {
                        getter = currentClass.getMethod(getterName);
                    } catch (NoSuchMethodException e2) {
                        // Try direct field access
                        try {
                            Field field = currentClass.getDeclaredField(fieldName);
                            field.setAccessible(true);
                            fieldValue.set(field.get(current));
                            return true;
                        } catch (Exception e3) {
                            fieldValue.set(null);
                            return false;
                        }
                    }
                }

                if (getter != null) {
                    Object value = getter.invoke(current);
                    fieldValue.set(value);
                    return true;
                }

            } catch (Exception e) {
                fieldValue.set(null);
                return false;
            }

            fieldValue.set(null);
            return false;
        }
        //ExEnd:MyMailMergeDataSourceGetValue

        /// <summary>
        /// Moves to the next record in the collection.
        /// </summary>            
        //ExStart:MyMailMergeDataSourceMoveNext
        @Override
        public boolean moveNext() {
            if (iterator.hasNext()) {
                current = iterator.next();
                return true;
            }
            return false;
        }
        //ExEnd:MyMailMergeDataSourceMoveNext

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        //ExStart:MyMailMergeDataSourceTableName
        @Override
        public String getTableName() {
            return tableName;
        }
        //ExEnd:MyMailMergeDataSourceTableName

        @Override
        public IMailMergeDataSource getChildDataSource(String tableName) {
            return null;
        }
    }
}
