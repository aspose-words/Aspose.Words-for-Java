package DocsExamples.Mail_Merge_And_Reporting.Complex_examples_and_helpers;

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import java.util.ArrayList;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.ref.Ref;

@Test
public class NestedMailMergeCustom extends DocsExamplesBase
{
    @Test
    public void customMailMerge() throws Exception
    {
        //ExStart:NestedMailMergeCustom
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField(" MERGEFIELD TableStart:Customer");

        builder.write("Full name:\t");
        builder.insertField(" MERGEFIELD FullName ");
        builder.write("\nAddress:\t");
        builder.insertField(" MERGEFIELD Address ");
        builder.write("\nOrders:\n");

        builder.insertField(" MERGEFIELD TableStart:Order");

        builder.write("\tItem name:\t");
        builder.insertField(" MERGEFIELD Name ");
        builder.write("\n\tQuantity:\t");
        builder.insertField(" MERGEFIELD Quantity ");
        builder.insertParagraph();

        builder.insertField(" MERGEFIELD TableEnd:Order");

        builder.insertField(" MERGEFIELD TableEnd:Customer");

        ArrayList<Customer> customers = new ArrayList<Customer>();
        {
            customers.add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
            customers.add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));
        }

        customers.get(0).getOrders().add(new Order("Rugby World Cup Cap", 2));
        customers.get(0).getOrders().add(new Order("Rugby World Cup Ball", 1));
        customers.get(1).getOrders().add(new Order("Rugby World Cup Guide", 1));

        // To be able to mail merge from your data source,
        // it must be wrapped into an object that implements the IMailMergeDataSource interface.
        CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

        doc.getMailMerge().executeWithRegions(customersDataSource);

        doc.save(getArtifactsDir() + "NestedMailMergeCustom.CustomMailMerge.docx");
        //ExEnd:NestedMailMergeCustom
    }

    /// <summary>
    /// An example of a "data entity" class in your application.
    /// </summary>
    public static class Customer {
        public Customer(String aFullName, String anAddress) {
            setFullName(aFullName);
            setAddress(anAddress);
            setOrders(new ArrayList<>());
        }

        public String getFullName() {
            return mFullName;
        }

        public void setFullName(String value) {
            mFullName = value;
        }

        public String getAddress() {
            return mAddress;
        }

        public void setAddress(String value) {
            mAddress = value;
        }

        public ArrayList<Order> getOrders() {
            return mOrders;
        }

        public void setOrders(ArrayList<Order> value) {
            mOrders = value;
        }

        private ArrayList<Order> mOrders;
        private String mAddress;
        private String mFullName;
    }

    /// <summary>
    /// An example of a child "data entity" class in your application.
    /// </summary>
    public static class Order {
        public Order(String oName, int oQuantity) {
            setName(oName);
            setQuantity(oQuantity);
        }

        public String getName() {
            return mName;
        }

        public void setName(String value) {
            mName = value;
        }

        public int getQuantity() {
            return mQuantity;
        }

        public void setQuantity(int value) {
            mQuantity = value;
        }

        private int mQuantity;
        private String mName;
    }

    /// <summary>
    /// A custom mail merge data source that you implement to allow Aspose.Words
    /// to mail merge data from your Customer objects into Microsoft Word documents.
    /// </summary>
    public static class CustomerMailMergeDataSource implements IMailMergeDataSource
    {
        public CustomerMailMergeDataSource(ArrayList<Customer> customers)
        {
            mCustomers = customers;

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, Ref<Object> fieldValue)
        {
            switch (fieldName)
            {
                case "FullName":
                    fieldValue.set(mCustomers.get(mRecordIndex).getFullName());
                    return true;
                case "Address":
                    fieldValue.set(mCustomers.get(mRecordIndex).getAddress());
                    return true;
                case "Order":
                    fieldValue.set(mCustomers.get(mRecordIndex).getOrders());
                    return true;
                default:
                    fieldValue.set(null);
                    return false;
            }
        }

        @Override
        public String getTableName() {
            return "Customer";
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        public boolean moveNext()
        {
            if (!IsEof())
                mRecordIndex++;

            return !IsEof();
        }

        //ExStart:GetChildDataSourceExample           
        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            switch (tableName)
            {
                // Get the child collection to merge it with the region provided with tableName variable.
                case "Order":
                    return new OrderMailMergeDataSource(mCustomers.get(mRecordIndex).getOrders());
                default:
                    return null;
            }
        }
        //ExEnd:GetChildDataSourceExample

        private boolean IsEof()
        {
            return mRecordIndex >= mCustomers.size();
        }

        private ArrayList<Customer> mCustomers;
        private int mRecordIndex;
    }

    public static class OrderMailMergeDataSource implements IMailMergeDataSource
    {
        public OrderMailMergeDataSource(ArrayList<NestedMailMergeCustom.Order> orders)
        {
            mOrders = orders;

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            switch (fieldName)
            {
                case "Name":
                    fieldValue.set(mOrders.get(mRecordIndex).getName());
                    return true;
                case "Quantity":
                    fieldValue.set(mOrders.get(mRecordIndex).getQuantity());
                    return true;
                default:
                    fieldValue.set(null);
                    return false;
            }
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        @Override
        public String getTableName() {
            return "Order";
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        public boolean moveNext()
        {
            if (!isEof())
                mRecordIndex++;

            return !isEof();
        }

        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            // Return null because we haven't any child elements for this sort of object.
            return null;
        }

        private boolean isEof()
        {
            return mRecordIndex >= mOrders.size() ;
        }

        private ArrayList<NestedMailMergeCustom.Order> mOrders;
        private int mRecordIndex;
    }
}
