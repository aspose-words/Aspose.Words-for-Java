package DocsExamples.Mail_Merge_and_Reporting.Custom_examples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import java.util.ArrayList;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.ref.Ref;


class NestedMailMergeCustom extends DocsExamplesBase
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
    public static class Customer
    {
        public Customer(String aFullName, String anAddress)
        {
            setFullName(aFullName);
            setAddress(anAddress);
            setOrders(new ArrayList<Order>());
        }

        public String getFullName() { return mFullName; }; public void setFullName(String value) { mFullName = value; };

        private String mFullName;
        public String getAddress() { return mAddress; }; public void setAddress(String value) { mAddress = value; };

        private String mAddress;
        public ArrayList<Order> getOrders() { return mOrders; }; public void setOrders(ArrayList<Order> value) { mOrders = value; };

        private ArrayList<Order> mOrders;
    }

    /// <summary>
    /// An example of a child "data entity" class in your application.
    /// </summary>
    public static class Order
    {
        public Order(String oName, int oQuantity)
        {
            setName(oName);
            setQuantity(oQuantity);
        }

        public String getName() { return mName; }; public void setName(String value) { mName = value; };

        private String mName;
        public int getQuantity() { return mQuantity; }; public void setQuantity(int value) { mQuantity = value; };

        private int mQuantity;
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
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String TableName => "Customer";

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            switch (gStringSwitchMap.of(fieldName))
            {
                case /*"FullName"*/0:
                    fieldValue.set(mCustomers.get(mRecordIndex).getFullName());
                    return true;
                case /*"Address"*/1:
                    fieldValue.set(mCustomers.get(mRecordIndex).getAddress());
                    return true;
                case /*"Order"*/2:
                    fieldValue.set(mCustomers.get(mRecordIndex).getOrders());
                    return true;
                default:
                    fieldValue.set(null);
                    return false;
            }
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        public boolean moveNext()
        {
            if (!IsEof)
                mRecordIndex++;

            return !IsEof;
        }

        //ExStart:GetChildDataSourceExample           
        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            switch (gStringSwitchMap.of(tableName))
            {
                // Get the child collection to merge it with the region provided with tableName variable.
                case /*"Order"*/2:
                    return new OrderMailMergeDataSource(mCustomers.get(mRecordIndex).getOrders());
                default:
                    return null;
            }
        }
        //ExEnd:GetChildDataSourceExample

        private boolean IsEof => (mRecordIndex >= mCustomers.Count);

        private /*final*/ ArrayList<Customer> mCustomers;
        private int mRecordIndex;
    }

    public static class OrderMailMergeDataSource implements IMailMergeDataSource
    {
        public OrderMailMergeDataSource(ArrayList<Order> orders)
        {
            mOrders = orders;

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String TableName => "Order";

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            switch (gStringSwitchMap.of(fieldName))
            {
                case /*"Name"*/3:
                    fieldValue.set(mOrders[mRecordIndex].Name);
                    return true;
                case /*"Quantity"*/4:
                    fieldValue.set(mOrders[mRecordIndex].Quantity);
                    return true;
                default:
                    fieldValue.set(null);
                    return false;
            }
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        public boolean moveNext()
        {
            if (!IsEof)
                mRecordIndex++;

            return !IsEof;
        }

        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            // Return null because we haven't any child elements for this sort of object.
            return null;
        }

        private boolean IsEof => mRecordIndex >= private mOrders.CountmOrders;

        private /*final*/ ArrayList<Order> mOrders;
        private int mRecordIndex;
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"FullName",
		"Address",
		"Order",
		"Name",
		"Quantity"
	);

}
