// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.Collections.msArrayList;
import java.util.ArrayList;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.ref.Ref;


@Test
public class ExMailMergeCustomNested extends ApiExampleBase
{
    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSource)
    //ExSummary:Shows how to use mail merge regions to execute a nested mail merge.
    @Test //ExSkip
    public void customDataSource() throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        // Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        // Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        builder.insertField(" MERGEFIELD TableStart:Customers");

        // These MERGEFIELDs are inside the mail merge region of the "Customers" table.
        // When we execute the mail merge, this field will receive data from rows in a data source named "Customers".
        builder.write("Full name:\t");
        builder.insertField(" MERGEFIELD FullName ");
        builder.write("\nAddress:\t");
        builder.insertField(" MERGEFIELD Address ");
        builder.write("\nOrders:\n");

        // Create a second mail merge region inside the outer region for a data source named "Orders".
        // The "Orders" data entries have a many-to-one relationship with the "Customers" data source.
        builder.insertField(" MERGEFIELD TableStart:Orders");

        builder.write("\tItem name:\t");
        builder.insertField(" MERGEFIELD Name ");
        builder.write("\n\tQuantity:\t");
        builder.insertField(" MERGEFIELD Quantity ");
        builder.insertParagraph();

        builder.insertField(" MERGEFIELD TableEnd:Orders");
        builder.insertField(" MERGEFIELD TableEnd:Customers");

        // Create related data with names that match those of our mail merge regions.
        CustomerList customers = new CustomerList();
        msArrayList.add(customers, new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        msArrayList.add(customers, new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        customers.get(0).getOrders().add(new Order("Rugby World Cup Cap", 2));
        customers.get(0).getOrders().add(new Order("Rugby World Cup Ball", 1));
        customers.get(1).getOrders().add(new Order("Rugby World Cup Guide", 1));

        // To mail merge from your data source, we must wrap it into an object that implements the IMailMergeDataSource interface.
        CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);
        
        doc.getMailMerge().executeWithRegions(customersDataSource);

        doc.save(getArtifactsDir() + "NestedMailMergeCustom.CustomDataSource.docx");
        testCustomDataSource(customers, new Document(getArtifactsDir() + "NestedMailMergeCustom.CustomDataSource.docx")); //ExSkip
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
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    public static class CustomerList extends ArrayList
    {
        public /*new*/ Customer get(int index) { return (Customer) super.get(index); }
        public /*new*/void set(int index, Customer value) { super.set(index, value); }
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
        public CustomerMailMergeDataSource(CustomerList customers)
        {
            mCustomers = customers;

            // When we initialize the data source, its position must be before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String getTableName() { return "Customers"; }

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
                    // Return "false" to the Aspose.Words mail merge engine to signify
                    // that we could not find a field with this name.
                    fieldValue.set(null);
                    return false;
            }
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
            switch (gStringSwitchMap.of(tableName))
            {
                // Get the child data source, whose name matches the mail merge region that uses its columns.
                case /*"Orders"*/3:
                    return new OrderMailMergeDataSource(mCustomers.get(mRecordIndex).getOrders());
                default:
                    return null;
            }
        }

        private boolean isEof() { return (mRecordIndex >= mCustomers.size()); }

        private /*final*/ CustomerList mCustomers;
        private int mRecordIndex;
    }

    public static class OrderMailMergeDataSource implements IMailMergeDataSource
    {
        public OrderMailMergeDataSource(ArrayList<Order> orders)
        {
            mOrders = orders;

            // When we initialize the data source, its position must be before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String getTableName() { return "Orders"; }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            switch (gStringSwitchMap.of(fieldName))
            {
                case /*"Name"*/4:
                    fieldValue.set(mOrders.get(mRecordIndex).getName());
                    return true;
                case /*"Quantity"*/5:
                    fieldValue.set(mOrders.get(mRecordIndex).getQuantity());
                    return true;
                default:
                    // Return "false" to the Aspose.Words mail merge engine to signify
                    // that we could not find a field with this name.
                    fieldValue.set(null);
                    return false;
            }
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

        /// <summary>
        /// Return null because we do not have any child elements for this sort of object.
        /// </summary>
        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            return null;
        }

        private boolean isEof() { return (mRecordIndex >= mOrders.size()); }

        private /*final*/ ArrayList<Order> mOrders;
        private int mRecordIndex;
    }
    //ExEnd

    private void testCustomDataSource(CustomerList customers, Document doc)
    {
        ArrayList<String[]> mailMergeData = new ArrayList<String[]>();

        for (Customer customer : (Iterable<Customer>) customers)
        {
            for (Order order : customer.getOrders())
                mailMergeData.add(new String[]{ order.getName(), Integer.toString(order.getQuantity()) });
            mailMergeData.add(new String[] {customer.getFullName(), customer.getAddress()});
        }

        TestUtil.mailMergeMatchesArray(msArrayList.toArray(mailMergeData, new String[][0]), doc, false);
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"FullName",
		"Address",
		"Order",
		"Orders",
		"Name",
		"Quantity"
	);

}
