// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import org.testng.annotations.Test;
import com.aspose.ms.System.Collections.msArrayList;
import com.aspose.words.Document;
import java.util.ArrayList;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.ref.Ref;


@Test
public class ExNestedMailMergeCustom extends ApiExampleBase
{
    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSource)
    //ExSummary:Performs mail merge with regions from a custom data source.
    @Test //ExSkip
    public void mailMergeCustomDataSource() throws Exception
    {
        // Create some data that we will use in the mail merge.
        CustomerList customers = new CustomerList();
        msArrayList.add(customers, new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        msArrayList.add(customers, new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        // Create some data for nesting in the mail merge.
        msArrayList.add(customers.get(0).getOrders(), new Order("Rugby World Cup Cap", 2));
        msArrayList.add(customers.get(0).getOrders(), new Order("Rugby World Cup Ball", 1));
        msArrayList.add(customers.get(1).getOrders(), new Order("Rugby World Cup Guide", 1));

        // Open the template document.
        Document doc = new Document(getMyDir() + "NestedMailMerge.CustomDataSource.doc");

        // To be able to mail merge from your own data source, it must be wrapped
        // into an object that implements the IMailMergeDataSource interface.
        CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

        // Now you can pass your data source into Aspose.Words.
        doc.getMailMerge().executeWithRegions(customersDataSource);

        doc.save(getArtifactsDir() + "NestedMailMerge.CustomDataSource.doc");
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
            setOrders(new OrderList());
        }

        public String getFullName() { return mFullName; }; public void setFullName(String value) { mFullName = value; };

        private String mFullName;
        public String getAddress() { return mAddress; }; public void setAddress(String value) { mAddress = value; };

        private String mAddress;
        public OrderList getOrders() { return mOrders; }; public void setOrders(OrderList value) { mOrders = value; };

        private OrderList mOrders;
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
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    public static class OrderList extends ArrayList
    {
        public /*new*/ Order get(int index) { return (Order) super.get(index); }
        public /*new*/void set(int index, Order value) { super.set(index, value); }
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

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String getTableName() { return "Customer"; }

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
                    // A field with this name was not found, 
                    // return false to the Aspose.Words mail merge engine.
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

            return (!isEof());
        }

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

        private boolean isEof() { return (mRecordIndex >= mCustomers.size()); }

        private /*final*/ CustomerList mCustomers;
        private int mRecordIndex;
    }

    public static class OrderMailMergeDataSource implements IMailMergeDataSource
    {
        public OrderMailMergeDataSource(OrderList orders)
        {
            mOrders = orders;

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String getTableName() { return "Order"; }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            switch (gStringSwitchMap.of(fieldName))
            {
                case /*"Name"*/3:
                    fieldValue.set(mOrders.get(mRecordIndex).getName());
                    return true;
                case /*"Quantity"*/4:
                    fieldValue.set(mOrders.get(mRecordIndex).getQuantity());
                    return true;
                default:
                    // A field with this name was not found, 
                    // return false to the Aspose.Words mail merge engine.
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

            return (!isEof());
        }

        // Return null because we haven't any child elements for this sort of object.
        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            return null;
        }

        private boolean isEof() { return (mRecordIndex >= mOrders.size()); }

        private /*final*/ OrderList mOrders;
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

    //ExEnd
}
