//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
package Examples;

import org.testng.annotations.Test;
import com.aspose.words.Document;
import java.util.ArrayList;
import com.aspose.words.IMailMergeDataSource;


public class ExMailMergeCustom extends ExBase
{
    @Test
    public void mailMergeCustomDataSourceCaller() throws Exception
    {
        mailMergeCustomDataSource();
    }

    //ExStart
    //ExFor:IMailMergeDataSource
    //ExFor:IMailMergeDataSource.TableName
    //ExFor:IMailMergeDataSource.MoveNext
    //ExFor:IMailMergeDataSource.GetChildDataSource
    //ExFor:IMailMergeDataSource.GetValue
    //ExFor:MailMerge.Execute(IMailMergeDataSource)
    //ExSummary:Performs mail merge from a custom data source.
    public void mailMergeCustomDataSource() throws Exception
    {
        // Create some data that we will use in the mail merge.
        CustomerList customers = new CustomerList();
        customers.add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        customers.add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        // Open the template document.
        Document doc = new Document(getMyDir() + "MailMerge.CustomDataSource.doc");

        // To be able to mail merge from your own data source, it must be wrapped
        // into an object that implements the IMailMergeDataSource interface.
        CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

        // Now you can pass your data source into Aspose.Words.
        doc.getMailMerge().execute(customersDataSource);

        doc.save(getMyDir() + "MailMerge.CustomDataSource Out.doc");
    }

    /**
     * An example of a "data entity" class in your application.
     */
    public class Customer
    {
        public Customer(String aFullName, String anAddress) throws Exception
        {
            mFullName = aFullName;
            mAddress = anAddress;
        }

        public String getFullName() throws Exception { return mFullName; }
        public void setFullName(String value) throws Exception { mFullName = value; }

        public String getAddress() throws Exception { return mAddress; }
        public void setAddress(String value) throws Exception { mAddress = value; }

        private String mFullName;
        private String mAddress;
    }

    /**
     * An example of a typed collection that contains your "data" objects.
     */
    public class CustomerList extends ArrayList
    {
        public Customer get(int index) { return (Customer)super.get(index); }
        public void set(int index, Customer value) { super.set(index, value); }
    }

    /**
     * A custom mail merge data source that you implement to allow Aspose.Words
     * to mail merge data from your Customer objects into Microsoft Word documents.
     */
    public class CustomerMailMergeDataSource implements IMailMergeDataSource
    {
        public CustomerMailMergeDataSource(CustomerList customers) throws Exception
        {
            mCustomers = customers;

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex= -1;
        }

        /**
         * The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
         */
        public String getTableName() throws Exception { return "Customer"; }

        /**
         * Aspose.Words calls this method to get a value for every data field.
         */
        public boolean getValue(String fieldName, Object[] fieldValue) throws Exception
        {
            if (fieldName.equals("FullName"))
            {
                fieldValue[0] = mCustomers.get(mRecordIndex).getFullName();
                return true;
            }
            else if (fieldName.equals("Address"))
            {
                fieldValue[0] = mCustomers.get(mRecordIndex).getAddress();
                return true;
            }
            else
            {
                // A field with this name was not found,
                // return false to the Aspose.Words mail merge engine.
                fieldValue[0] = null;
                return false;
            }
        }

        /**
         * A standard implementation for moving to a next record in a collection.
         */
        public boolean moveNext() throws Exception
        {
            if (!isEof())
                mRecordIndex++;

            return (!isEof());
        }

        public IMailMergeDataSource getChildDataSource(String tableName) throws Exception
        {
            return null;
        }

        private boolean isEof() throws Exception { return (mRecordIndex >= mCustomers.size()); }

        private final CustomerList mCustomers;
        private int mRecordIndex;
    }
    //ExEnd
}

