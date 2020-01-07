package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.ref.Ref;
import org.testng.annotations.Test;

import java.util.ArrayList;

public class ExMailMergeCustom extends ApiExampleBase {
    //ExStart
    //ExFor:IMailMergeDataSource
    //ExFor:IMailMergeDataSource.TableName
    //ExFor:IMailMergeDataSource.MoveNext
    //ExFor:IMailMergeDataSource.GetValue
    //ExFor:IMailMergeDataSource.GetChildDataSource
    //ExFor:MailMerge.Execute(IMailMergeDataSource)
    //ExSummary:Performs mail merge from a custom data source.
    @Test //ExSkip
    public void mailMergeCustomDataSource() throws Exception {
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

        doc.save(getArtifactsDir() + "MailMerge.CustomDataSource.doc");
    }

    /**
     * An example of a "data entity" class in your application.
     */
    public class Customer {
        public Customer(final String aFullName, final String anAddress) {
            mFullName = aFullName;
            mAddress = anAddress;
        }

        public String getFullName() {
            return mFullName;
        }

        public void setFullName(final String value) {
            mFullName = value;
        }

        public String getAddress() {
            return mAddress;
        }

        public void setAddress(final String value) {
            mAddress = value;
        }

        private String mFullName;
        private String mAddress;
    }

    /**
     * An example of a typed collection that contains your "data" objects.
     */
    public class CustomerList extends ArrayList {
        public Customer get(final int index) {
            return (Customer) super.get(index);
        }

        public void set(final int index, final Customer value) {
            super.set(index, value);
        }
    }

    /**
     * A custom mail merge data source that you implement to allow Aspose.Words
     * to mail merge data from your Customer objects into Microsoft Word documents.
     */
    public class CustomerMailMergeDataSource implements IMailMergeDataSource {
        public CustomerMailMergeDataSource(final CustomerList customers) {
            mCustomers = customers;

            // When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1;
        }

        /**
         * The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
         */
        public String getTableName() {
            return "Customer";
        }

        /**
         * Aspose.Words calls this method to get a value for every data field.
         */
        public boolean getValue(final String fieldName, final Ref<Object> fieldValue) throws Exception {
            if (fieldName.equals("FullName")) {
                fieldValue.set(mCustomers.get(mRecordIndex).getFullName());
                return true;
            } else if (fieldName.equals("Address")) {
                fieldValue.set(mCustomers.get(mRecordIndex).getAddress());
                return true;
            } else {
                // A field with this name was not found,
                // return false to the Aspose.Words mail merge engine.
                fieldValue.set(null);
                return false;
            }
        }

        /**
         * A standard implementation for moving to a next record in a collection.
         */
        public boolean moveNext() throws Exception {
            if (!isEof()) mRecordIndex++;

            return (!isEof());
        }

        public IMailMergeDataSource getChildDataSource(final String tableName) {
            return null;
        }

        private boolean isEof() {
            return (mRecordIndex >= mCustomers.size());
        }

        private final CustomerList mCustomers;
        private int mRecordIndex;
    }
    //ExEnd
}

