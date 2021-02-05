package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.IMailMergeDataSourceRoot;
import com.aspose.words.ref.Ref;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.HashMap;

public class ExMailMergeCustom extends ApiExampleBase {
    //ExStart
    //ExFor:IMailMergeDataSource
    //ExFor:IMailMergeDataSource.TableName
    //ExFor:IMailMergeDataSource.MoveNext
    //ExFor:IMailMergeDataSource.GetValue
    //ExFor:IMailMergeDataSource.GetChildDataSource
    //ExFor:MailMerge.Execute(IMailMergeDataSourceCore)
    //ExSummary:Performs mail merge from a custom data source.
    @Test //ExSkip
    public void customDataSource() throws Exception {
        // Create a destination document for the mail merge
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.insertField(" MERGEFIELD FullName ");
        builder.insertParagraph();
        builder.insertField(" MERGEFIELD Address ");

        // Create some data that we will use in the mail merge
        CustomerList customers = new CustomerList();
        customers.add(new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        customers.add(new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        // To be able to mail merge from your own data source, it must be wrapped
        // into an object that implements the IMailMergeDataSource interface
        CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

        // Now you can pass your data source into Aspose.Words
        doc.getMailMerge().execute(customersDataSource);

        doc.save(getArtifactsDir() + "MailMergeCustom.CustomDataSource.docx");
        testCustomDataSource(customers, new Document(getArtifactsDir() + "MailMergeCustom.CustomDataSource.docx")); //ExSkip
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
                // return false to the Aspose.Words mail merge engine
                fieldValue.set(null);
                return false;
            }
        }

        /**
         * A standard implementation for moving to a next record in a collection.
         */
        public boolean moveNext() {
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

    private void testCustomDataSource(CustomerList customerList, Document doc) {
        String[][] mergeData = new String[customerList.size()][];

        for (int i = 0; i < customerList.size(); i++)
            mergeData[i] = new String[]{customerList.get(i).getFullName(), customerList.get(i).getAddress()};

        TestUtil.mailMergeMatchesArray(mergeData, doc, true);
    }

    //ExStart
    //ExFor:IMailMergeDataSourceRoot
    //ExFor:IMailMergeDataSourceRoot.GetDataSource(String)
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSourceRoot)
    //ExSummary:Performs mail merge from a custom data source with master-detail data.
    @Test //ExSkip
    public void customDataSourceRoot() throws Exception {
        // Create a document with two mail merge regions named "Washington" and "Seattle"
        Document doc = createSourceDocumentWithMailMergeRegions(new String[]{"Washington", "Seattle"});

        // Create two data sources
        EmployeeList employeesWashingtonBranch = new EmployeeList();
        employeesWashingtonBranch.add(new Employee("John Doe", "Sales"));
        employeesWashingtonBranch.add(new Employee("Jane Doe", "Management"));

        EmployeeList employeesSeattleBranch = new EmployeeList();
        employeesWashingtonBranch.add(new Employee("John Cardholder", "Management"));
        employeesWashingtonBranch.add(new Employee("Joe Bloggs", "Sales"));

        // Register our data sources by name in a data source root
        DataSourceRoot sourceRoot = new DataSourceRoot();
        sourceRoot.registerSource("Washington", new EmployeeListMailMergeSource(employeesWashingtonBranch));
        sourceRoot.registerSource("Seattle", new EmployeeListMailMergeSource(employeesSeattleBranch));

        // Since we have consecutive mail merge regions, we would normally have to perform two mail merges
        // However, one mail merge source data root call every relevant data source and merge automatically
        doc.getMailMerge().executeWithRegions(sourceRoot);

        doc.save(getArtifactsDir() + "MailMergeCustom.CustomDataSourceRoot.docx");
    }

    /// <summary>
    /// Create document that contains consecutive mail merge regions, with names designated by the input array,
    /// for a data table of employees.
    /// </summary>
    private static Document createSourceDocumentWithMailMergeRegions(String[] regions) throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (String s : regions) {
            builder.writeln("\n" + s + " branch: ");
            builder.insertField(" MERGEFIELD TableStart:" + s);
            builder.insertField(" MERGEFIELD FullName");
            builder.write(", ");
            builder.insertField(" MERGEFIELD Department");
            builder.insertField(" MERGEFIELD TableEnd:" + s);
        }

        return doc;
    }

    /// <summary>
    /// An example of a "data entity" class in your application.
    /// </summary>
    private static class Employee {
        public Employee(String aFullName, String aDepartment) {
            mFullName = aFullName;
            mDepartment = aDepartment;
        }

        public String getFullName() {
            return mFullName;
        }

        private final String mFullName;

        public String getDepartment() {
            return mDepartment;
        }

        private final String mDepartment;
    }

    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    private static class EmployeeList extends ArrayList {
        public Employee get(int index) {
            return (Employee) super.get(index);
        }

        public void set(int index, Employee value) {
            super.set(index, value);
        }
    }

    /// <summary>
    /// Data source root that can be passed directly into a mail merge which can register and contain many child data sources.
    /// These sources must all implement IMailMergeDataSource, and are registered and differentiated by a name
    /// which corresponds to a mail merge region that will read the respective data.
    /// </summary>
    private static class DataSourceRoot implements IMailMergeDataSourceRoot {
        public IMailMergeDataSource getDataSource(String tableName) {
            EmployeeListMailMergeSource source = mSources.get(tableName);
            source.reset();
            return mSources.get(tableName);
        }

        public void registerSource(String sourceName, EmployeeListMailMergeSource source) {
            mSources.put(sourceName, source);
        }

        private final HashMap<String, EmployeeListMailMergeSource> mSources = new HashMap<>();
    }

    /// <summary>
    /// Custom mail merge data source.
    /// </summary>
    private static class EmployeeListMailMergeSource implements IMailMergeDataSource {
        public EmployeeListMailMergeSource(EmployeeList employees) {
            mEmployees = employees;
            mRecordIndex = -1;
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        public boolean moveNext() {
            if (!isEof())
                mRecordIndex++;

            return (!isEof());
        }

        private boolean isEof() {
            return (mRecordIndex >= mEmployees.size());
        }

        public void reset() {
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String getTableName() {
            return "Employees";
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, Ref<Object> fieldValue) {
            switch (fieldName) {
                case "FullName":
                    fieldValue.set(mEmployees.get(mRecordIndex).getFullName());
                    return true;
                case "Department":
                    fieldValue.set(mEmployees.get(mRecordIndex).getDepartment());
                    return true;
                default:
                    // A field with this name was not found,
                    // return false to the Aspose.Words mail merge engine
                    fieldValue.set(null);
                    return false;
            }
        }

        /// <summary>
        /// Child data sources are for nested mail merges.
        /// </summary>
        public IMailMergeDataSource getChildDataSource(String tableName) {
            throw new UnsupportedOperationException();
        }

        private final EmployeeList mEmployees;
        private int mRecordIndex;
    }
    //ExEnd

}

