// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.DocumentBuilder;
import com.aspose.words.IMailMergeDataSourceRoot;
import com.aspose.ms.System.Collections.msDictionary;
import java.util.HashMap;
import com.aspose.words.ref.Ref;


@Test
public class ExMailMergeCustom extends ApiExampleBase
{
    //ExStart
    //ExFor:IMailMergeDataSource
    //ExFor:IMailMergeDataSource.TableName
    //ExFor:IMailMergeDataSource.MoveNext
    //ExFor:IMailMergeDataSource.GetValue
    //ExFor:IMailMergeDataSource.GetChildDataSource
    //ExFor:MailMerge.Execute(IMailMergeDataSourceCore)
    //ExSummary:Performs mail merge from a custom data source.
    @Test //ExSkip
    public void mailMergeCustomDataSource() throws Exception
    {
        // Create some data that we will use in the mail merge
        CustomerList customers = new CustomerList();
        msArrayList.add(customers, new Customer("Thomas Hardy", "120 Hanover Sq., London"));
        msArrayList.add(customers, new Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"));

        // Open the template document
        Document doc = new Document(getMyDir() + "MailMerge.CustomDataSource.doc");

        // To be able to mail merge from your own data source, it must be wrapped
        // into an object that implements the IMailMergeDataSource interface
        CustomerMailMergeDataSource customersDataSource = new CustomerMailMergeDataSource(customers);

        // Now you can pass your data source into Aspose.Words
        doc.getMailMerge().execute(customersDataSource);

        doc.save(getArtifactsDir() + "MailMerge.CustomDataSource.doc");
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
        }

        public String getFullName() { return mFullName; }; public void setFullName(String value) { mFullName = value; };

        private String mFullName;
        public String getAddress() { return mAddress; }; public void setAddress(String value) { mAddress = value; };

        private String mAddress;
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
            return null;
        }

        private boolean isEof() { return (mRecordIndex >= mCustomers.size()); }

        private /*final*/ CustomerList mCustomers;
        private int mRecordIndex;
    }
    //ExEnd

    //ExStart
    //ExFor:IMailMergeDataSourceRoot
    //ExFor:IMailMergeDataSourceRoot.GetDataSource(String)
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSourceRoot)
    //ExSummary:Performs mail merge from a custom data source with master-detail data.
    @Test //ExSkip
    public void mailMergeCustomDataSourceRoot() throws Exception
    {
        // Create a document with two mail merge regions named "Washington" and "Seattle"
        Document doc = createSourceDocumentWithMailMergeRegions(new String[] { "Washington", "Seattle" });

        // Create two data sources
        EmployeeList employeesWashingtonBranch = new EmployeeList();
        msArrayList.add(employeesWashingtonBranch, new Employee("John Doe", "Sales"));
        msArrayList.add(employeesWashingtonBranch, new Employee("Jane Doe", "Management"));

        EmployeeList employeesSeattleBranch = new EmployeeList();
        msArrayList.add(employeesSeattleBranch, new Employee("John Cardholder", "Management"));
        msArrayList.add(employeesSeattleBranch, new Employee("Joe Bloggs", "Sales"));

        // Register our data sources by name in a data source root
        DataSourceRoot sourceRoot = new DataSourceRoot();
        sourceRoot.registerSource("Washington", new EmployeeListMailMergeSource(employeesWashingtonBranch));
        sourceRoot.registerSource("Seattle", new EmployeeListMailMergeSource(employeesSeattleBranch));

        // Since we have consecutive mail merge regions, we would normally have to perform two mail merges
        // However, one mail merge source data root call every relevant data source and merge automatically 
        doc.getMailMerge().executeWithRegions(sourceRoot);

        doc.save(getArtifactsDir() + "MailMerge.MailMergeCustomDataSourceRoot.docx");
    }

    /// <summary>
    /// Create document that contains consecutive mail merge regions, with names designated by the input array,
    /// for a data table of employees.
    /// </summary>
    private static Document createSourceDocumentWithMailMergeRegions(String[] regions) throws Exception
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (String s : regions)
        {
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
    private static class Employee
    {
        public Employee(String aFullName, String aDepartment)
        {
            mFullName = aFullName;
            mDepartment = aDepartment;
        }

        public String getFullName() { return mFullName; };

        private  String mFullName;
        public String getDepartment() { return mDepartment; };

        private  String mDepartment;
    }

    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    private static class EmployeeList extends ArrayList
    {
        public /*new*/ Employee get(int index) { return (Employee)super.get(index); }
        public /*new*/void set(int index, Employee value) { super.set(index, value); }
    }

    /// <summary>
    /// Data source root that can be passed directly into a mail merge which can register and contain many child data sources.
    /// These sources must all implement IMailMergeDataSource, and are registered and differentiated by a name
    /// which corresponds to a mail merge region that will read the respective data.
    /// </summary>
    private static class DataSourceRoot implements IMailMergeDataSourceRoot
    {
        public IMailMergeDataSource getDataSource(String tableName)
        {
            EmployeeListMailMergeSource source = mSources.get(tableName);
            source.reset();
            return mSources.get(tableName);
        }

        public void registerSource(String sourceName, EmployeeListMailMergeSource source)
        {
            msDictionary.add(mSources, sourceName, source);
        }

        private /*final*/ HashMap<String, EmployeeListMailMergeSource> mSources = new HashMap<String, EmployeeListMailMergeSource>();
    }

    /// <summary>
    /// Custom mail merge data source.
    /// </summary>
    private static class EmployeeListMailMergeSource implements IMailMergeDataSource
    {
        public EmployeeListMailMergeSource(EmployeeList employees)
        {
            mEmployees = employees;
            mRecordIndex = -1;
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

        private boolean isEof() { return (mRecordIndex >= mEmployees.size()); }

        public void reset()
        {
            mRecordIndex = -1;
        }

        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        public String getTableName() { return "Employees"; }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        public boolean getValue(String fieldName, /*out*/Ref<Object> fieldValue)
        {
            switch (gStringSwitchMap.of(fieldName))
            {
                case /*"FullName"*/0:
                    fieldValue.set(mEmployees.get(mRecordIndex).getFullName());
                    return true;
                case /*"Department"*/2:
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
        public IMailMergeDataSource getChildDataSource(String tableName)
        {
            throw new UnsupportedOperationException();
        }

        private /*final*/ EmployeeList mEmployees;
        private int mRecordIndex;
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"FullName",
		"Address",
		"Department"
	);

    //ExEnd
}
