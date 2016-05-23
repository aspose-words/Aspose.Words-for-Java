package com.aspose.words.examples.linq;
import java.io.File;
import java.io.FileInputStream;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;

import java.sql.*;
import java.util.*;
import java.util.Date;

public class Common {
    public static List<Manager> managers = new ArrayList<Manager>();

    /// <summary>
    /// Return first manager from Managers which is an enumeration of instances of the Manager class.
    /// </summary>
    public static Manager GetManager(){
        for (Manager manager : GetManagers()) {
            return manager;
        }
        return null;
    }

    /// <summary>
    /// Return an dataset of the Client class.
    /// </summary>
    public static DataSet GetClients() throws Exception
    {
        // Create a new data set
        DataSet dataSet = new DataSet("DS");

        // Add a new table to store clients
        DataTable dt = new DataTable("clients");

        // Add columns
        dt.getColumns().add("Name");
        dataSet.getTables().add(dt);

        // Populate the data in table
        for (Manager manager : GetManagers()) {
            List<Contract> listOfContracts = manager.getContracts();
            for (Contract contract : listOfContracts) {
                DataRow row = dt.newRow();
                row.set("Name", contract.getClient().getName());
                dt.getRows().add(row);
            }
        }
        return dataSet;
    }

    /// <summary>
    /// Return an enumeration of instances of the Manager class.
    /// </summary>
    private static List<Manager> GetManagers() {

        Manager manager = new Manager();
        manager.setName("John Smith");
        manager.setAge(36);
        manager.setPhoto(Photo());

        Contract contract1 = new Contract();
        Client client1 = new Client();
        client1.setName("A Company");
        contract1.setClient(client1);
        contract1.setManager(manager);
        contract1.setPrice(1200000);
        contract1.setDate(new Date(2015, 1, 1));

        Contract contract2 = new Contract();
        Client client2 = new Client();
        client2.setName("B Ltd.");
        contract2.setClient(client2);
        contract2.setManager(manager);
        contract2.setPrice(750000);
        contract2.setDate(new Date(2015, 4, 1));

        Contract contract3 = new Contract();
        Client client3 = new Client();
        client3.setName("C & D");
        contract3.setClient(client3);
        contract3.setManager(manager);
        contract3.setPrice(350000);
        contract3.setDate(new Date(2015, 7, 1));

        ArrayList<Contract> contracts = new ArrayList<Contract>();
        contracts.add(contract1);
        contracts.add(contract2);
        contracts.add(contract3);

        manager.setContracts(contracts);
        managers.add(manager);

        manager = new Manager();
        manager.setName("Tony Anderson");
        manager.setAge(37);
        manager.setPhoto(Photo());
        Contract contract4 = new Contract();
        Client client4 = new Client();
        client4.setName("E Corp.");
        contract4.setClient(client4);
        contract4.setManager(manager);
        contract4.setPrice(650000);
        Date date = new Date(2015, 2, 1);
        contract4.setDate(date);
        Contract contract5 = new Contract();
        Client client5 = new Client();
        client5.setName("F & Partners");
        contract5.setClient(client5);
        contract5.setManager(manager);
        contract5.setPrice(550000);
        contract5.setDate(new Date(2015, 8, 1));

        ArrayList<Contract> contracts2 = new ArrayList<Contract>();
        contracts2.add(contract4);
        contracts2.add(contract5);
        manager.setContracts(contracts2);
        managers.add(manager);

        manager = new Manager();
        manager.setName("July James");
        manager.setAge(38);
        manager.setPhoto(Photo());
        Contract contract6 = new Contract();
        Client client6 = new Client();
        client6.setName("G & Co.");
        contract6.setClient(client6);
        contract6.setManager(manager);
        contract6.setPrice(350000);
        contract6.setDate(new Date(2015, 2, 1));
        Contract contract7 = new Contract();
        Client client7 = new Client();
        client7.setName("H Group");
        contract7.setClient(client7);
        contract7.setManager(manager);
        contract7.setPrice(250000);
        contract7.setDate(new Date(2015, 5, 1));
        Contract contract8 = new Contract();
        Client client8 = new Client();
        client8.setName("I & Sons");
        contract8.setClient(client8);
        contract8.setManager(manager);
        contract8.setPrice(100000);
        contract8.setDate(new Date(2015, 7, 1));
        Contract contract9 = new Contract();
        Client client9 = new Client();
        client9.setName("J Ent.");
        contract9.setClient(client9);
        contract9.setManager(manager);
        contract9.setPrice(100000);
        contract9.setDate(new Date(2015, 8, 1));

        ArrayList<Contract> contracts3 = new ArrayList<Contract>();
        contracts3.add(contract6);
        contracts3.add(contract7);

        contracts3.add(contract8);
        contracts3.add(contract9);

        manager.setContracts(contracts3);
        managers.add(manager);
        return managers;
    }
    /// <summary>
    /// Return an array of photo bytes.
    /// </summary>
    private static byte[] Photo()
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Common.class);
        File file = new File(dataDir + "photo.png");
        return readContentIntoByteArray(file);
    }
    private static byte[] readContentIntoByteArray(File file)
    {
        FileInputStream fileInputStream = null;
        byte[] bFile = new byte[(int) file.length()];
        try
        {
            //convert file into array of bytes
            fileInputStream = new FileInputStream(file);
            fileInputStream.read(bFile);
            fileInputStream.close();
            for (int i = 0; i < bFile.length; i++)
            {
                //System.out.print((char) bFile[i]);
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        return bFile;
    }
    /// <summary>
    ///  Return an dataset of the Contract class.
    /// </summary>
    public static DataSet GetContracts() throws Exception
    {
        // Create a new data set
        DataSet ds = new DataSet("ds");

        // Add a new table to store contracts
        DataTable dtContracts = new DataTable("Contracts");

        // Add a new table to store managers
        DataTable dtManagers = new DataTable("Managers");

        // Add a new table to store clients
        DataTable dtClients = new DataTable("Clients");

        // Add columns to Managers table
        dtManagers.getColumns().add("Id", int.class);
        dtManagers.getColumns().add("Name");
        dtManagers.getColumns().add("Age", int.class);
        dtManagers.getColumns().add("Photo", byte[].class);
        ds.getTables().add(dtManagers);

        // Add columns to Contracts table
        dtContracts.getColumns().add("Id", int.class);
        dtContracts.getColumns().add("ClientId", int.class);
        dtContracts.getColumns().add("ManagerId", int.class);
        dtContracts.getColumns().add("Price", float.class);
        dtContracts.getColumns().add("Date", Date.class);
        ds.getTables().add(dtContracts);

        // Add columns to Clients table
        dtClients.getColumns().add("Id", int.class);
        dtClients.getColumns().add("Name");
        ds.getTables().add(dtClients);
        ds.getRelations().add(dtClients,dtContracts, "Id","ClientId");
        ds.getRelations().add(dtManagers,dtContracts, "Id","ManagerId");


        int managerCounter = 1;
        int contractCounter =1;
        int clientCounter = 1;
        for (Manager manager : GetManagers()) {
            // Add data row to managers table.
            DataRow managerRow = dtManagers.newRow();
            managerRow.set("Id", managerCounter);
            managerRow.set("Name", manager.getName());
            managerRow.set("Age", manager.getAge());
            managerRow.set("Photo", manager.getPhoto());
            dtManagers.getRows().add(managerRow);

            for (Contract contract : manager.getContracts()) {
                DataRow contractRow = dtContracts.newRow();
                DataRow clientRow = dtClients.newRow();

                clientRow.set("Id", clientCounter);
                clientRow.set("Name", contract.getClient().getName());
                dtClients.getRows().add(clientRow);

                contractRow.set("Id", contractCounter);
                contractRow.set("ClientId", clientCounter);
                contractRow.set("ManagerId", managerCounter);
                contractRow.set("Price", contract.getPrice());
                contractRow.set("Date", contract.getDate());
                dtContracts.getRows().add(contractRow);
                clientCounter += 1;
                contractCounter += 1;


            }
            managerCounter += 1;
        }
        return ds;
    }


}
