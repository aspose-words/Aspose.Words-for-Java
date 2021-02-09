package TestData;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.TestClasses.ClientTestClass;
import TestData.TestClasses.ContractTestClass;
import TestData.TestClasses.ManagerTestClass;

import java.time.LocalDate;
import java.util.ArrayList;

public class Common {
    private static ArrayList<ManagerTestClass> managers = new ArrayList<>();
    private static ArrayList<ContractTestClass> contracts = new ArrayList<>();
    private static ArrayList<ClientTestClass> clients = new ArrayList<>();

    static {
        // --------------------------------------------------
        // First manager
        // --------------------------------------------------
        ManagerTestClass firstManager = new ManagerTestClass();
        firstManager.setName("John Smith");
        firstManager.setAge(36);

        ArrayList<ContractTestClass> contracts = new ArrayList<>();
        {
            contracts.add(new ContractTestClass());
            {
                contracts.get(0).setManager(firstManager);
                contracts.get(0).setPrice(1200000f);
                contracts.get(0).setDate(LocalDate.of(2017, 1, 1));
                contracts.get(0).setClient(new ClientTestClass("A Company", "Australia", "219-241 Cleveland St STRAWBERRY HILLS  NSW  1427"));
            }
            contracts.add(new ContractTestClass());
            {
                contracts.get(1).setManager(firstManager);
                contracts.get(1).setPrice(750000f);
                contracts.get(1).setDate(LocalDate.of(2017, 4, 1));
                contracts.get(1).setClient(new ClientTestClass("B Ltd.", "Brazil", "Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"));
            }
            contracts.add(new ContractTestClass());
            {
                contracts.get(2).setManager(firstManager);
                contracts.get(2).setPrice(350000f);
                contracts.get(2).setDate(LocalDate.of(2017, 7, 1));
                contracts.get(2).setClient(new ClientTestClass("C & D", "Canada", "101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6"));
            }
        }

        firstManager.setContracts(contracts);

        // --------------------------------------------------
        // Second manager
        // --------------------------------------------------
        ManagerTestClass secondManager = new ManagerTestClass();
        secondManager.setName("Tony Anderson");
        secondManager.setAge(37);

        contracts = new ArrayList<>();
        {
            contracts.add(new ContractTestClass());
            {
                contracts.get(0).setManager(secondManager);
                contracts.get(0).setPrice(650000f);
                contracts.get(0).setDate(LocalDate.of(2017, 2, 1));
                contracts.get(0).setClient(new ClientTestClass("E Corp.", "445 Mount Eden Road Mount Eden Auckland 1024"));
            }
            contracts.add(new ContractTestClass());
            {
                contracts.get(1).setManager(secondManager);
                contracts.get(1).setPrice(550000f);
                contracts.get(1).setDate(LocalDate.of(2017, 8, 1));
                contracts.get(1).setClient(new ClientTestClass("F & Partners", "20 Greens Road Tuahiwi Kaiapoi 7691"));
            }
        }

        secondManager.setContracts(contracts);

        // --------------------------------------------------
        // Third manager
        // --------------------------------------------------
        ManagerTestClass thirdManager = new ManagerTestClass();
        thirdManager.setName("July James");
        thirdManager.setAge(38);

        contracts = new ArrayList<>();
        {
            contracts.add(new ContractTestClass());
            {
                contracts.get(0).setManager(thirdManager);
                contracts.get(0).setPrice(350000f);
                contracts.get(0).setDate(LocalDate.of(2017, 2, 1));
                contracts.get(0).setClient(new ClientTestClass("G & Co.", "Greece", "Karkisias 6 GR-111 42  ATHINA GRÉCE"));
            }
            contracts.add(new ContractTestClass());
            {
                contracts.get(1).setManager(thirdManager);
                contracts.get(1).setPrice(250000f);
                contracts.get(1).setDate(LocalDate.of(2017, 5, 1));
                contracts.get(1).setClient(new ClientTestClass("H Group", "Hungary", "Budapest Fiktív utca 82., IV. em./28.2806"));
            }
            contracts.add(new ContractTestClass());
            {
                contracts.get(2).setManager(thirdManager);
                contracts.get(2).setPrice(100000f);
                contracts.get(2).setDate(LocalDate.of(2017, 7, 1));
                contracts.get(2).setClient(new ClientTestClass("I & Sons", "43 Vogel Street Roslyn Palmerston North 4414"));
            }
            contracts.add(new ContractTestClass());
            {
                contracts.get(3).setManager(thirdManager);
                contracts.get(3).setPrice(100000f);
                contracts.get(3).setDate(LocalDate.of(2017, 8, 1));
                contracts.get(3).setClient(new ClientTestClass("J Ent.", "Japan", "Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"));
            }
        }

        thirdManager.setContracts(contracts);

        managers.add(firstManager);
        managers.add(secondManager);
        managers.add(thirdManager);
    }

    public static ArrayList<ManagerTestClass> getEmptyManagers() {
        return new ArrayList<>();
    }

    public static ArrayList<ManagerTestClass> getManagers() {
        return managers;
    }

    public static ArrayList<ClientTestClass> getClients() {
        for (ManagerTestClass manager : getManagers()) {
            for (ContractTestClass contract : manager.getContracts())
                clients.add(contract.getClient());
        }

        return clients;
    }

    public static ArrayList<ContractTestClass> getContracts() {
        for (ManagerTestClass manager : getManagers()) {
            for (ContractTestClass contract : manager.getContracts())
                contracts.add(contract);
        }

        return contracts;
    }
}
