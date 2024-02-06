package TestData;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import TestData.TestClasses.*;

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

    public static ArrayList<ShareTestClass> getShares() {
        ArrayList<ShareTestClass> shares = new ArrayList<ShareTestClass>() {
            {
                new ShareTestClass("Technology", "Consumer Electronics", "AAPL", 6.602835, -0.0054);
                new ShareTestClass("Technology", "Software - Infrastructure", "MSFT", 5.832072, -0.005);
                new ShareTestClass("Technology", "Software - Infrastructure", "ADBE", 0.562561, -0.0274);
                new ShareTestClass("Technology", "Semiconductors", "NVDA", 1.335994, -0.0074);
                new ShareTestClass("Technology", "Semiconductors", "QCOM", 0.462198, 0.0248);
                new ShareTestClass("Communication Services", "Internet Content & Information", "GOOG", 3.771651, 0.011);
                new ShareTestClass("Communication Services", "Entertainment", "DIS", 0.575768, 0.0102);
                new ShareTestClass("Communication Services", "Entertainment", "WBD", 0.116579, -0.0165);
                new ShareTestClass("Consumer Cyclical", "Internet Retail", "AMZN", 3.011482, 0.044);
                new ShareTestClass("Consumer Cyclical", "Auto Manufactures", "TSLA", 1.816734, -0.0018);
                new ShareTestClass("Consumer Cyclical", "Auto Manufactures", "GM", 0.160205, 0.0026);
                new ShareTestClass("Financial", "Credit Services", "V", 1.1, 0.005);
            }
        };

        return shares;
    }

    public static ArrayList<ShareQuoteTestClass> getShareQuotes() {
        ArrayList<ShareQuoteTestClass> shareQuotes = new ArrayList<ShareQuoteTestClass>() {
            {
                new ShareQuoteTestClass(45131, 15232450, 171.32, 172.50, 170.69, 171.98);
                new ShareQuoteTestClass(45132, 13962990, 172.20, 172.70, 171.40, 171.86);
                new ShareQuoteTestClass(45133, 14902060, 171.86, 171.93, 170.31, 171.35);
                new ShareQuoteTestClass(45134, 16962540, 171.64, 173.10, 171.35, 172.00);
                new ShareQuoteTestClass(45135, 15588280, 171.98, 172.40, 170.00, 171.44);
            }
        };

        return shareQuotes;
    }
}
