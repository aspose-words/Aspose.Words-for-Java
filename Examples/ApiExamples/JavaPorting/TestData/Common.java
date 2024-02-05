package ApiExamples.TestData;

// ********* THIS FILE IS AUTO PORTED *********

import ApiExamples.TestData.TestClasses.ManagerTestClass;
import ApiExamples.TestData.TestClasses.ContractTestClass;
import ApiExamples.TestData.TestClasses.ClientTestClass;
import com.aspose.ms.System.DateTime;


public /*static*/ class Common
{
	/* Simulation of static class by using private constructor */
	private Common()
	{}

    public static Iterable<ManagerTestClass> getManagers()
    {
        ManagerTestClass manager = new ManagerTestClass();
        {
            manager.setName("John Smith");
            manager.setAge(36);
        }

        manager.setContracts(new ContractTestClass[]
        {
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("A Company");
                        manager.getContracts().getClient().setCountry("Australia");
                        manager.getContracts().getClient().setLocalAddress("219-241 Cleveland St STRAWBERRY HILLS  NSW  1427");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(1200000f);
                manager.getContracts().setDate(new DateTime(2017, 1, 1));
            },
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("B Ltd.");
                        manager.getContracts().getClient().setCountry("Brazil");
                        manager.getContracts().getClient().setLocalAddress("Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(750000f);
                manager.getContracts().setDate(new DateTime(2017, 4, 1));
            },
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("C & D");
                        manager.getContracts().getClient().setCountry("Canada");
                        manager.getContracts().getClient().setLocalAddress("101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(350000f);
                manager.getContracts().setDate(new DateTime(2017, 7, 1));
            }
        });

        yield return manager;

        manager = new ManagerTestClass();
        {
            manager.setName("Tony Anderson");
            manager.setAge(37);
        }

        manager.setContracts(new ContractTestClass[]
        {
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("E Corp.");
                        manager.getContracts().getClient().setLocalAddress("445 Mount Eden Road Mount Eden Auckland 1024");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(650000f);
                manager.getContracts().setDate(new DateTime(2017, 2, 1));
            },
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("F & Partners");
                        manager.getContracts().getClient().setLocalAddress("20 Greens Road Tuahiwi Kaiapoi 7691 ");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(550000f);
                manager.getContracts().setDate(new DateTime(2017, 8, 1));
            }
        });

        yield return manager;

        manager = new ManagerTestClass();
        {
            manager.setName("July James");
            manager.setAge(38);
        }

        manager.setContracts(new ContractTestClass[]
        {
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("G & Co.");
                        manager.getContracts().getClient().setCountry("Greece");
                        manager.getContracts().getClient().setLocalAddress("Karkisias 6 GR-111 42  ATHINA GRÉCE");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(350000f);
                manager.getContracts().setDate(new DateTime(2017, 2, 1));
            },
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("H Group");
                        manager.getContracts().getClient().setCountry("Hungary");
                        manager.getContracts().getClient().setLocalAddress("Budapest Fiktív utca 82., IV. em./28.2806");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(250000f);
                manager.getContracts().setDate(new DateTime(2017, 5, 1));
            },
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("I & Sons");
                        manager.getContracts().getClient().setLocalAddress("43 Vogel Street Roslyn Palmerston North 4414");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(100000f);
                manager.getContracts().setDate(new DateTime(2017, 7, 1));
            },
            new ContractTestClass();
            {
                manager.getContracts().setClient(new ClientTestClass());
                    {
                        manager.getContracts().getClient().setName("J Ent.");
                        manager.getContracts().getClient().setCountry("Japan");
                        manager.getContracts().getClient().setLocalAddress("Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan");
                    }
                manager.getContracts().setManager(manager);
                manager.getContracts().setPrice(100000f);
                manager.getContracts().setDate(new DateTime(2017, 8, 1));
            }
        });

        yield return manager;
    }

    public static Iterable<ManagerTestClass> getEmptyManagers()
    {
        return Enumerable.<ManagerTestClass>Empty();
    }

    public static Iterable<ClientTestClass> getClients()
    {
        for (ManagerTestClass manager : getManagers())
        {
            for (ContractTestClass contract : manager.getContracts())
                yield return contract.getClient();
        }
    }

    public static Iterable<ContractTestClass> getContracts()
    {
        for (ManagerTestClass manager : getManagers())
        {
            for (ContractTestClass contract : manager.getContracts())
                yield return contract;
        }
    }

    public static ShareTestClass[] getShares()
    {
        return new ShareTestClass[]
        {
            new ShareTestClass("Technology", "Consumer Electronics", "AAPL", 6.602835, -0.0054),
            new ShareTestClass("Technology", "Software - Infrastructure", "MSFT", 5.832072, -0.005),
            new ShareTestClass("Technology", "Software - Infrastructure", "ADBE", 0.562561, -0.0274),
            new ShareTestClass("Technology", "Semiconductors", "NVDA", 1.335994, -0.0074),
            new ShareTestClass("Technology", "Semiconductors", "QCOM", 0.462198, 0.0248),
            new ShareTestClass("Communication Services", "Internet Content & Information", "GOOG", 3.771651, 0.011),
            new ShareTestClass("Communication Services", "Entertainment", "DIS", 0.575768, 0.0102),
            new ShareTestClass("Communication Services", "Entertainment", "WBD", 0.116579, -0.0165),
            new ShareTestClass("Consumer Cyclical", "Internet Retail", "AMZN", 3.011482, 0.044),
            new ShareTestClass("Consumer Cyclical", "Auto Manufactures", "TSLA", 1.816734, -0.0018),
            new ShareTestClass("Consumer Cyclical", "Auto Manufactures", "GM", 0.160205, 0.0026),
            new ShareTestClass("Financial", "Credit Services", "V", 1.1, 0.005)
        };
    }

    public static ShareQuoteTestClass[] getShareQuotes()
    {
        return new ShareQuoteTestClass[]
        {
            new ShareQuoteTestClass(45131, 15232450, 171.32, 172.50, 170.69, 171.98),
            new ShareQuoteTestClass(45132, 13962990, 172.20, 172.70, 171.40, 171.86),
            new ShareQuoteTestClass(45133, 14902060, 171.86, 171.93, 170.31, 171.35),
            new ShareQuoteTestClass(45134, 16962540, 171.64, 173.10, 171.35, 172.00),
            new ShareQuoteTestClass(45135, 15588280, 171.98, 172.40, 170.00, 171.44)
        };
    }
}
