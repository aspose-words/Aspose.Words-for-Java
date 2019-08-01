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
}
