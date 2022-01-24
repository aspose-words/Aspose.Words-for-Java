package DocsExamples.LINQ_Reporting_Engine.Helpers;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.Manager;
import java.util.Iterator;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.Client;
import DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects.Contract;
import com.aspose.ms.System.DateTime;
import com.aspose.ms.System.IO.File;


class Common extends DocsExamplesBase
{
    /// <summary>
    /// Return the first manager from Managers, which is an enumeration of instances of the Manager class. 
    /// </summary>        
    public static Manager getManager() throws Exception
    {
        //ExStart:GetManager
        Iterator<Manager> managers = getManagers().iterator();
        managers.hasNext();
        
        return managers.next();
        //ExEnd:GetManager
    }

    /// <summary>
    /// Return an enumeration of instances of the Client class. 
    /// </summary>        
    public static Iterable<Client> getClients() throws Exception
    {
        //ExStart:GetClients
        for (Manager manager : getManagers())
        {
            for (Contract contract : manager.getContracts())
                yield return contract.getClient();
        }
        //ExEnd:GetClients
    }

    /// <summary>
    ///  Return an enumeration of instances of the Manager class.
    /// </summary>
    public static Iterable<Manager> getManagers() throws Exception
    {
        //ExStart:GetManagers
        Manager manager = new Manager(); { manager.setName("John Smith"); manager.setAge(36); manager.setPhoto(photo()); }
        manager.setContracts(new Contract[]
        {
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                    {
                        manager.getContracts().getClient().setName("A Company"); manager.getContracts().getClient().setCountry("Australia");
                        manager.getContracts().getClient().setLocalAddress("219-241 Cleveland St STRAWBERRY HILLS  NSW  1427");
                    }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(1200000f); manager.getContracts().setDate(new DateTime(2015, 1, 1));
            },
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                    {
                        manager.getContracts().getClient().setName("B Ltd."); manager.getContracts().getClient().setCountry("Brazil");
                        manager.getContracts().getClient().setLocalAddress("Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680");
                    }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(750000f); manager.getContracts().setDate(new DateTime(2015, 4, 1));
            },
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                    {
                        manager.getContracts().getClient().setName("C & D"); manager.getContracts().getClient().setCountry("Canada");
                        manager.getContracts().getClient().setLocalAddress("101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6");
                    }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(350000f); manager.getContracts().setDate(new DateTime(2015, 7, 1));
            }
        });
        yield return manager;

        manager = new Manager(); { manager.setName("Tony Anderson"); manager.setAge(37); manager.setPhoto(photo()); }
        manager.setContracts(new Contract[]
        {
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                        { manager.getContracts().getClient().setName("E Corp."); manager.getContracts().getClient().setLocalAddress("445 Mount Eden Road Mount Eden Auckland 1024"); }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(650000f); manager.getContracts().setDate(new DateTime(2015, 2, 1));
            },
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                        { manager.getContracts().getClient().setName("F & Partners"); manager.getContracts().getClient().setLocalAddress("20 Greens Road Tuahiwi Kaiapoi 7691 "); }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(550000f); manager.getContracts().setDate(new DateTime(2015, 8, 1));
            },
        });
        yield return manager;

        manager = new Manager(); { manager.setName("July James"); manager.setAge(38); manager.setPhoto(photo()); }
        manager.setContracts(new Contract[]
        {
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                        { manager.getContracts().getClient().setName("G & Co."); manager.getContracts().getClient().setCountry("Greece"); manager.getContracts().getClient().setLocalAddress("Karkisias 6 GR-111 42  ATHINA GRÉCE"); }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(350000f); manager.getContracts().setDate(new DateTime(2015, 2, 1));
            },
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                    {
                        manager.getContracts().getClient().setName("H Group"); manager.getContracts().getClient().setCountry("Hungary");
                        manager.getContracts().getClient().setLocalAddress("Budapest Fiktív utca 82., IV. em./28.2806");
                    }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(250000f); manager.getContracts().setDate(new DateTime(2015, 5, 1));
            },
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                        { manager.getContracts().getClient().setName("I & Sons"); manager.getContracts().getClient().setLocalAddress("43 Vogel Street Roslyn Palmerston North 4414"); }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(100000f); manager.getContracts().setDate(new DateTime(2015, 7, 1));
            },
            new Contract();
            {
                manager.getContracts().setClient(new Client());
                    {
                        manager.getContracts().getClient().setName("J Ent."); manager.getContracts().getClient().setCountry("Japan");
                        manager.getContracts().getClient().setLocalAddress("Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan");
                    }
                manager.getContracts().setManager(manager); manager.getContracts().setPrice(100000f); manager.getContracts().setDate(new DateTime(2015, 8, 1));
            }
        });
        yield return manager;
        //ExEnd:GetManagers
    }

    /// <summary>
    /// Return an array of photo bytes. 
    /// </summary>
    private static byte[] photo() throws Exception
    {
        //ExStart:Photo
        // Load the photo and read all bytes
        byte[] logo = com.aspose.ms.System.IO.File.readAllBytes(getImagesDir() + "Logo.jpg");
        
        return logo;
        //ExEnd:Photo
    }

    /// <summary>
    ///  Return an enumeration of instances of the Contract class.
    /// </summary>
    public static Iterable<Contract> getContracts() throws Exception
    {
        //ExStart:GetContracts
        for (Manager manager : getManagers())
        {
            for (Contract contract : manager.getContracts())
                yield return contract;
        }
        //ExEnd:GetContracts
    }
}
