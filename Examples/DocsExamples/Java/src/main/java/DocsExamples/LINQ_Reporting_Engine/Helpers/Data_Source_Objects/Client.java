package DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

//ExStart:Client
public class Client
{
    private String mName;
    private String mCountry;
    private String mLocalAddress;

    public Client(final String name, final String localAddress) {
        setName(name);
        setLocalAddress(localAddress);
    }

    public Client(final String name, final String country, final String localAddress) {
        setName(name);
        setCountry(country);
        setLocalAddress(localAddress);
    }

    public String getName() { return mName; }
    public String getCountry() { return mCountry; }
    public String getLocalAddress() { return mLocalAddress; }

    public void setName(String value) { mName = value; }
    public void setCountry(String value) { mCountry = value; }
    public void setLocalAddress(String value) { mLocalAddress = value; }
}
//ExEnd:Client
