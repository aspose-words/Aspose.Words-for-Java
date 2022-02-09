package DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

// ********* THIS FILE IS AUTO PORTED *********



//ExStart:Manager
public class Manager
{
    public String getName() { return mName; }; public void setName(String value) { mName = value; };

    private String mName;
    public int getAge() { return mAge; }; public void setAge(int value) { mAge = value; };

    private int mAge;
    public byte[] getPhoto() { return mPhoto; }; public void setPhoto(byte[] value) { mPhoto = value; };

    private byte[] mPhoto;
    public Iterable<Contract> getContracts() { return mContracts; }; public void setContracts(Iterable<Contract> value) { mContracts = value; };

    private Iterable<Contract> mContracts;
}
//ExEnd:Manager
