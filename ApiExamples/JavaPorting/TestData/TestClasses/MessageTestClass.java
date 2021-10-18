package ApiExamples.TestData.TestClasses;

// ********* THIS FILE IS AUTO PORTED *********


public class MessageTestClass
{
    public String getName() { return mName; }; public void setName(String value) { mName = value; };

    private String mName;
    public String getMessage() { return mMessage; }; public void setMessage(String value) { mMessage = value; };

    private String mMessage;

    public MessageTestClass(String name, String message)
    {
        setName(name);
        setMessage(message);
    }
}
