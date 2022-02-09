package DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

//ExStart:PointDataClass
public class PointData
{
    private String mTime;
    private int mFlow;
    private int mRainfall;

    public String getTime() { return mTime; }
    public int getFlow() { return mFlow; }
    public int getRainfall() { return mRainfall; }

    public void setTime(String value) { mTime = value; }
    public void setFlow(int value) { mFlow = value; }
    public void setRainfall(int value) { mRainfall = value; }
}
//ExEnd:PointDataClass

