package DocsExamples.LINQ_Reporting_Engine.Helpers.Data_Source_Objects;

import com.aspose.email.system.DateTime;

import java.time.LocalDate;

//ExStart:Contract
public class Contract
{
    private Manager mManager;
    private Client mClient;
    private float mPrice;
    private LocalDate mDate;

    public Manager getManager() { return mManager; }
    public Client getClient() { return mClient; }
    public float getPrice() { return mPrice; }
    public LocalDate getDate() { return mDate; }

    public void setManager(Manager value) { mManager = value; }
    public void setClient(Client value) { mClient = value; }
    public void setPrice(float value) { mPrice = value; }
    public void setDate(LocalDate value) { mDate = value; }
}
//ExEnd:Contract 
