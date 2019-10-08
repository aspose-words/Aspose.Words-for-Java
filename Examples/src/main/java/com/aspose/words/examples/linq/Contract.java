package com.aspose.words.examples.linq;

import java.util.Date;

//ExStart:Contract
public class Contract {
    private Manager Manager;

    public final Manager getManager() {
        return Manager;
    }

    public final void setManager(Manager value) {
        Manager = value;
    }

    private Client Client;

    public final Client getClient() {
        return Client;
    }

    public final void setClient(Client value) {
        Client = value;
    }

    private float Price;

    public final float getPrice() {
        return Price;
    }

    public final void setPrice(float value) {
        Price = value;
    }

    private Date Date = new Date();

    public final Date getDate() {
        return Date;
    }

    public final void setDate(Date value) {
        Date = value;
    }
}
//ExEnd:Contract
