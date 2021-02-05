package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import java.time.LocalDate;

public class ContractTestClass {
    private ManagerTestClass mManager;
    private ClientTestClass mClient;
    private float mPrice;
    private LocalDate mDate;

    public ContractTestClass() {
    }

    public ContractTestClass(final ManagerTestClass manager, final ClientTestClass client, final float price, final LocalDate date) {
        setManager(manager);
        setClient(client);
        setPrice(price);
        setDate(date);
    }

    public void setManager(final ManagerTestClass value) {
        mManager = value;
    }

    public void setClient(final ClientTestClass value) {
        mClient = value;
    }

    public void setPrice(final float value) {
        mPrice = value;
    }

    public void setDate(final LocalDate value) {
        mDate = value;
    }

    public ManagerTestClass getManager() {
        return mManager;
    }

    public ClientTestClass getClient() {
        return mClient;
    }

    public float getPrice() {
        return mPrice;
    }

    public LocalDate getDate() {
        return mDate;
    }


}
