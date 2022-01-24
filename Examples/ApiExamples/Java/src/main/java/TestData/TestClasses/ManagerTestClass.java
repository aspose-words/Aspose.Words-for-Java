package TestData.TestClasses;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import java.util.ArrayList;

public class ManagerTestClass {
    private String mName;
    private int mAge;
    private ArrayList<ContractTestClass> mContracts;

    public ManagerTestClass() {
    }

    public ManagerTestClass(final String name, final int age, final ArrayList<ContractTestClass> contracts) {
        setName(name);
        setAge(age);
        setContracts(contracts);
    }

    public void setName(final String value) {
        mName = value;
    }

    public void setAge(final int value) {
        mAge = value;
    }

    public void setContracts(final ArrayList<ContractTestClass> value) {
        mContracts = value;
    }

    public String getName() {
        return mName;
    }

    public int getAge() {
        return mAge;
    }

    public ArrayList<ContractTestClass> getContracts() {
        return mContracts;
    }

    public int getContractsSum() {
        int contractsSum = 0;

        for (ContractTestClass contract : getContracts()) {
            contractsSum = (int) (contractsSum + contract.getPrice());
        }

        return contractsSum;
    }
}
