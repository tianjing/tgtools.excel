package tgtools.excel.listener.event;

import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：10:06
 */
public class ReadSheetEvent extends Event {
    private String mSheetName;
    private boolean mCancel=false;
    protected int mBeginDataRow=1;
    protected int mBeginTitleRow=0;
    private String mTableName="";

    public String getSheetName() {
        return mSheetName;
    }

    public void setSheetName(String pSheetName) {
        mSheetName = pSheetName;
    }

    public boolean getCancel() {
        return mCancel;
    }

    public void setCancel(boolean pCancel) {
        mCancel = pCancel;
    }

    public int getBeginDataRow() {
        return mBeginDataRow;
    }

    public void setBeginDataRow(int pBeginDataRow) {
        mBeginDataRow = pBeginDataRow;
    }

    public int getBeginTitleRow() {
        return mBeginTitleRow;
    }

    public void setBeginTitleRow(int pBeginTitleRow) {
        mBeginTitleRow = pBeginTitleRow;
    }

    public String getTableName() {
        return mTableName;
    }

    public void setTableName(String pTableName) {
        mTableName = pTableName;
    }
}
