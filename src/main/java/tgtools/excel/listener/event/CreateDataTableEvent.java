package tgtools.excel.listener.event;

import tgtools.data.DataTable;
import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:08
 */
public class CreateDataTableEvent extends Event {
    private DataTable mTable;
    private String[] mColumns;
    private String[] mColumnNames;
    private boolean mIsExcute=true;

    public boolean getIsExcute() {
        return mIsExcute;
    }

    public void setIsExcute(boolean pIsExcute) {
        mIsExcute = pIsExcute;
    }

    public String[] getColumns() {
        return mColumns;
    }

    public void setColumns(String[] pColumns) {
        mColumns = pColumns;
    }

    public String[] getColumnNames() {
        return mColumnNames;
    }

    public void setColumnNames(String[] pColumnNames) {
        mColumnNames = pColumnNames;
    }

    public DataTable getTable() {
        return mTable;
    }

    public void setTable(DataTable pTable) {
        mTable = pTable;
    }
}
