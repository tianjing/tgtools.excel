package tgtools.excel.listener.event;

import com.fasterxml.jackson.databind.node.ArrayNode;
import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：16:45
 */
public class ExportExcelEvent extends Event {
    private ArrayNode mDatas;
    private int mRowIndex;
    private int mcellIndex;
    private Object mValue;
    private String mTableSql;
    public ArrayNode getDatas() {
        return mDatas;
    }

    public void setDatas(ArrayNode pDatas) {
        mDatas = pDatas;
    }

    public int getRowIndex() {
        return mRowIndex;
    }

    public void setRowIndex(int pRowIndex) {
        mRowIndex = pRowIndex;
    }

    public int getCellIndex() {
        return mcellIndex;
    }

    public void setCellIndex(int pcellIndex) {
        mcellIndex = pcellIndex;
    }

    public Object getValue() {
        return mValue;
    }

    public void setValue(Object pValue) {
        mValue = pValue;
    }

    public String getTableSql() {
        return mTableSql;
    }

    public void setTableSql(String pTableSql) {
        mTableSql = pTableSql;
    }
}
