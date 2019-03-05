package tgtools.excel.listener.event;

import com.fasterxml.jackson.databind.JsonNode;
import tgtools.interfaces.Event;

import java.util.LinkedHashMap;
import java.util.List;

/**
 * 导入时的事件参数
 * Created by tian_ on 2016-06-20.
 */
public class ImportEvent extends Event {

    private LinkedHashMap<String,String> pColumns;
    private JsonNode mRow;
    private Object mValue;
    private int mRowIndex;
    private int mColumnIndex;
    private boolean mIsExcute=true;
    private String mFilter;
    private List<String> mFilterNames;
    private List<String> mFilterValues;
    private boolean mIsSucess=false;
    private String mSql=null;

    public JsonNode getRow() {
        return mRow;
    }

    public void setRow(JsonNode pRow) {
        mRow = pRow;
    }

    public Object getValue() {
        return mValue;
    }

    public void setValue(Object pValue) {
        mValue = pValue;
    }

    public int getRowIndex() {
        return mRowIndex;
    }

    public void setRowIndex(int pRowIndex) {
        mRowIndex = pRowIndex;
    }

    public int getColumnIndex() {
        return mColumnIndex;
    }

    public void setColumnIndex(int pColumnIndex) {
        mColumnIndex = pColumnIndex;
    }

    public boolean getIsExcute() {
        return mIsExcute;
    }

    public void setIsExcute(boolean pIsExcute) {
        mIsExcute = pIsExcute;
    }

    public String getFilter() {
        return mFilter;
    }

    public void setFilter(String pFilter) {
        mFilter = pFilter;
    }

    public List<String> getFilterNames() {
        return mFilterNames;
    }

    public void setFilterNames(List<String> pFilterNames) {
        mFilterNames = pFilterNames;
    }

    public List<String> getFilterValues() {
        return mFilterValues;
    }

    public void setFilterValues(List<String> pFilterValues) {
        mFilterValues = pFilterValues;
    }

    public boolean getIsSucess() {
        return mIsSucess;
    }

    public void setIsSucess(boolean pIsSucess) {
        mIsSucess = pIsSucess;
    }

    public LinkedHashMap<String, String> getpColumns() {
        return pColumns;
    }

    public void setpColumns(LinkedHashMap<String, String> pColumns) {
        this.pColumns = pColumns;
    }

    public String getSql() {
        return mSql;
    }

    public void setSql(String pSql) {
        mSql = pSql;
    }
}
