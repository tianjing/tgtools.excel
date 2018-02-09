package tgtools.excel.Listener.event;

import tgtools.data.DataRow;
import tgtools.interfaces.Event;

import java.util.List;

/**
 * 导入时的事件参数
 * Created by tian_ on 2016-06-20.
 */
public class ImportEvent extends Event {

    private String[] m_AttrName;
    private String[] m_Attr;
    private DataRow m_Row;
    private String m_Value;
    private int m_RowIndex;
    private int m_ColumnIndex;
    private boolean m_IsExcute=true;
    private String m_Filter;
    private List<String> m_FilterNames;
    private List<String> m_FilterValues;
    private boolean m_IsSucess=false;

    public boolean isSucess() {
        return m_IsSucess;
    }

    public void setIsSucess(boolean p_IsSucess) {
        m_IsSucess = p_IsSucess;
    }

    public String getFilter() {
        return m_Filter;
    }

    public void setFilter(String p_Filter) {
        m_Filter = p_Filter;
    }

    public List<String> getFilterNames() {
        return m_FilterNames;
    }

    public void setFilterNames(List<String> p_FilterNames) {
        m_FilterNames = p_FilterNames;
    }

    public List<String> getFilterValues() {
        return m_FilterValues;
    }

    public void setFilterValues(List<String> p_FilterValues) {
        m_FilterValues = p_FilterValues;
    }

    public DataRow getRow() {
        return m_Row;
    }

    public void setRow(DataRow p_Row) {
        m_Row = p_Row;
    }

    public boolean isExcute() {
        return m_IsExcute;
    }

    public void setIsExcute(boolean p_IsExcute) {
        m_IsExcute = p_IsExcute;
    }

    public String getValue() {
        return m_Value;
    }

    public void setValue(String p_Value) {
        m_Value = p_Value;
    }

    public int getColumnIndex() {
        return m_ColumnIndex;
    }

    public void setColumnIndex(int p_ColumnIndex) {
        m_ColumnIndex = p_ColumnIndex;
    }

    public int getRowIndex() {
        return m_RowIndex;
    }

    public void setRowIndex(int p_RowIndex) {
        m_RowIndex = p_RowIndex;
    }


    public String[] getAttr() {
        return m_Attr;
    }

    public void setAttr(String[] p_Attr) {
        this.m_Attr = p_Attr;
    }

    public String[] getAttrName() {
        return m_AttrName;
    }

    public void setAttrName(String[] p_AttrName) {
        this.m_AttrName = p_AttrName;
    }


}
