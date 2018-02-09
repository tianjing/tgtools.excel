package tgtools.excel.Listener.event;

import tgtools.data.DataTable;
import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：16:45
 */
public class ExportExcelEvent extends Event {
    private DataTable m_Table;
    private int m_RowIndex;
    private int m_cellIndex;
    private Object m_Value;
    private String m_TableSql;
    public DataTable getTable() {
        return m_Table;
    }

    public void setTable(DataTable p_Table) {
        m_Table = p_Table;
    }

    public int getRowIndex() {
        return m_RowIndex;
    }

    public void setRowIndex(int p_RowIndex) {
        m_RowIndex = p_RowIndex;
    }

    public int getcellIndex() {
        return m_cellIndex;
    }

    public void setcellIndex(int p_cellIndex) {
        m_cellIndex = p_cellIndex;
    }

    public Object getValue() {
        return m_Value;
    }

    public void setValue(Object p_Value) {
        this.m_Value = p_Value;
    }

    public String getTableSql() {
        return m_TableSql;
    }

    public void setTableSql(String p_TableSql) {
        m_TableSql = p_TableSql;
    }
}
