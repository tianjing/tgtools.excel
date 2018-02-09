package tgtools.excel.Listener.event;

import tgtools.data.DataTable;
import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:08
 */
public class CreateDataTableEvent extends Event {
    private DataTable m_Table;
    private String[] m_Columns;
    private String[] m_ColumnNames;
    private boolean m_IsExcute=true;

    public boolean getIsExcute() {
        return m_IsExcute;
    }

    public void setIsExcute(boolean p_IsExcute) {
        m_IsExcute = p_IsExcute;
    }

    public String[] getColumns() {
        return m_Columns;
    }

    public void setColumns(String[] p_Columns) {
        m_Columns = p_Columns;
    }

    public String[] getColumnNames() {
        return m_ColumnNames;
    }

    public void setColumnNames(String[] p_ColumnNames) {
        m_ColumnNames = p_ColumnNames;
    }

    public DataTable getTable() {
        return m_Table;
    }

    public void setTable(DataTable p_Table) {
        m_Table = p_Table;
    }
}
