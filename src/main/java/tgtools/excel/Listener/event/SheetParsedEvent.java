package tgtools.excel.Listener.event;

import tgtools.data.DataTable;
import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：10:11
 */
public class SheetParsedEvent extends Event {
private DataTable m_Table;
    private String m_SheetName;

    public DataTable getTable() {
        return m_Table;
    }

    public void setTable(DataTable p_Table) {
        m_Table = p_Table;
    }

    public String getSheetName() {
        return m_SheetName;
    }

    public void setSheetName(String p_SheetName) {
        m_SheetName = p_SheetName;
    }
}
