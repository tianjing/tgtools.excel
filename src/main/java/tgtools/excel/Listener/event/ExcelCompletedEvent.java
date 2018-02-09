package tgtools.excel.Listener.event;

import tgtools.data.DataTable;
import tgtools.interfaces.Event;

import java.util.List;

/**
 * 名  称：表示导入或导出结果事件参数，如果导入那么Table有数据，导出workbook有数据
 * 编写者：田径
 * 功  能：
 * 时  间：10:24
 */
public class ExcelCompletedEvent extends Event {
    private DataTable m_Table;
    private List<DataTable> m_Tables;
    private Object Workbook;

    public DataTable getTable() {
        return m_Table;
    }

    public void setTable(DataTable p_Table) {
        m_Table = p_Table;
    }

    public List<DataTable> getTables() {
        return m_Tables;
    }

    public void setTables(List<DataTable> p_Tables) {
        m_Tables = p_Tables;
    }

    public Object getWorkbook() {
        return Workbook;
    }

    public void setWorkbook(Object workbook) {
        Workbook = workbook;
    }
}
