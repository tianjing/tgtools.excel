package tgtools.excel.Listener.event;

import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:06
 */
public class CreateWorkbookEvent extends Event {
    private Object m_Workbook;
    private Object m_Data;

    public Object getWorkbook() {
        return m_Workbook;
    }

    public void setWorkbook(Object p_Workbook) {
        m_Workbook = p_Workbook;
    }

    public Object getData() {
        return m_Data;
    }

    public void setData(Object p_Data) {
        m_Data = p_Data;
    }
}
