package tgtools.excel.Listener.event;

import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：10:06
 */
public class ReadSheetEvent extends Event {
    private String m_SheetName;
    private boolean m_Cancel=false;
    protected int m_beginDataRow=1;
    protected int m_beginTitleRow=0;
    private String m_TableName="";

    public String getTableName() {
        return m_TableName;
    }

    public void setTableName(String p_TableName) {
        m_TableName = p_TableName;
    }

    public int getbeginDataRow() {
        return m_beginDataRow;
    }

    public void setbeginDataRow(int p_beginDataRow) {
        m_beginDataRow = p_beginDataRow;
    }

    public int getbeginTitleRow() {
        return m_beginTitleRow;
    }

    public void setbeginTitleRow(int p_beginTitleRow) {
        m_beginTitleRow = p_beginTitleRow;
    }

    public String getSheetName() {
        return m_SheetName;
    }

    public void setSheetName(String p_SheetName) {
        m_SheetName = p_SheetName;
    }

    public boolean getCancel() {
        return m_Cancel;
    }

    public void setCancel(boolean p_Cancel) {
        m_Cancel = p_Cancel;
    }
}
