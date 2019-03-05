package tgtools.excel.listener.event;

import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:06
 */
public class CreateWorkbookEvent extends Event {
    private Object mWorkbook;
    private Object mData;

    public Object getWorkbook() {
        return mWorkbook;
    }

    public void setWorkbook(Object pWorkbook) {
        mWorkbook = pWorkbook;
    }

    public Object getData() {
        return mData;
    }

    public void setData(Object pData) {
        mData = pData;
    }
}
