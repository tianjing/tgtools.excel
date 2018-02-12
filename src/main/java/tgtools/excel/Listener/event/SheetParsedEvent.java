package tgtools.excel.listener.event;

import com.fasterxml.jackson.databind.node.ArrayNode;
import tgtools.interfaces.Event;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：10:11
 */
public class SheetParsedEvent extends Event {
    private ArrayNode mData;
    private String mSheetName;

    public ArrayNode getData() {
        return mData;
    }

    public void setData(ArrayNode pData) {
        mData = pData;
    }

    public String getSheetName() {
        return mSheetName;
    }

    public void setSheetName(String pSheetName) {
        mSheetName = pSheetName;
    }
}
