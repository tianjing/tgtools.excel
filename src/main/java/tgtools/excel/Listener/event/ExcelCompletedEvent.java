package tgtools.excel.Listener.event;

import com.fasterxml.jackson.databind.node.ArrayNode;
import tgtools.interfaces.Event;

import java.util.List;

/**
 * 名  称：表示导入或导出结果事件参数，如果导入那么Table有数据，导出workbook有数据
 * 编写者：田径
 * 功  能：
 * 时  间：10:24
 */
public class ExcelCompletedEvent extends Event {
    private List<ArrayNode> mDatas;
    private Object mWorkbook;

    public List<ArrayNode> getDatas() {
        return mDatas;
    }

    public void setDatas(List<ArrayNode> pDatas) {
        mDatas = pDatas;
    }

    public Object getWorkbook() {
        return mWorkbook;
    }

    public void setWorkbook(Object pWorkbook) {
        mWorkbook = pWorkbook;
    }
}
