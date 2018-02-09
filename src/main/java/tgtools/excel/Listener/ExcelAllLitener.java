package tgtools.excel.Listener;


import tgtools.excel.Listener.event.CreateDataTableEvent;
import tgtools.excel.Listener.event.CreateWorkbookEvent;
import tgtools.excel.Listener.event.ExcelCompletedEvent;


/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:04
 */
public interface ExcelAllLitener {

    /**
     * 创建excel workbook后对workbook的事件
     * @param p_Event
     */
    void onCreateWorkbook(CreateWorkbookEvent p_Event);

    /**
     * 创建DataTabel表格之前
     * @param p_Event
     */
    void onCreateDataTable(CreateDataTableEvent p_Event);


    /**
     * 整个任务完成后事件
     * @param p_Event
     */
    void onCompleted(ExcelCompletedEvent p_Event);



}
