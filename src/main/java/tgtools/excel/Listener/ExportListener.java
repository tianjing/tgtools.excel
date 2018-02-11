package tgtools.excel.Listener;


import tgtools.excel.Listener.event.CreateWorkbookEvent;
import tgtools.excel.Listener.event.ExcelCompletedEvent;
import tgtools.excel.Listener.event.ExportExcelEvent;

/**
 * Created by tian_ on 2016-06-17.
 */
public interface ExportListener {
    /**
     * 当写入Excel标题时事件 通过cellindex datatable 和 value（jxl.write.Lable）对写入Excel的值做自定义处理
     * @param p_ExportExcelEvent
     */
    void onWriteTitle(ExportExcelEvent p_ExportExcelEvent);

    /**
     * 当写入Excel内容时事件 通过cellindex rowindex datatable 和 value（jxl.write.Lable）对写入Excel的值做自定义处理
     * @param p_ExportExcelEvent
     */
    void onWriteCell(ExportExcelEvent p_ExportExcelEvent);
    /**
     * 创建excel workbook后对workbook的事件
     * @param p_Event
     */
    void onCreateWorkbook(CreateWorkbookEvent p_Event);

    /**
     * 整个任务完成后事件
     * @param p_Event
     */
    void onCompleted(ExcelCompletedEvent p_Event);


}
