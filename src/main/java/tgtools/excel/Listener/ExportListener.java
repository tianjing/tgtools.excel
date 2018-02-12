package tgtools.excel.listener;


import tgtools.excel.listener.event.CreateWorkbookEvent;
import tgtools.excel.listener.event.ExcelCompletedEvent;
import tgtools.excel.listener.event.ExportExcelEvent;

/**
 * Created by tian_ on 2016-06-17.
 */
public interface ExportListener {
    /**
     * 当写入Excel标题时事件 通过cellindex datatable 和 value（jxl.write.Lable）对写入Excel的值做自定义处理
     * @param pEvent
     */
    void onWriteTitle(ExportExcelEvent pEvent);

    /**
     * 当写入Excel内容时事件 通过cellindex rowindex datatable 和 value（jxl.write.Lable）对写入Excel的值做自定义处理
     * @param pEvent
     */
    void onWriteCell(ExportExcelEvent pEvent);
    /**
     * 创建excel workbook后对workbook的事件
     * @param pEvent
     */
    void onCreateWorkbook(CreateWorkbookEvent pEvent);

    /**
     * 整个任务完成后事件
     * @param pEvent
     */
    void onCompleted(ExcelCompletedEvent pEvent);


}
