package tgtools.excel.Listener;


import tgtools.excel.Listener.event.ExportExcelEvent;

/**
 * Created by tian_ on 2016-06-17.
 */
public interface ExportLisener {
    /**
     * 获取平台表格展示列的中英文（PlatformImportExcel 用到）
     * @param p_ColumnName 列的中文名
     * @param p_Column 列名
     */
    void onGetAtted(String[] p_ColumnName, String[] p_Column);

    /**
     * 获取表格数据事件 这里可以对DataTable做调整或者重新set新的DataTable
     * @param p_ExportExcelEvent
     */
    void onGetDataed(ExportExcelEvent p_ExportExcelEvent);

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

}
