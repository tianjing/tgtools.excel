package tgtools.excel.listener;


import tgtools.excel.listener.event.*;

/**
 * Created by tian_ on 2016-06-20.
 */
public interface ImportListener {
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

    /**
     * 加载过滤条件事件 通过 列名 列中文名 过滤条件 解析后的过滤条件名集合 过滤条件值集合，重新处理得到新的过滤条件名集合 过滤条件值集合
     * @param pEvent
     */
    void onLoadFilter(ImportEvent pEvent);
    /**
     * 获取列事件 获取平台表格展示列的中英文（PlatformImportExcel 用到）
     * @param pEvent
     */
    void onGetAtted(ImportEvent pEvent);

    /**
     * 获取值事件 通过列名 列中文名 rowindex ColumnIndex 和value （String） 自定义修改值
     * @param pEvent
     */
    void onGetValue(ImportEvent pEvent);

    /**
     * 执行sql事件  通过 列名 列中文名 row行数据 isExcute（true执行默认sql，false：不执行默认sql）调整sql的执行
     * @param pEvent
     */
    void onExcuteSQL(ImportEvent pEvent);

    /**
     * 读取sheet事件  通过 sheetname Cancel 决定是跳过当前的sheet
     * @param pEvent
     */
    void onReadSheet(ReadSheetEvent pEvent);

    /**
     * 一个sheet解析后事件 通过sheetname 和 DataTable 设置DataTable 一般需要设置DataTable的Name为表名
     * @param pEvent
     */
    void onSheetParsed(SheetParsedEvent pEvent);


}

