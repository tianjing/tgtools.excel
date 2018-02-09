package tgtools.excel.poi;


import org.apache.poi.hssf.record.CellValueRecordInterface;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import tgtools.data.DataColumn;
import tgtools.data.DataRow;
import tgtools.data.DataTable;
import tgtools.excel.ImportExcel;
import tgtools.excel.Listener.ExcelAllLitener;
import tgtools.excel.Listener.ImportLisener;
import tgtools.excel.Listener.event.*;
import tgtools.exceptions.APPErrorException;
import tgtools.util.LogHelper;
import tgtools.util.StringUtil;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.sql.Types;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

/**
 * excel标准导入类
 * Created by tian_ on 2016-06-20.
 */
public class PoiImportExcel implements ImportExcel {

    protected int m_beginRow = 0;
    protected int m_beginTitleRow = 0;
    protected String m_tabName = "";//表名

    protected String[] m_attrName;//列中文名
    protected String[] m_attr;//列名


    protected ImportResult m_ImportResult;
    protected Workbook m_Workbook;

    protected ImportLisener m_Lisener;
    protected ExcelAllLitener m_AllLitener;

    public static void main(String[] args) throws APPErrorException, IOException {

        //FloatingDecimal.readJavaFormatString()
        tgtools.db.DataBaseFactory.add("DM", new Object[]{"jdbc:dm://192.168.88.128:5235/dqmis", "BQ_SYS", "BQ_SYS123"});

        String[] attrs = new String[]{"ID", "NAME", "BIR", "MONY1", "MONY2", "MONY3", "MONY4", "MONY5", "MONY6", "MONY7", "MONY8"};
        String[] attrnames = new String[]{"电压等级", "责任单位", "停电场所", "工作内容", "计划时间", "停电开始时间", "停电结束时间", "工作来源", "上次检修日期", "上次关联检修申请", "备注"};
        String tablename = "BQ_SYS.IMPORTTEST";
        int beginrow = 1;
        PoiImportExcel importex = new PoiImportExcel();
        importex.init(attrs, attrnames, tablename, 0, beginrow);

        importex.importExcel(new File("C:\\tianjing\\Desktop\\11.xlsx"));
        importex.close();
    }

    /**
     * 获取监听对象
     *
     * @return
     */
    public ImportLisener getLisener() {
        return m_Lisener;
    }

    /**
     * 设置监听对象
     *
     * @param p_Lisener
     */
    @Override
    public void setLisener(ImportLisener p_Lisener) {
        m_Lisener = p_Lisener;
    }

    @Override
    public void setExcelAllLisener(ExcelAllLitener p_Litener) {
        m_AllLitener = p_Litener;
    }

    /**
     * 获取导入的结果
     *
     * @return
     */
    @Override
    public String getImportResult() {
        if (null != m_ImportResult) {
            return m_ImportResult.toString();
        }
        return StringUtil.EMPTY_STRING;
    }

    @Override
    public void init(String[] p_Column, String[] p_ColumnName, String p_TableName) {
        m_attr = p_Column;
        m_attrName = p_ColumnName;
        m_tabName = p_TableName;
    }

    /**
     * @param p_Column     表列名
     * @param p_ColumnName 列中文名
     * @param p_TableName  表名
     * @param p_beginRow   起始行
     */
    @Override
    public void init(String[] p_Column, String[] p_ColumnName, String p_TableName, int p_BeginTitleRow, int p_beginRow) {
        m_attr = p_Column;
        m_attrName = p_ColumnName;
        m_beginRow = p_beginRow == 0 ? p_beginRow + 1 : p_beginRow;
        m_tabName = p_TableName;
        m_beginTitleRow = p_BeginTitleRow;
    }

    @Override
    public void init(String[] p_Column, String[] p_ColumnName, String p_TableName, int p_beginRow) {
        init(p_Column, p_ColumnName, p_TableName, p_beginRow > 0 ? p_beginRow - 1 : p_beginRow, p_beginRow);
    }

    /**
     * 导入excel
     *
     * @param p_ExcelFile excel 文件对象
     *
     * @throws APPErrorException
     */
    @Override
    public void importExcel(File p_ExcelFile) throws APPErrorException {
        if (null == p_ExcelFile || !p_ExcelFile.exists()) {
            throw new APPErrorException("excel 文件不存在");
        }
        try {
            m_Workbook = WorkbookFactory.createWorkbook(p_ExcelFile);

            CreateWorkbookEvent event = new CreateWorkbookEvent();
            event.setData(p_ExcelFile);
            event.setWorkbook(m_Workbook);
            onCreateWorkbook(event);
        } catch (Exception ex) {
            throw new APPErrorException("创建excel错误", ex);
        }
        try {
            doImportExcel();
        } catch (Exception ex) {
            throw new APPErrorException("导入excel错误；原因：" + ex.getMessage(), ex);
        }
    }

    /**
     * 导入excel
     *
     * @param excel excel 文件二进制数据
     *
     * @throws APPErrorException
     */
    @Override
    public void importExcel(byte[] excel) throws APPErrorException {
        try {
            m_Workbook = WorkbookFactory.createWorkbook(excel);

            CreateWorkbookEvent event = new CreateWorkbookEvent();
            event.setData(excel);
            event.setWorkbook(m_Workbook);
            onCreateWorkbook(event);
        } catch (Exception ex) {
            throw new APPErrorException("创建excel错误", ex);
        }

        try {
            doImportExcel();
        } catch (Exception ex) {
            throw new APPErrorException("导入excel错误；原因：" + ex.getMessage(), ex);
        }
    }

    protected void doImportExcel() throws APPErrorException {
        if (m_beginRow < 0) {
            throw new APPErrorException("数据行不能小于0");
        }
        if (m_beginTitleRow < 0) {
            throw new APPErrorException("标题行不能小于0");
        }
        if (m_beginRow <= m_beginTitleRow) {
            throw new APPErrorException("数据行只能大于标题行；标题行：" + String.valueOf(m_beginTitleRow) + ";数据行：" + String.valueOf(m_beginRow));
        }
        if (null != m_attr && m_attr.length > 0 && null != m_attrName && m_attrName.length > 0
                && m_beginRow > -1 && !StringUtil.isNullOrEmpty(m_tabName)) {
            parseExcel();
        }
    }

    public void onSheetParsed(SheetParsedEvent p_Event) {
        if (null != m_Lisener) {
            m_Lisener.onSheetParsed(p_Event);
        }
    }

    public void onReadSheet(ReadSheetEvent p_Event) {
        if (null != m_Lisener) {
            m_Lisener.onReadSheet(p_Event);
        }
    }

    protected void onExcuteSQL(ImportEvent p_Event) {
        if (null != m_Lisener) {
            m_Lisener.onExcuteSQL(p_Event);
        }
    }

    protected void onGetValue(ImportEvent p_Event) {
        if (null != m_Lisener) {
            m_Lisener.onGetValue(p_Event);
        }
    }

    /**
     * 创建excel workbook后对workbook的事件
     *
     * @param p_Event
     */
    protected void onCreateWorkbook(CreateWorkbookEvent p_Event) {
        if (null != m_AllLitener) {
            m_AllLitener.onCreateWorkbook(p_Event);
        }
    }

    /**
     * 创建DataTabel表格之前
     *
     * @param p_Event
     */
    protected void onCreateDataTable(CreateDataTableEvent p_Event) {
        if (null != m_AllLitener) {
            m_AllLitener.onCreateDataTable(p_Event);
        }
    }

    /**
     * 整个任务完成后事件
     *
     * @param p_Event
     */
    protected void onCompleted(ExcelCompletedEvent p_Event) {
        if (null != m_AllLitener) {
            m_AllLitener.onCompleted(p_Event);
        }
    }

    /**
     * 根据表格列信息创建表格
     *
     * @return
     */
    protected DataTable createDataTable() {
        DataTable dt = new DataTable();

        for (int i = 0; i < m_attr.length; i++) {
            DataColumn temp = dt.appendColumn(m_attr[i]);
            temp.setColumnType(Types.VARCHAR);
        }
        return dt;
    }

    /**
     * 解析excel
     *
     * @return
     *
     * @throws APPErrorException
     */
    public void parseExcel() throws APPErrorException {
        m_ImportResult = new ImportResult();
        List<DataTable> tables = null;
        try {
            tables = parseDataTable();
        } catch (Exception ex) {
            LogHelper.error("", "excel解析错误22", "ImportExcel。parseExcel1", ex);
        }
        try {
            if (null == tables || tables.size() < 1) {
                throw new APPErrorException("excel没有有效内容");
            }
            int count = 0;
            int sucess = 0;
            for (int i = 0; i < tables.size(); i++) {
                DataTable table = tables.get(i);
                count += table.getRowCount();
                String tablename = table.getTableName();
                for (int rownum = 0; rownum < table.getRowCount(); rownum++) {
                    DataRow row = table.getRow(rownum);
                    try {
                        ImportEvent event = new ImportEvent();
                        event.setAttrName(m_attrName);
                        event.setAttr(m_attr);
                        event.setIsExcute(true);
                        event.setRow(row);

                        onExcuteSQL(event);

                        String sql = tgtools.util.DataTableSqlFactory.buildInsertSql(row, tablename);
                        if (event.isExcute()) {
                            sucess += execute(sql);
                        } else {
                            sucess += event.isSucess() ? 1 : 0;
                        }

                    } catch (Exception e) {
                        LogHelper.error("", "excel导入出错", "ImportExcel。parseExcel1", e);
                    }
                }
            }
            m_ImportResult.setCount(count);
            m_ImportResult.setSucess(sucess);
            ExcelCompletedEvent event = new ExcelCompletedEvent();
            event.setTables(tables);
            event.setWorkbook(m_Workbook);
            onCompleted(event);
        } catch (Exception ex) {
            throw new APPErrorException("解析excel错误，原因：" + ex.getMessage(), ex);
        }
    }

    /**
     * 解析数据集
     *
     * @return
     *
     * @throws APPErrorException
     */
    private List<DataTable> parseDataTable() throws APPErrorException {
        List<DataTable> res = new ArrayList<DataTable>();
        int sheetcount = m_Workbook.getNumberOfSheets();
        for (int i = 0; i < sheetcount; i++) {//循环excel
            Sheet sheet = m_Workbook.getSheetAt(i);
            if (null == sheet) {
                throw new APPErrorException("获取第" + i + "个表错误");
            }
            int titlerow = 0;
            int datarow = 0;
            int rowcount = 0;
            Row titleRow = null;
            DataTable dt = null;
            try {
                ReadSheetEvent readevent = new ReadSheetEvent();
                readevent.setCancel(false);
                readevent.setCancel(false);
                readevent.setbeginDataRow(m_beginRow);
                readevent.setbeginTitleRow(m_beginTitleRow);
                readevent.setSheetName(sheet.getSheetName());
                onReadSheet(readevent);
                if (readevent.getCancel()) {
                    continue;
                }


                CreateDataTableEvent tableevent = new CreateDataTableEvent();
                tableevent.setColumnNames(m_attrName);
                tableevent.setColumns(m_attr);
                onCreateDataTable(tableevent);
                if (tableevent.getIsExcute()) {
                    m_attrName = tableevent.getColumnNames();
                    m_attr = tableevent.getColumns();
                    try {
                        dt = createDataTable();
                    } catch (Exception ex) {
                        throw new APPErrorException("解析第" + i + "表错误；起始行：" + m_beginRow + "；标题行：" + m_beginTitleRow);
                    }
                } else {
                    dt = tableevent.getTable();
                }

                titlerow = readevent.getbeginTitleRow();
                datarow = readevent.getbeginDataRow();
                dt.setTableName(m_tabName);

                rowcount = sheet.getLastRowNum() + 1;//获取总行数（由于获取行号从0开始，所以行数应为行号加1）
                LogHelper.info("", "正在解析总行数：" + rowcount, "");
                titleRow = sheet.getRow(titlerow);
            } catch (Exception ex) {
                throw new APPErrorException("解析第" + i + "表错误；起始行：" + m_beginRow + "；标题行：" + m_beginTitleRow);
            }
            for (int r = datarow; r < rowcount; r++) {//循环添加行
                try {
                    Row sheetrow = sheet.getRow(r);
                    if (null == sheetrow) {
                        continue;
                    }
                    LogHelper.info("", "正在解析行：" + r, "");
                    int colcount = sheetrow.getPhysicalNumberOfCells();
                    LogHelper.info("", "正在解析总列数：" + colcount, "");
                    DataRow row = dt.appendRow();
                    PlatformEngine.getCommonListBll().initNewRow(row);
                    for (int c = 0; c < colcount; c++)//循环添加列值
                    {
                        if (null == titleRow.getCell(c)) {
                            throw new APPErrorException("无效的标题cell：" + c + "sheetname:" + sheet.getSheetName());
                        }
                        LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + titleRow.getCell(c).getStringCellValue(), "");
                        String colName = titleRow.getCell(c).getStringCellValue();//取得excel列名
                        String col = getAttrName(colName);

                        if (StringUtil.isNullOrEmpty(col)) {
                            continue;
                        }
                        String value = "";
                        try {
                            value = getCellValue(sheet.getRow(r).getCell(c));
                            LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + titleRow.getCell(c).getStringCellValue() + "；值：" + value, "");
                        } catch (Exception ex) {
                            throw new APPErrorException("取值错误：" + ex.getMessage() + ";行：" + String.valueOf(r) + "；列：" + String.valueOf(c) + ";sheet:" + sheet.toString() + ";row:" + sheet.getRow(r).toString(), ex);
                        }
                        try {
                            ImportEvent event = new ImportEvent();
                            event.setAttrName(m_attrName);
                            event.setAttr(m_attr);
                            event.setRowIndex(r);
                            event.setColumnIndex(c);
                            event.setValue(value);

                            onGetValue(event);
                            value = event.getValue();

                            row.setValue(col, value);
                        } catch (Exception ex) {
                            throw new APPErrorException("添加数据到datatable出错：" + ex.getMessage() + ";行：" + String.valueOf(r) + "；列：" + String.valueOf(c) + ";sheet:" + sheet.toString() + ";row:" + sheet.getRow(r).toString(), ex);
                        }
                    }
                } catch (Exception ex) {
                    throw new APPErrorException("添加数据到datatable出错：" + ex.getMessage() + ";行：" + String.valueOf(r) + ";sheet:" + sheet.toString(), ex);
                }
            }
            LogHelper.info("", "ppppppppppppppppppp", "");

            SheetParsedEvent sheetevent = new SheetParsedEvent();
            sheetevent.setSheetName(sheet.getSheetName());
            sheetevent.setTable(dt);
            onSheetParsed(sheetevent);
            if (null != dt && dt.getRowCount() > 0) {
                res.add(dt);
            }
        }
        return res;
    }

    /**
     * 根据excel列名获取表列明
     *
     * @param p_ExcelColumnName
     *
     * @return
     */
    private String getAttrName(String p_ExcelColumnName) {
        for (int i = 0; i < m_attrName.length; i++) {
            if (m_attrName[i].equals(p_ExcelColumnName)) {
                return m_attr[i];
            }
        }
        return null;
    }

    /**
     * 获取单元格值
     *
     * @param p_Cell
     *
     * @return
     */
    @SuppressWarnings("deprecation")
    private String getCellValue(Cell p_Cell) {
        String cellValue = "";
        if (null == p_Cell) {
            return StringUtil.EMPTY_STRING;
        }
        switch (p_Cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                cellValue = p_Cell.getRichStringCellValue().getString().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(p_Cell)) {
                    cellValue = getDateCellValue(DateUtil.getJavaDate(p_Cell.getNumericCellValue()));
                    break;
                }
                cellValue = String.valueOf(p_Cell.getNumericCellValue());
                if (!StringUtil.isNullOrEmpty(cellValue) && (cellValue.indexOf("E+") > 0 || cellValue.indexOf("E-") > 0)) {
                    cellValue = getFixNumericValue(p_Cell, cellValue);
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue = String.valueOf(p_Cell.getBooleanCellValue()).trim();
                break;
            case Cell.CELL_TYPE_FORMULA:
                //cellValue = p_Cell.getCellFormula();
                cellValue = getFORMULACellValue(p_Cell);
                break;
            default:
                cellValue = "";
        }
        return cellValue;

    }

    /**
     * 用于xls 2003 数据变成科学计数后的问题
     *
     * @param p_Cell
     * @param p_DefaultValue
     *
     * @return
     */
    private String getFixNumericValue(Cell p_Cell, String p_DefaultValue) {
        String result = p_DefaultValue;
        try {
            //excel 2003
            if (p_Cell instanceof HSSFCell) {
                Method method = p_Cell.getClass().getDeclaredMethod("getCellValueRecord");
                if (null != method) {
                    method.setAccessible(true);
                    CellValueRecordInterface obj = (CellValueRecordInterface) method.invoke(p_Cell);
                    String value = obj.toString();
                    value = value.substring(value.indexOf(".value=") + 7);
                    result = value.substring(0, value.indexOf("\n"));
                }
            } else {
                //excel 2007
                BigDecimal bigDecimal = new BigDecimal(result);
                String str = bigDecimal.toPlainString();
                String pattern = "(\\d.*[1-9])";
                String res = tgtools.util.RegexHelper.regexFirst(str, pattern);
                try {
                    new BigDecimal(res);
                    result = res;
                } catch (Exception ex) {
                    LogHelper.error("系统", "转换科学基数出错，原值："+p_DefaultValue+";;转换后："+res, "ImportExcel.getFixNumericValue", ex);
                }
            }
        } catch (Exception e) {
            LogHelper.error("系统", "转换科学基数出错，不支持的excel类型，请尝试excel2003", "ImportExcel.getFixNumericValue", e);
        }
        return result;
    }

    /**
     * 获取公式的值
     *
     * @param p_Cell
     *
     * @return
     */
    @SuppressWarnings("deprecation")
    private String getFORMULACellValue(Cell p_Cell) {
        String cellValue = "";
        FormulaEvaluator evaluator = p_Cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        CellValue evalValue = evaluator.evaluate(p_Cell);
        switch (evalValue.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                cellValue = evalValue.getStringValue().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                cellValue = String.valueOf(evalValue.getNumberValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cellValue = String.valueOf(evalValue.getBooleanValue());
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }

    /**
     * 获取日期型单元格值
     *
     * @param p_Date
     *
     * @return
     */
    private String getDateCellValue(Date p_Date) {
        Date date = p_Date;
        TimeZone zone = TimeZone.getTimeZone("Asia/Shanghai");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        sdf.setTimeZone(zone);
        return sdf.format(date);
    }

    @Override
    public void close() {

        if (null != m_Workbook) {
            try {
                m_Workbook.close();
            } catch (Exception e) {
                LogHelper.error("", "关闭时出错；原因：" + e.getMessage(), "ImportExcel.Dispose", e);
            }
        }

        m_Workbook = null;
    }

    private int execute(String sql) throws APPErrorException {
        sql = tgtools.util.SqlStrHelper.processKeyWord(sql);
        return tgtools.db.DataBaseFactory.getDefault().executeUpdate(sql);
    }



    public class ImportResult {
        private String m_Name;
        private int m_Count = 0;
        private int m_Sucess = 0;

        public String getName() {
            return m_Name;
        }

        public void setName(String p_Name) {
            m_Name = p_Name;
        }

        public int getCount() {
            return m_Count;
        }

        public void setCount(int p_Count) {
            m_Count = p_Count;
        }

        public int getSucess() {
            return m_Sucess;
        }

        public void setSucess(int p_Sucess) {
            m_Sucess = p_Sucess;
        }

        @Override
        public String toString() {
            return m_Sucess + "/" + m_Count;
        }
    }
}
