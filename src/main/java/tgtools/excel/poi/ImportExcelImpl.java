package tgtools.excel.poi;


import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.hssf.record.CellValueRecordInterface;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.ss.usermodel.*;
import tgtools.excel.ImportExcel;
import tgtools.excel.listener.ImportListener;
import tgtools.excel.listener.event.*;
import tgtools.excel.util.JsonUtil;
import tgtools.exceptions.APPErrorException;
import tgtools.util.FileUtil;
import tgtools.util.LogHelper;
import tgtools.util.StringUtil;

import java.io.*;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：21:47
 */
public class ImportExcelImpl implements ImportExcel {

    protected int mbeginDataRow = 0;
    protected int mbeginTitleRow = 0;
    protected LinkedHashMap<String, String> mCoulmns;
    protected ObjectNode mImportResult;
    protected Workbook mWorkbook;
    protected ImportListener mListener;
    protected Map<String, String> mSheetTable;
    protected ArrayList<String> mTableNames;
    protected String mDatabaseType = null;
    protected LinkedHashMap<String, ArrayNode> mParseData = null;

    @Override
    public void setListener(ImportListener pListener) {
        mListener = pListener;
    }

    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String, String> pSheetTable) {
        init(pColumns, pSheetTable, null, 0, 1);
    }

    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String, String> pSheetTable, String pDatabaseType) {
        init(pColumns, pSheetTable, pDatabaseType, 0, 1);
    }


    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String, String> pSheetTable, String pDatabaseType, int pBeginDataRow) {
        init(pColumns, pSheetTable, pDatabaseType, 0, pBeginDataRow);
    }

    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String, String> pSheetTable, String pDatabaseType, int pBeginTitleRow, int pBeginDataRow) {
        mCoulmns = pColumns;
        mDatabaseType = pDatabaseType;
        mbeginTitleRow = pBeginTitleRow;
        mbeginDataRow = pBeginDataRow;
        mSheetTable = pSheetTable;
    }

    @Override
    public void importExcel(File pFile) throws APPErrorException {

        if (null == pFile || !pFile.exists()) {
            throw new APPErrorException("文件不存在！p_File：" + (null == pFile ? "null" : pFile.getAbsolutePath()));
        }
        String ext = FileUtil.getExtensionName(pFile.getAbsolutePath());
        try {
            importExcel(new FileInputStream(pFile), ext);
        }catch (FileNotFoundException e)
        {
            throw new APPErrorException("文件不存在！p_File：" + (null == pFile ? "null" : pFile.getAbsolutePath()));
        }
    }

    @Override
    public void importExcel(byte[] pDatas, String pVersion) throws APPErrorException {
        if (null == pDatas || pDatas.length < 1) {
            throw new APPErrorException("无效的文件内容p_Datas");
        }
        importExcel(new ByteArrayInputStream(pDatas), pVersion);
    }

    @Override
    public void importExcel(InputStream pInputStream, String pVersion) throws APPErrorException {
        mImportResult = JsonUtil.getEmptyObjectNode();
        mTableNames = new ArrayList<String>();
        mParseData = new LinkedHashMap<String, ArrayNode>();
        try {
            mWorkbook = WorkbookFactory.createWorkbook(pInputStream, pVersion);
            CreateWorkbookEvent event = new CreateWorkbookEvent();
            event.setData(pInputStream);
            event.setWorkbook(mWorkbook);
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
        if (mbeginDataRow < 0) {
            throw new APPErrorException("数据行不能小于0");
        }
        if (mbeginTitleRow < 0) {
            throw new APPErrorException("标题行不能小于0");
        }
        if (mbeginDataRow <= mbeginTitleRow) {
            throw new APPErrorException("数据行只能大于标题行；标题行：" + String.valueOf(mbeginTitleRow) + ";数据行：" + String.valueOf(mbeginDataRow));
        }
        if (null != mCoulmns && mCoulmns.size() > 0) {
            parseExcel();
        }
    }

    /**
     * 根据excel列名获取表列明
     *
     * @param pExcelColumnName
     *
     * @return
     */
    private String getAttrName(String pExcelColumnName) {
        for (Map.Entry<String, String> item : mCoulmns.entrySet()) {
            if (item.getValue().equals(pExcelColumnName)) {
                return item.getKey();
            }
        }
        return null;
    }

    /**
     * 获取日期型单元格值
     *
     * @param pDate
     *
     * @return
     */
    private String getDateCellValue(Date pDate) {
        Date date = pDate;
        TimeZone zone = TimeZone.getTimeZone("Asia/Shanghai");
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        sdf.setTimeZone(zone);
        return sdf.format(date);
    }

    /**
     * 用于xls 2003 数据变成科学计数后的问题
     *
     * @param pCell
     * @param pDefaultValue
     *
     * @return
     */
    private String getFixNumericValue(Cell pCell, String pDefaultValue) {
        String result = pDefaultValue;
        try {
            //excel 2003
            if (pCell instanceof HSSFCell) {
                Method method = pCell.getClass().getDeclaredMethod("getCellValueRecord");
                if (null != method) {
                    method.setAccessible(true);
                    CellValueRecordInterface obj = (CellValueRecordInterface) method.invoke(pCell);
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
                    LogHelper.error("系统", "转换科学基数出错，原值：" + pDefaultValue + ";;转换后：" + res, "ImportExcel.getFixNumericValue", ex);
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
     * @param pCell
     *
     * @return
     */
    @SuppressWarnings("deprecation")
    private Object getFORMULACellValue(Cell pCell) {
        String cellValue = "";
        FormulaEvaluator evaluator = pCell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
        CellValue evalValue = evaluator.evaluate(pCell);
        switch (evalValue.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                cellValue = evalValue.getStringValue().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                return evalValue.getNumberValue();
            case Cell.CELL_TYPE_BOOLEAN:
                return evalValue.getBooleanValue();
            default:
                cellValue = "";
        }
        return cellValue;
    }

    /**
     * 获取单元格值
     *
     * @param pCell
     *
     * @return
     */
    @SuppressWarnings("deprecation")
    private Object getCellValue(Cell pCell) {
        String cellValue = "";
        if (null == pCell) {
            return StringUtil.EMPTY_STRING;
        }
        switch (pCell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                cellValue = pCell.getRichStringCellValue().getString().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(pCell)) {
                    cellValue = getDateCellValue(DateUtil.getJavaDate(pCell.getNumericCellValue()));
                    break;
                }

                cellValue = new HSSFDataFormatter().formatCellValue(pCell);
                if(cellValue.indexOf("%") >0)
                {
                    cellValue=String.valueOf(pCell.getNumericCellValue());
                }
                if (cellValue.indexOf(".") < 0) {
                    try {
                        return Integer.parseInt(cellValue);
                    }catch (Exception e)
                    {
                        return Long.parseLong(cellValue);
                    }
                }
                if (!StringUtil.isNullOrEmpty(cellValue) && (cellValue.indexOf("E+") > 0 || cellValue.indexOf("E-") > 0)) {
                    cellValue = getFixNumericValue(pCell, cellValue);
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                return pCell.getBooleanCellValue();
            case Cell.CELL_TYPE_FORMULA:
                return getFORMULACellValue(pCell);
            default:
                cellValue = "";
        }
        return cellValue;

    }

    /**
     * 解析数据集
     *
     * @return
     *
     * @throws APPErrorException
     */
    private void parseDatas() throws APPErrorException {

        int sheetcount = mWorkbook.getNumberOfSheets();
        //循环excel
        for (int i = 0; i < sheetcount; i++) {
            Sheet sheet = mWorkbook.getSheetAt(i);
            if (null == sheet) {
                throw new APPErrorException("获取第" + i + "个表错误");
            }
            int titlerow = mbeginTitleRow;
            int datarow = mbeginDataRow;
            int rowcount = 0;
            Row titleRow = null;
            ArrayNode dt = JsonUtil.getEmptyArrayNode();
            try {
                ReadSheetEvent readevent = new ReadSheetEvent();
                readevent.setCancel(false);
                readevent.setBeginDataRow(mbeginDataRow);
                readevent.setBeginTitleRow(mbeginTitleRow);
                readevent.setSheetName(sheet.getSheetName());
                onReadSheet(readevent);
                if (readevent.getCancel()) {
                    continue;
                }

                //获取总行数（由于获取行号从0开始，所以行数应为行号加1）
                rowcount = sheet.getLastRowNum() + 1;
                LogHelper.info("", "正在解析总行数：" + rowcount, "");
                titleRow = sheet.getRow(titlerow);
            } catch (Exception ex) {
                throw new APPErrorException("解析第" + i + "表错误；起始行：" + mbeginDataRow + "；标题行：" + mbeginTitleRow);
            }
            //循环添加行
            for (int r = datarow; r < rowcount; r++) {
                try {
                    Row sheetrow = sheet.getRow(r);
                    if (null == sheetrow) {
                        continue;
                    }
                    LogHelper.info("", "正在解析行：" + r, "");
                    //int colcount = sheetrow.getPhysicalNumberOfCells();
                    int colcount = sheetrow.getLastCellNum();

                    LogHelper.info("", "正在解析总列数：" + colcount, "");
                    ObjectNode row = JsonUtil.getEmptyObjectNode();
                    //循环添加列值
                    for (int c = 0; c < colcount; c++) {
                        if (null == titleRow.getCell(c)) {
                            break;
                            //throw new APPErrorException("无效的标题cell：" + c + "sheetname:" + sheet.getSheetName());
                        }
                        LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + titleRow.getCell(c).getStringCellValue(), "");
                        //取得excel列名
                        String colName = titleRow.getCell(c).getStringCellValue();
                        String col = getAttrName(colName);

                        if (StringUtil.isNullOrEmpty(col)) {
                            continue;
                        }
                        if(null==sheet.getRow(r).getCell(c))
                        {
                            continue;
                        }
                        Object value = "";
                        try {
                            if(r==19&&c==4)
                            {
                                System.out.println("fddd");
                            }
                            value = getCellValue(sheet.getRow(r).getCell(c));
                            LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + titleRow.getCell(c).getStringCellValue() + "；值：" + value, "");
                        } catch (Exception ex) {
                            throw new APPErrorException("取值错误：" + ex.getMessage() + ";行：" + String.valueOf(r) + "；列：" + String.valueOf(c) + ";sheet:" + sheet.toString() + ";row:" + sheet.getRow(r).toString(), ex);
                        }
                        try {
                            ImportEvent event = new ImportEvent();
                            event.setpColumns(mCoulmns);
                            event.setRowIndex(r);
                            event.setColumnIndex(c);
                            event.setValue(value);

                            onGetValue(event);

                            value = event.getValue();
                            if (value instanceof String) {
                                row.put(col, (String) value);
                            } else {
                                row.putPOJO(col, value);
                            }
                        } catch (Exception ex) {
                            throw new APPErrorException("添加数据到datatable出错：" + ex.getMessage() + ";行：" + String.valueOf(r) + "；列：" + String.valueOf(c) + ";sheet:" + sheet.toString() + ";row:" + sheet.getRow(r).toString(), ex);
                        }
                    }
                    dt.add(row);
                } catch (Exception ex) {
                    throw new APPErrorException("添加数据到datatable出错：" + ex.getMessage() + ";行：" + String.valueOf(r) + ";sheet:" + sheet.toString(), ex);
                }
            }

            SheetParsedEvent sheetevent = new SheetParsedEvent();
            sheetevent.setSheetName(sheet.getSheetName());
            sheetevent.setData(dt);
            onSheetParsed(sheetevent);
            if (null != dt && dt.size() > 0) {
                mParseData.put(sheet.getSheetName(), dt);
                mTableNames.add(sheet.getSheetName());
            }
        }

    }

    /**
     * 解析excel
     *
     * @return
     *
     * @throws APPErrorException
     */
    public void parseExcel() throws APPErrorException {
        try {
            parseDatas();
        } catch (Exception ex) {
            LogHelper.error("", "excel解析错误22", "ImportExcel。parseExcel1", ex);
        }
        if (StringUtil.isNullOrEmpty(mDatabaseType)) {
            return;
        }
        parseToDatabase();

    }
    protected void parseToDatabase() throws APPErrorException {
        try {

            if (null == mParseData || mParseData.size() < 1) {
                throw new APPErrorException("excel没有有效内容");
            }
            int count = 0;
            int sucess = 0;
            int error = 0;
            int i = 0;
            for (Map.Entry<String, ArrayNode> item : mParseData.entrySet()) {
                ArrayNode table = item.getValue();
                count += table.size();
                String tablename = mTableNames.get(i);
                for (int rownum = 0; rownum < table.size(); rownum++) {
                    JsonNode row = table.get(rownum);
                    try {
                        ImportEvent event = new ImportEvent();
                        event.setpColumns(mCoulmns);
                        event.setIsExcute(true);
                        event.setRow(row);
                        String sql = tgtools.util.JsonSqlFactory.parseInsertSql(row, mDatabaseType, tablename);
                        event.setSql(sql);
                        onExcuteSQL(event);

                        if (event.getIsExcute()) {

                            sucess += execute(event.getSql());
                        } else {
                            sucess += event.getIsSucess() ? 1 : 0;
                        }

                    } catch (Exception e) {
                        error = error + 1;
                        LogHelper.error("", "excel导入出错", "ImportExcel.parseExcel1", e);
                    }
                }
                i = i + 1;
            }
            mImportResult.put("count", count);
            mImportResult.put("success", sucess);
            mImportResult.put("error", error);
            ExcelCompletedEvent event = new ExcelCompletedEvent();
            event.setDatas(mParseData);
            event.setWorkbook(mWorkbook);
            onCompleted(event);
        } catch (Exception ex) {
            throw new APPErrorException("解析excel错误，原因：" + ex.getMessage(), ex);
        }
    }
    private int execute(String sql) throws APPErrorException {
        sql = tgtools.util.SqlStrHelper.processKeyWord(sql);
        return tgtools.db.DataBaseFactory.getDefault().executeUpdate(sql);
    }

    @Override
    public ObjectNode getImportResult() {
        return mImportResult;
    }

    @Override
    public LinkedHashMap<String, ArrayNode> getParseData() {
        return new LinkedHashMap<String, ArrayNode>(){{putAll(mParseData);}};
    }

    @Override
    public void close() throws IOException {
        if(null!=mParseData)
        {
            mParseData.clear();
        }
        if (null != mCoulmns) {
            mCoulmns.clear();
        }
        if (null != mImportResult) {
            mImportResult.removeAll();
        }
        if (null != mWorkbook) {
            try {
                mWorkbook.close();
            } catch (Exception e) {
            }
        }
        if (null != mSheetTable) {
            mSheetTable.clear();
        }
        if (null != mTableNames) {
            mTableNames.clear();
        }
        mCoulmns = null;
        mImportResult = null;
        mWorkbook = null;
        mListener = null;
        mSheetTable = null;
        mTableNames = null;
        mParseData=null;
    }


    //------------------------------ Listener  ------------------------------------

    /**
     * 创建excel workbook后对workbook的事件
     *
     * @param pEvent
     */
    protected void onCreateWorkbook(CreateWorkbookEvent pEvent) {
        if (null != mListener) {
            mListener.onCreateWorkbook(pEvent);
        }
    }

    public void onReadSheet(ReadSheetEvent pEvent) {
        if (null != mListener) {
            mListener.onReadSheet(pEvent);
        }
    }

    /**
     * 整个任务完成后事件
     *
     * @param pEvent
     */
    protected void onCompleted(ExcelCompletedEvent pEvent) {
        if (null != mListener) {
            mListener.onCompleted(pEvent);
        }
    }

    protected void onExcuteSQL(ImportEvent pEvent) {
        if (null != mListener) {
            mListener.onExcuteSQL(pEvent);
        }
    }

    public void onSheetParsed(SheetParsedEvent pEvent) {
        if (null != mListener) {
            mListener.onSheetParsed(pEvent);
        }
    }

    protected void onGetValue(ImportEvent pEvent) {
        if (null != mListener) {
            mListener.onGetValue(pEvent);
        }
    }

    public static void main(String[] args) {
        String filepath = "C:\\Users\\tian_\\Desktop\\业务联系电话.xls";
        ImportExcelImpl importExcel = new ImportExcelImpl();
        LinkedHashMap<String, String> columns = new LinkedHashMap<String, String>();
        columns.put("DW", "单位");
        columns.put("SJ", "值班电话");
        columns.put("XMMC", "班组驻地电话");
        columns.put("KMFL", "外线");
        columns.put("KMXF", "联系人及联系方式");
//        columns.put("JE", "金额");
//        columns.put("SL", "税率");
//        columns.put("SE", "税额");
//        columns.put("JSHJ", "价税合计");
//        columns.put("FPHM", "发票号码");
//        columns.put("SKRMC", "收款人名称");
//        columns.put("KHH", "开户行");
//        columns.put("YHZH", "银行账户");
//        columns.put("CS", "处室");
//        columns.put("JBR", "经办人");
//        columns.put("BZ", "备注");
        HashMap<String, String> table = new HashMap<String, String>();
        table.put("sheet1", "MQ_SYS.ACT_ID_USER");
        //默认不做数据库操作 之转换成json
        importExcel.init(columns, table,null,1,2);
        importExcel.setListener(new ImportListener(){

            @Override
            public void onCreateWorkbook(CreateWorkbookEvent pEvent) {

            }

            @Override
            public void onCompleted(ExcelCompletedEvent pEvent) {

            }

            @Override
            public void onLoadFilter(ImportEvent pEvent) {

            }

            @Override
            public void onGetAtted(ImportEvent pEvent) {

            }

            @Override
            public void onGetValue(ImportEvent pEvent) {
                System.out.println(pEvent.getValue());
                if("18651244052".equals(pEvent.getValue()))
                {
                    System.out.println(pEvent.getValue());
                }
            }

            @Override
            public void onExcuteSQL(ImportEvent pEvent) {

            }

            @Override
            public void onReadSheet(ReadSheetEvent pEvent) {

            }

            @Override
            public void onSheetParsed(SheetParsedEvent pEvent) {

            }
        });
        //设置数据库类型后进行sql 操作
        //importExcel.init(columns, table,"dm");
        try {
            importExcel.importExcel(new File(filepath));
            Map<String, ArrayNode> ds = importExcel.getParseData();
            importExcel.close();
            System.out.println(ds);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
