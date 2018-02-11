package tgtools.excel.poi;


import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.hssf.record.CellValueRecordInterface;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import tgtools.excel.ImportExcel;
import tgtools.excel.Listener.ImportListener;
import tgtools.excel.Listener.event.*;
import tgtools.excel.util.JsonUtil;
import tgtools.exceptions.APPErrorException;
import tgtools.json.JSONObject;
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
    protected Map<String,String> mSheetTable;
    protected ArrayList<String> mTableNames;

    @Override
    public void setListener(ImportListener pListener) {
        mListener = pListener;
    }

    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String,String> pSheetTable) {
        init(pColumns, pSheetTable,0, 1);
    }


    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String,String> pSheetTable, int p_BeginDataRow) {
        init(pColumns,pSheetTable, 0, p_BeginDataRow);
    }
    @Override
    public void init(LinkedHashMap<String, String> pColumns, Map<String,String> pSheetTable, int p_BeginTitleRow, int p_BeginDataRow) {
        mCoulmns=pColumns;
        mbeginTitleRow=p_BeginTitleRow;
        mbeginDataRow=p_BeginDataRow;
        mSheetTable=pSheetTable;
    }

    @Override
    public void importExcel(File pFile) throws Exception {

        if (null == pFile || !pFile.exists()) {
            throw new Exception("文件不存在！p_File：" + (null == pFile ? "null" : pFile.getAbsolutePath()));
        }
        importExcel(new FileInputStream(pFile));
    }

    @Override
    public void importExcel(byte[] pDatas) throws Exception {
        if(null==pDatas||pDatas.length<1)
        {
            throw new Exception("无效的文件内容p_Datas");
        }
        importExcel(new ByteArrayInputStream(pDatas));
    }

    @Override
    public void importExcel(InputStream pInputStream) throws Exception {
        mImportResult = JsonUtil.getEmptyObjectNode();
        mTableNames=new ArrayList<String>();
        try {
            mWorkbook = WorkbookFactory.createWorkbook(pInputStream);
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
        if (null != mCoulmns&&mCoulmns.size()>0) {
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
        for(Map.Entry<String,String> item :mCoulmns.entrySet())
        {
            if(item.getValue().equals(pExcelColumnName))
            {return item.getKey();}
        }
        return null;
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
     * 解析数据集
     *
     * @return
     *
     * @throws APPErrorException
     */
    private List<ArrayNode> parseDatas() throws APPErrorException {
        List<ArrayNode> res =new ArrayList<ArrayNode>();

        int sheetcount = mWorkbook.getNumberOfSheets();
        for (int i = 0; i < sheetcount; i++) {//循环excel
            Sheet sheet = mWorkbook.getSheetAt(i);
            if (null == sheet) {
                throw new APPErrorException("获取第" + i + "个表错误");
            }
            int titlerow = 0;
            int datarow = 0;
            int rowcount = 0;
            Row titleRow = null;
            ArrayNode dt = null;
            try {
                ReadSheetEvent readevent = new ReadSheetEvent();
                readevent.setCancel(false);
                readevent.setbeginDataRow(mbeginDataRow);
                readevent.setbeginTitleRow(mbeginTitleRow);
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
                    int colcount = sheetrow.getPhysicalNumberOfCells();
                    LogHelper.info("", "正在解析总列数：" + colcount, "");
                    ObjectNode row=JsonUtil.getEmptyObjectNode();
                    //循环添加列值
                    for (int c = 0; c < colcount; c++)
                    {
                        if (null == titleRow.getCell(c)) {
                            throw new APPErrorException("无效的标题cell：" + c + "sheetname:" + sheet.getSheetName());
                        }
                        LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + titleRow.getCell(c).getStringCellValue(), "");
                        //取得excel列名
                        String colName = titleRow.getCell(c).getStringCellValue();
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
                            event.setpColumns(mCoulmns);
                            event.setRowIndex(r);
                            event.setColumnIndex(c);
                            event.setValue(value);

                            onGetValue(event);

                            value = event.getValue();

                            row.put(col, value);
                        } catch (Exception ex) {
                            throw new APPErrorException("添加数据到datatable出错：" + ex.getMessage() + ";行：" + String.valueOf(r) + "；列：" + String.valueOf(c) + ";sheet:" + sheet.toString() + ";row:" + sheet.getRow(r).toString(), ex);
                        }
                    }
                } catch (Exception ex) {
                    throw new APPErrorException("添加数据到datatable出错：" + ex.getMessage() + ";行：" + String.valueOf(r) + ";sheet:" + sheet.toString(), ex);
                }
            }

            SheetParsedEvent sheetevent = new SheetParsedEvent();
            sheetevent.setSheetName(sheet.getSheetName());
            sheetevent.setData(dt);
            onSheetParsed(sheetevent);
            if (null != dt && dt.size() > 0) {
                res.add(dt);
                mTableNames.add(sheet.getSheetName());
            }
        }
        return res;
    }
    /**
     * 解析excel
     *
     * @return
     *
     * @throws APPErrorException
     */
    public void parseExcel() throws APPErrorException {
        List<ArrayNode> datas = null;
        try {
            datas = parseDatas();
        } catch (Exception ex) {
            LogHelper.error("", "excel解析错误22", "ImportExcel。parseExcel1", ex);
        }

        try {

            if (null == datas || datas.size() < 1) {
                throw new APPErrorException("excel没有有效内容");
            }
            int count = 0;
            int sucess = 0;
            int error=0;
            for (int i = 0; i < datas.size(); i++) {
                ArrayNode table = datas.get(i);
                count += table.size();
                String tablename = mTableNames.get(i);
                for (int rownum = 0; rownum < table.size(); rownum++) {
                    JsonNode row = table.get(rownum);
                    try {
                        ImportEvent event = new ImportEvent();
                        event.setpColumns(mCoulmns);
                        event.setIsExcute(true);
                        event.setRow(row);

                        onExcuteSQL(event);

                        String sql = tgtools.util.JsonSqlFactory.parseInsertSql(new JSONObject(row.toString()), tablename);
                        if (event.isExcute()) {
                            sucess += execute(sql);
                        } else {
                            sucess += event.isSucess() ? 1 : 0;
                        }

                    } catch (Exception e) {
                        error=error+1;
                        LogHelper.error("", "excel导入出错", "ImportExcel.parseExcel1", e);
                    }
                }
            }
            mImportResult.put("count",count);
            mImportResult.put("success",sucess);
            mImportResult.put("error",error);
            ExcelCompletedEvent event = new ExcelCompletedEvent();
            event.setDatas(datas);
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
    public void close() throws IOException {
        if(null!=mCoulmns) {
            mCoulmns.clear();
        }
        if(null!=mImportResult) {
            mImportResult.removeAll();
        }
        if(null!=mWorkbook) {
            try {
                mWorkbook.close();
            }catch (Exception e)
            {}
        }
        if(null!=mSheetTable) {
            mSheetTable.clear();
        }
        if(null!=mTableNames) {
            mTableNames.clear();
        }
        mCoulmns=null;
        mImportResult=null;
        mWorkbook = null;
        mListener=null;
        mSheetTable=null;
        mTableNames=null;

    }


    //------------------------------ Listener  ------------------------------------
    /**
     * 创建excel workbook后对workbook的事件
     *
     * @param p_Event
     */
    protected void onCreateWorkbook(CreateWorkbookEvent p_Event) {
        if (null != mListener) {
            mListener.onCreateWorkbook(p_Event);
        }
    }
    public void onReadSheet(ReadSheetEvent p_Event) {
        if (null != mListener) {
            mListener.onReadSheet(p_Event);
        }
    }
    /**
     * 整个任务完成后事件
     *
     * @param p_Event
     */
    protected void onCompleted(ExcelCompletedEvent p_Event) {
        if (null != mListener) {
            mListener.onCompleted(p_Event);
        }
    }
    protected void onExcuteSQL(ImportEvent p_Event) {
        if (null != mListener) {
            mListener.onExcuteSQL(p_Event);
        }
    }
    public void onSheetParsed(SheetParsedEvent p_Event) {
        if (null != mListener) {
            mListener.onSheetParsed(p_Event);
        }
    }
    protected void onGetValue(ImportEvent p_Event) {
        if (null != mListener) {
            mListener.onGetValue(p_Event);
        }
    }
}
