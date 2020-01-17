package tgtools.excel.jxl;


import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import tgtools.excel.ImportExcel;
import tgtools.excel.listener.ImportListener;
import tgtools.excel.listener.event.*;
import tgtools.excel.util.JsonUtil;
import tgtools.exceptions.APPErrorException;
import tgtools.util.FileUtil;
import tgtools.util.LogHelper;
import tgtools.util.StringUtil;

import java.io.*;
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
     * 获取单元格值
     *
     * @param pCell
     *
     * @return
     */
    @SuppressWarnings("deprecation")
    private Object getCellValue(Cell pCell) {
        return pCell.getContents();
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
            Sheet sheet = mWorkbook.getSheet(i);
            if (null == sheet) {
                throw new APPErrorException("获取第" + i + "个表错误");
            }
            int titlerow = mbeginTitleRow;
            int datarow = mbeginDataRow;
            int rowcount = 0;
            Cell[] titleRow = null;
            ArrayNode dt = JsonUtil.getEmptyArrayNode();
            try {
                ReadSheetEvent readevent = new ReadSheetEvent();
                readevent.setCancel(false);
                readevent.setBeginDataRow(mbeginDataRow);
                readevent.setBeginTitleRow(mbeginTitleRow);
                readevent.setSheetName(sheet.getName());
                onReadSheet(readevent);
                if (readevent.getCancel()) {
                    continue;
                }

                //获取总行数（由于获取行号从0开始，所以行数应为行号加1）
                rowcount = sheet.getRows();
                LogHelper.info("", "正在解析总行数：" + rowcount, "");
                titleRow = sheet.getRow(titlerow);
            } catch (Exception ex) {
                throw new APPErrorException("解析第" + i + "表错误；起始行：" + mbeginDataRow + "；标题行：" + mbeginTitleRow);
            }
            //循环添加行
            for (int r = datarow; r < rowcount; r++) {
                try {
                    Cell[] sheetrow = sheet.getRow(r);
                    if (null == sheetrow) {
                        continue;
                    }
                    LogHelper.info("", "正在解析行：" + r, "");
                    int colcount = sheetrow.length;

                    LogHelper.info("", "正在解析总列数：" + colcount, "");
                    ObjectNode row = JsonUtil.getEmptyObjectNode();
                    //循环添加列值
                    for (int c = 0; c < colcount; c++) {
                        if (null == titleRow[c]) {
                            break;
                            //throw new APPErrorException("无效的标题cell：" + c + "sheetname:" + sheet.getSheetName());
                        }
                        String colName = titleRow[c].getContents();
                        LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + colName, "");
                        //取得excel列名
                        String col = getAttrName(colName);

                        if (StringUtil.isNullOrEmpty(col)) {
                            continue;
                        }
                        if(null==sheet.getCell(c,r))
                        {
                            continue;
                        }
                        Object value = "";
                        try {
                            if(r==19&&c==4)
                            {
                                System.out.println("fddd");
                            }
                            value = getCellValue(sheet.getCell(c,r));
                            LogHelper.info("", "正在解析行：" + r + "正在解析列：" + c + "标题：" + colName + "；值：" + value, "");
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
            sheetevent.setSheetName(sheet.getName());
            sheetevent.setData(dt);
            onSheetParsed(sheetevent);
            if (null != dt && dt.size() > 0) {
                mParseData.put(sheet.getName(), dt);
                mTableNames.add(sheet.getName());
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
                String tablename = getTableName(mTableNames.get(i));
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
    protected String getTableName(String pSheetName)
    {
        if(!StringUtil.isNullOrEmpty(pSheetName)&&mSheetTable.containsKey(pSheetName))
        {
            return mSheetTable.get(pSheetName);
        }
        return pSheetName;
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


}
