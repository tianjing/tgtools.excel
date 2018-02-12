package tgtools.excel.poi;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import org.apache.poi.ss.usermodel.*;
import tgtools.excel.ExportExcel;
import tgtools.excel.listener.ExportListener;
import tgtools.excel.listener.event.CreateWorkbookEvent;
import tgtools.excel.listener.event.ExcelCompletedEvent;
import tgtools.excel.listener.event.ExportExcelEvent;
import tgtools.exceptions.APPErrorException;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author 田径
 * @Title
 * @Description
 * @date 8:40
 */
public class ExportExcelImpl implements ExportExcel {
    protected Workbook mWorkbook;
    protected ExportListener mListener;
    protected String mVersion = null;
    protected int mDataIndex = 0;
    protected ArrayNode mDatas = null;
    protected LinkedHashMap<String, String> mColumns = null;
    protected boolean mIsExportTitle = true;

    @Override
    public void init(String pVersion) throws APPErrorException {
        mVersion = pVersion;
        createWorkbook();
    }


    /**
     * 创建excel对象
     *
     * @throws Exception
     */
    protected void createWorkbook() throws APPErrorException {
        mWorkbook = WorkbookFactory.createWorkbook(mVersion);
        CreateWorkbookEvent event = new CreateWorkbookEvent();
        event.setData(mDatas);
        event.setWorkbook(mWorkbook);
        onCreateWorkbook(event);
    }

    @Override
    public void setLisener(ExportListener pListener) throws APPErrorException {
        mListener = pListener;
    }

    @Override
    public void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson) throws APPErrorException {
        appendData(pColumns, pJson, true);
    }

    @Override
    public void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle) throws APPErrorException {
        appendData(pColumns, pJson, pIsExportTitle, "sheet1", 0, 0);
    }

    @Override
    public void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle, String pSheetName, int pIndex, int pDataIndex) throws APPErrorException {
        mDataIndex = pDataIndex;
        mDatas = pJson;
        mColumns = pColumns;
        mIsExportTitle = pIsExportTitle;
        Sheet sheet = createSheet(pSheetName, pIndex);
        writeExcel(sheet);
    }

    /**
     * 将数据写入sheet
     *
     * @param sheet1
     *
     * @throws Exception
     */
    protected void writeExcel(Sheet sheet1) throws APPErrorException {
        final Sheet sheet2=sheet1;
        try {
            if (mIsExportTitle) {
                writeTitle(sheet1);
            }
            writeContent(sheet1);

        } finally {
            ExcelCompletedEvent event = new ExcelCompletedEvent();
            event.setWorkbook(mWorkbook);
            event.setDatas(new LinkedHashMap<String, ArrayNode>() {{
                put(sheet2.getSheetName(),mDatas);
            }});
            onCompleted(event);
        }
    }

    /**
     * 向excel写入数据
     *
     * @param sheet1
     */
    protected void writeContent(Sheet sheet1) {
        for (int i = 0, count = mDatas.size(); i < count; i++) {

            Row row = sheet1.createRow(mDataIndex + i + 1);
            int k = 0;
            for (Map.Entry<String, String> item : mColumns.entrySet()) {
                String name = item.getKey();
                JsonNode value = mDatas.get(i).get(name);

                Cell cell = row.createCell(k);
                if (value.isTextual()) {
                    cell.setCellValue(value.asText());
                } else if (value.isBigDecimal()) {
                    cell.setCellValue(value.asDouble());
                } else if (value.isBigInteger()) {
                    cell.setCellValue(value.asLong());
                } else if (value.isInt()) {
                    cell.setCellValue(value.asInt());
                    CellStyle style=mWorkbook.createCellStyle();
                    DataFormat df = mWorkbook.createDataFormat();
                    style.setDataFormat(df.getFormat("#,#0"));
                    cell.setCellStyle(style);
                }else if (value.isBoolean()) {
                    cell.setCellValue(value.asBoolean());
                } else {
                    cell.setCellValue(value.toString());
                }

                ExportExcelEvent event = new ExportExcelEvent();
                event.setDatas(mDatas);
                event.setRowIndex(i);
                event.setCellIndex(k);
                event.setValue(cell);
                onWriteCell(event);
                k = k + 1;
            }
        }

    }

    /**
     * 向excel写入标题行
     *
     * @param sheet1
     */
    protected void writeTitle(Sheet sheet1) {
        Row row = sheet1.createRow(0);
        int k = 0;
        for (Map.Entry<String, String> item : mColumns.entrySet()) {
            String nickname = item.getValue();
            Cell cell = row.createCell(k);
            cell.setCellValue(nickname);

            ExportExcelEvent event = new ExportExcelEvent();
            event.setValue(cell);
            onWriteTitle(event);

            k = k + 1;
        }

    }

    /**
     * 创建excel对象
     *
     * @param pSheetName
     * @param pIndex
     *
     * @return
     *
     * @throws Exception
     */
    protected Sheet createSheet(String pSheetName, int pIndex) throws APPErrorException {
        Sheet s = mWorkbook.getSheet(pSheetName);
        if (null == s) {
            s = mWorkbook.createSheet(pSheetName);
            mWorkbook.setSheetName(pIndex, pSheetName);
        }
        return s;
    }

    @Override
    public byte[] getBytes() throws APPErrorException {
        return ((ByteArrayOutputStream) getOutputStream()).toByteArray();
    }

    @Override
    public OutputStream getOutputStream() throws APPErrorException {
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            mWorkbook.write(out);
            return out;
        }catch (Exception e)
        {
            throw new APPErrorException("Workbook write ByteArrayOutputStream 出错，原因："+e.getMessage(),e);
        }
    }

    @Override
    public Object getExcel() {
        return mWorkbook;
    }

    /**
     * 释放对象
     *
     * @param pIsRelease 是否完全清空引用
     */
    protected void close(boolean pIsRelease) {
        try {
            mWorkbook.close();
        } catch (Exception e) {
        }

        if (pIsRelease) {
            mListener = null;
            mWorkbook = null;
        }
    }

    @Override
    public void close() throws IOException {
        close(true);
    }

    //-------------------------------listen--------------------------------

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

    /**
     * 写入单元格事件
     *
     * @param pEvent
     */
    protected void onWriteCell(ExportExcelEvent pEvent) {
        if (null != mListener) {
            mListener.onWriteCell(pEvent);
        }
    }

    /**
     * 写入标题事件
     *
     * @param pEvent
     */
    protected void onWriteTitle(ExportExcelEvent pEvent) {
        if (null != mListener) {
            mListener.onWriteTitle(pEvent);
        }
    }
}
