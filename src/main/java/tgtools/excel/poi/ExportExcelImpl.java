package tgtools.excel.poi;

import com.fasterxml.jackson.databind.node.ArrayNode;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import tgtools.excel.ExportExcel;
import tgtools.excel.Listener.ExportListener;
import tgtools.excel.Listener.event.CreateWorkbookEvent;
import tgtools.excel.Listener.event.ExcelCompletedEvent;
import tgtools.excel.Listener.event.ExportExcelEvent;
import tgtools.util.FileUtil;
import tgtools.util.StringUtil;

import java.io.*;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author 田径
 * @Title
 * @Description
 * @date 8:40
 */
public class ExportExcelImpl implements ExportExcel {
    protected OutputStream mOutputStream;
    protected Workbook mWorkbook;
    protected ExportListener mListener;
    protected boolean mIsFile = false;
    protected String mVersion = null;
    protected int mDataIndex = 1;
    protected ArrayNode mDatas = null;
    protected LinkedHashMap<String, String> mColumns = null;
    protected boolean mIsExportTitle = true;

    @Override
    public void init(String pVersion) throws Exception {
        init(pVersion, new ByteArrayOutputStream());
    }

    @Override
    public void init(File pFile) throws Exception {
        if (null == pFile) {
            throw new Exception("pFile 不能为空");
        }
        if (!pFile.exists()) {
            throw new Exception("文件不存在。pFile：" + pFile.getAbsolutePath());
        }
        String version = FileUtil.getExtensionName(pFile.getName());
        mIsFile = true;
        init(version, new FileOutputStream(pFile));
    }

    @Override
    public void init(String pVersion, OutputStream pOutputStream) throws Exception {
        if (StringUtil.isNullOrEmpty(pVersion)) {
            throw new Exception("pVersion 不能为空");
        }
        if (!ExportExcel.VERSION_EXCEL2003.equals(pVersion) || !ExportExcel.VERSION_EXCEL2007.equals(pVersion) ||
                !("." + ExportExcel.VERSION_EXCEL2003).equals(pVersion) || !("." + ExportExcel.VERSION_EXCEL2007).equals(pVersion)) {
            throw new Exception("pVersion 信息超出范围，请参看 ExportExcel");
        }
        if (null == pOutputStream) {
            throw new Exception("pOutputStream 不能为空");
        }
        this.mVersion = pVersion;
        this.mOutputStream = pOutputStream;
    }

    /**
     * 创建excel对象
     *
     * @throws Exception
     */
    protected void createWorkbook() throws Exception {
        mWorkbook = WorkbookFactory.createWorkbook(mVersion);

        CreateWorkbookEvent event = new CreateWorkbookEvent();
        event.setData(mOutputStream);
        event.setWorkbook(mWorkbook);
        onCreateWorkbook(event);
    }

    @Override
    public void setLisener(ExportListener pListener) throws Exception {
        mListener = pListener;
    }

    @Override
    public void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson) throws Exception {
        appendData(pColumns, pJson, true);
    }

    @Override
    public void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle) throws Exception {
        appendData(pColumns, pJson, pIsExportTitle, "sheet1", 0, 1);
    }

    @Override
    public void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle, String p_sheetName, int p_Index, int pDataIndex) throws Exception {
        mDataIndex = pDataIndex;
        mDatas = pJson;
        mColumns = pColumns;
        mIsExportTitle = pIsExportTitle;
        createWorkbook();
        Sheet sheet = createSheet(p_sheetName, p_Index);
        writeExcel(sheet);
    }

    /**
     * 将数据写入sheet
     *
     * @param sheet1
     *
     * @throws Exception
     */
    protected void writeExcel(Sheet sheet1) throws Exception {
        try {
            writeTitle(sheet1);
            writeContent(sheet1);

        } finally {
            ExcelCompletedEvent event = new ExcelCompletedEvent();
            event.setWorkbook(mWorkbook);
            event.setDatas(new ArrayList<ArrayNode>(){{add(mDatas);}});
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

            Row row = sheet1.createRow(i + 1);
            int k=0;
            for (Map.Entry<String, String> item : mColumns.entrySet()) {
                String name =item.getKey();
                Object value = mDatas.get(i).get(name);

                Cell cell = row.createCell(k);
                cell.setCellValue(value.toString());

                ExportExcelEvent event = new ExportExcelEvent();
                event.setDatas(mDatas);
                event.setRowIndex(i);
                event.setCellIndex(k);
                event.setValue(cell);
                onWriteCell(event);
                k=k+1;
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
        for (Map.Entry<String, String> item :mColumns.entrySet()) {
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
     * @param p_SheetName
     * @param p_Index
     *
     * @return
     *
     * @throws Exception
     */
    protected Sheet createSheet(String p_SheetName, int p_Index) throws Exception {
        Sheet s = mWorkbook.getSheet(p_SheetName);
        if (null == s) {
            s = mWorkbook.createSheet(p_SheetName);
            mWorkbook.setSheetName(p_Index, p_SheetName);
        }
        return s;
    }

    @Override
    public byte[] getBytes() throws Exception {
        if (mOutputStream instanceof ByteArrayOutputStream) {
            close(false);
            return ((ByteArrayOutputStream) mOutputStream).toByteArray();
        }
        throw new Exception("当前使用的不是内存流，请使用getOutputStream()");
    }

    @Override
    public OutputStream getOutputStream() throws Exception {
        return mOutputStream;
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
        try {
            mOutputStream.close();
        } catch (Exception e) {
        }
        if (pIsRelease) {
            mListener = null;
            mOutputStream = null;
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
     * @param p_Event
     */
    protected void onCreateWorkbook(CreateWorkbookEvent p_Event) {
        if (null != mListener) {
            mListener.onCreateWorkbook(p_Event);
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

    /**
     * 写入单元格事件
     * @param pEvent
     */
    protected void onWriteCell(ExportExcelEvent pEvent)
    {
        if (null != mListener) {
            mListener.onWriteCell(pEvent);
        }
    }

    /**
     * 写入标题事件
     * @param pEvent
     */
    protected void onWriteTitle(ExportExcelEvent pEvent)
    {
        if (null != mListener) {
            mListener.onWriteTitle(pEvent);
        }
    }
}
