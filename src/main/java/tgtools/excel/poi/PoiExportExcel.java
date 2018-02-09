package tgtools.excel.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import tgtools.data.DataTable;
import tgtools.excel.ExportExcel;
import tgtools.excel.Listener.ExcelAllLitener;
import tgtools.excel.Listener.ExportLisener;
import tgtools.excel.Listener.event.CreateWorkbookEvent;
import tgtools.excel.Listener.event.ExcelCompletedEvent;
import tgtools.excel.Listener.event.ExportExcelEvent;
import tgtools.exceptions.APPErrorException;
import tgtools.util.StringUtil;
import tgtools.web.entity.GridDataEntity;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * excel标准导出类 将表所有信息导出
 * 使用方式
 *
 *  excel.init();//初始化
 *  excel.setLisener();//设置监听，进行功能扩展（不必须）
 *  excel.buildExcel();//生成EXCEL
 *  （1）excel.getBytes();//结束写入并取回excel数据
 *  （2）excel.getExcel();//获取excel对象继续自定义写入
 *  excel.Dispose();//释放
 * Created by tian_ on 2016-06-17.
 */
public class PoiExportExcel implements ExportExcel {
    public PoiExportExcel(){
        this(WorkbookFactory.EXCEL_TYPE_XLSX);
    }
    public PoiExportExcel(String p_Type){
        m_ExcelType=p_Type;
    }
    protected String m_ExcelType;
    protected String m_SortsStr;
    protected String m_PageIndex;
    protected String m_PageSize;

    protected String m_type;
    protected String m_cutPage;

    protected String[] m_attrName;
    protected String[] m_attr;
    protected String m_tabName;
    protected ExcelAllLitener m_AllLitener;
    protected ByteArrayOutputStream m_OutputStream;
    protected Workbook m_Workbook;
    protected ExportLisener m_Lisener;

    public static void main(String[] args) throws APPErrorException, IOException {

        tgtools.db.DataBaseFactory.add("DM", new Object[]{"jdbc:dm://192.168.88.128:5235/dqmis", "BQ_SYS", "BQ_SYS123"});
        String SortsStr = "";
        String cutPage = "yes";
        String PageIndex = "1";
        String PageSize = "999999999";
        String type = "only";
        String SortOrder = "";
        PoiExportExcel importExcel = new PoiExportExcel();


        String[] attrs = new String[]{"USERNAME", "LOGTIME","LOGTYPE","BIZTYPE","LOGCONTENT"};
        String[] attrnames = new String[]{"名字", "日志时间","日志类型","日志业务","日志内容"};
        String tablename = "BQ_SYS.LOGINFO";
        importExcel.init(attrs, attrnames, tablename, SortsStr, SortOrder, PageIndex, PageSize, cutPage, type);
        importExcel.buildExcel();
        importExcel.close();

        byte[] res = importExcel.getBytes();

        importExcel.close();

        File file = new File("c:\\tianjing\\Desktop\\1.xls");
        if (file.exists()) {
            file.delete();
        }
        file.createNewFile();
        FileOutputStream out = new FileOutputStream(file);
        out.write(res);
        out.close();
    }

    /**
     * 获取监听对象
     *
     * @return
     */
    public ExportLisener getLisener() {
        return m_Lisener;
    }

    /**
     * 设置监听对象
     *
     * @param p_Lisener
     */
    @Override
    public void setLisener(ExportLisener p_Lisener) {
        m_Lisener = p_Lisener;
    }

    @Override
    public void setExcelAllLisener(ExcelAllLitener p_Litener) {
        m_AllLitener=p_Litener;
    }



    /**
     * 初始化参数
     *
     * @param p_Attr      表列名 必填
     * @param p_AttrNames 列中文名 必填
     * @param p_tabName   表名称 必填
     * @param p_SortsStr  排序
     * @param p_SortOrder
     * @param p_CutPage   yes 与 p_Type only 组合使用表示全部导出
     * @param p_PageIndex 当前索引 必填
     * @param p_PageSize  页大小 必填
     * @param p_Type      only 全部导出
     */
    @Override
    public void init(String[] p_Attr, String[] p_AttrNames, String p_tabName, String p_SortsStr, String p_SortOrder, String p_PageIndex, String p_PageSize, String p_CutPage, String p_Type) {
        m_attr = p_Attr;
        m_attrName = p_AttrNames;
        m_tabName = p_tabName;
        m_SortsStr = p_SortsStr;
        if (!StringUtil.isNullOrEmpty(p_SortOrder) && !StringUtil.isNullOrEmpty(p_SortsStr)) {
            m_SortsStr += " " + p_SortOrder;
        }
        m_PageIndex = p_PageIndex;
        m_PageSize = StringUtil.isNullOrEmpty(p_PageSize) ? "15" : p_PageSize;
        m_cutPage = p_CutPage;
        m_type = p_Type;


    }

    /**
     * 获取excel对象
     *
     * @return
     */
    public Object getExcel() {
        return m_Workbook;
    }

    /**
     * 获取生成excel文件的内容,并关闭流，不可再写入
     *
     * @return
     */
    @Override
    public byte[] getBytes() throws APPErrorException{
        try {
            m_Workbook.write(m_OutputStream);
            close();
            return m_OutputStream.toByteArray();
        }catch (Exception ex)
        {
           throw new APPErrorException("获取excel数据出错。原因："+ex.getMessage(),ex);
        }
    }

    private void validParams() throws APPErrorException {
        if (StringUtil.isNullOrEmpty(m_tabName)) {
            throw new APPErrorException("tabName 表名称不能为空");
        }
        if (StringUtil.isNullOrEmpty(m_PageIndex)) {
            throw new APPErrorException("PageIndex 当前页不能为空");
        }
        if (StringUtil.isNullOrEmpty(m_PageSize)) {
            throw new APPErrorException("PageSize 页大小不能唯恐不能为空");
        }
        if (!tgtools.util.RegexHelper.isNubmer(m_PageIndex)) {
            throw new APPErrorException("PageIndex 当前页必须是数字 PageIndex：" + m_PageIndex);
        }
        if (!tgtools.util.RegexHelper.isNubmer(m_PageSize)) {
            throw new APPErrorException("PageSize 页大小必须是数字 PageSize：" + m_PageSize);
        }
    }

    /**
     * 生成excel
     *
     * @throws APPErrorException
     */
    public void buildExcel() throws APPErrorException {
        buildExcel("sheet1", 0);
    }

    /**
     * 生成excel
     *
     * @param p_sheetName
     * @param p_Index
     * @throws APPErrorException
     */
    @Override
    public void buildExcel(String p_sheetName, int p_Index) throws APPErrorException {
        validParams();
        createWorkbook();
        Sheet sheet = createSheet(p_sheetName, p_Index);
        writeExcel(sheet);
    }

    @Override
    public void close() {
        try {
            m_Workbook.close();
        } catch (Exception e) {
        }
        try {
            m_OutputStream.close();
        } catch (Exception e) {
        }
        m_OutputStream = null;
        m_Workbook = null;
    }

    /**
     * 创建excel对象
     *
     * @param p_SheetName
     * @param p_Index
     * @return
     * @throws APPErrorException
     */
    protected Sheet createSheet(String p_SheetName, int p_Index) throws APPErrorException {
        Sheet s= m_Workbook.createSheet(p_SheetName);
        m_Workbook.setSheetName(p_Index,p_SheetName);
        return s;
    }

    /**
     * 创建excel对象
     *
     * @throws APPErrorException
     */
    protected void createWorkbook() throws APPErrorException {

            m_OutputStream = new ByteArrayOutputStream();
            m_Workbook =WorkbookFactory.createWorkbook(m_ExcelType);

            CreateWorkbookEvent event =new CreateWorkbookEvent();
            event.setData(m_OutputStream);
            event.setWorkbook(m_Workbook);
            onCreateWorkbook(event);
    }

    /**
     * 向excel写入标题行
     *
     * @param sheet1
     * @param p_Table
     */
    protected void writeTitle(Sheet sheet1, DataTable p_Table)  {
        Row row= sheet1.createRow(0);
        for (int k = 0; k < m_attrName.length; k++) {
            String name = m_attrName[k];
            Cell cell= row.createCell(k);
            cell.setCellValue(name);
            if (null != m_Lisener) {
                ExportExcelEvent event = new ExportExcelEvent();
                event.setValue(cell);
                event.setTable(p_Table);
                event.setcellIndex(k);
                m_Lisener.onWriteTitle(event);
            }
        }

    }

    /**
     * 向excel写入数据
     *
     * @param sheet1
     * @param p_Table
     */
    protected void writeContent(Sheet sheet1, DataTable p_Table)  {
        for (int i = 0, count = p_Table.getRows().size(); i < count; i++) {

            Row row= sheet1.createRow(i + 1);
            for (int k = 0, colcount = m_attr.length; k < colcount; k++) {
                String name = m_attr[k];
                Object value = p_Table.getRow(i).getValue(name);

                Cell cell= row.createCell(k);
                cell.setCellValue(value.toString());

                if (null != m_Lisener) {
                    ExportExcelEvent event = new ExportExcelEvent();
                    event.setTable(p_Table);
                    event.setRowIndex(i);
                    event.setcellIndex(k);
                    event.setValue(cell);
                    m_Lisener.onWriteCell(event);
                }
            }
        }

    }

    /**
     * 获取要导出的数据
     *
     * @param p_UsePage
     * @return
     * @throws APPErrorException
     */
    protected GridDataEntity getListData(boolean p_UsePage) throws APPErrorException {
        ViewModel viewmodel = new ViewModel();
        return viewmodel.getListData(m_tabName, Integer.valueOf(m_PageIndex), Integer.valueOf(m_PageSize), m_SortsStr, "", null, "", p_UsePage);
    }
    /**
     * 创建excel workbook后对workbook的事件
     * @param p_Event
     */
    protected void onCreateWorkbook(CreateWorkbookEvent p_Event){
        if(null!=m_AllLitener)
        {
            m_AllLitener.onCreateWorkbook(p_Event);
        }
    }



    /**
     * 整个任务完成后事件
     * @param p_Event
     */
    protected void onCompleted(ExcelCompletedEvent p_Event){
        if(null!=m_AllLitener)
        {
            m_AllLitener.onCompleted(p_Event);
        }
    }

    protected void onGetDataed(ExportExcelEvent p_Event) throws APPErrorException {
        if (null != m_Lisener) {
            m_Lisener.onGetDataed(p_Event);
        }
    }


    /**
     * 将数据写入sheet
     *
     * @param sheet1
     * @throws APPErrorException
     */
    protected void writeExcel(Sheet sheet1) throws APPErrorException {
        DataTable dt =null;
        try {
            boolean usepage = "only".equals(m_type) && "yes".equals(m_cutPage);
            ExportExcelEvent event = new ExportExcelEvent();
            GridDataEntity entity = getListData(usepage);
            event.setTable(entity.getData());
            onGetDataed(event);
            dt = null != event.getTable() ? event.getTable() : entity.getData();

            writeTitle(sheet1, dt);

            writeContent(sheet1, dt);

        }  catch (APPErrorException e) {
            throw new APPErrorException("获取表内容出错", e);
        } finally {
            ExcelCompletedEvent event =new ExcelCompletedEvent();
            event.setWorkbook(m_Workbook);
            event.setTable(dt);
            onCompleted(event);
        }
    }




}
