package tgtools.excel;


import tgtools.excel.Listener.ExcelAllLitener;
import tgtools.excel.Listener.ExportLisener;
import tgtools.exceptions.APPErrorException;
import tgtools.interfaces.IDispose;

import java.io.Closeable;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:38
 */
public interface ExportExcel extends Closeable {

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
    void init(String[] p_Attr, String[] p_AttrNames, String p_tabName, String p_SortsStr, String p_SortOrder, String p_PageIndex, String p_PageSize, String p_CutPage, String p_Type);

    void setLisener(ExportLisener p_Litener);

    /**
     * 设置扩展类的监听
     * @param p_Litener
     */
    void setExcelAllLisener(ExcelAllLitener p_Litener) ;

    /**
     * 生成excel
     * @throws APPErrorException
     */
    void buildExcel() throws APPErrorException;
    /**
     * 生成excel
     *
     * @param p_sheetName
     * @param p_Index
     * @throws APPErrorException
     */
    void buildExcel(String p_sheetName, int p_Index) throws APPErrorException;

    /**
     * 获取excel数据同时释放excel对象
     * @return
     * @throws APPErrorException
     */
    byte[] getBytes() throws APPErrorException;

}
