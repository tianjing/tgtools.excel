package tgtools.excel;


import tgtools.excel.Listener.ExcelAllLitener;
import tgtools.excel.Listener.ImportLisener;
import tgtools.exceptions.APPErrorException;
import tgtools.interfaces.IDispose;

import java.io.Closeable;
import java.io.File;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：21:47
 */
public interface ImportExcel extends Closeable {
    /**
     * 设置生成excel时监听
     * @param p_Litener
     */
    void setLisener(ImportLisener p_Litener) ;

    /**
     * 设置扩展类监听
     * @param p_Litener
     */
    void setExcelAllLisener(ExcelAllLitener p_Litener) ;
    /**
     * 初始化
     * @param p_Column 表列名
     * @param p_ColumnName 列中文名
     * @param p_TableName 表名
     */
    void init(String[] p_Column, String[] p_ColumnName, String p_TableName) ;

    /**
     * 初始化
     * @param p_Column 表列名
     * @param p_ColumnName 列中文名
     * @param p_TableName 表名
     * @param p_BeginTitleRow 标题行
     * @param p_BeginDataRow 数据起始行
     */
    void init(String[] p_Column, String[] p_ColumnName, String p_TableName,int p_BeginTitleRow,int p_BeginDataRow);
    /**
     *初始化
     * @param p_Column 表列名
     * @param p_ColumnName 列中文名
     * @param p_TableName 表名
     * @param p_beginRow 起始行
     */
    void init(String[] p_Column, String[] p_ColumnName, String p_TableName, int p_beginRow);

    /**
     * 导入
     * @param p_File
     * @throws APPErrorException
     */
    void importExcel(File p_File) throws APPErrorException;

    /**
     * 导入
     * @param p_Datas
     * @throws APPErrorException
     */
    void importExcel(byte[] p_Datas) throws APPErrorException;

    /**
     * 获取结果字符串 如 ：100/200
     * @return
     */
    String getImportResult();


}
