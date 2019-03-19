package tgtools.excel;


import com.fasterxml.jackson.databind.node.ArrayNode;
import tgtools.excel.listener.ExportListener;
import tgtools.exceptions.APPErrorException;

import java.io.Closeable;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.LinkedHashMap;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：9:38
 */
public interface ExportExcel extends Closeable {
    public static final String VERSION_EXCEL2003 = "xls";
    public static final String VERSION_EXCEL2007 = "xlsx";


    /**
     * 初始化参数
     *
     * @param pVersion Excel版本 参考 ExportExcel中 VERSION_EXCEL2003 VERSION_EXCEL2007
     *
     * @throws APPErrorException
     */
    void init(String pVersion) throws APPErrorException;
    /**
     * 初始化参数
     *
     * @param pFile Excel文件
     *
     * @throws APPErrorException
     */
    void init(File pFile) throws APPErrorException;

    /**
     * 初始化参数
     *
     * @param pInputStream Excel文件
     * @param pVersion  xls xlsx
     *
     * @throws APPErrorException
     */
    void init(InputStream pInputStream,String pVersion) throws APPErrorException;

    /**
     * 初始化参数
     *
     * @param pBytes Excel文件
     * @param pVersion  xls xlsx
     *
     * @throws APPErrorException
     */
    void init(byte[] pBytes,String pVersion) throws APPErrorException;



    /**
     * 设置自定义监听 可进行功能扩展
     *
     * @param pLitener
     *
     * @throws APPErrorException
     */
    void setLisener(ExportListener pLitener) throws APPErrorException;

    /**
     * 通过json 追加数据到默认"sheet1", 0行表头，1行开始写入数据） 中
     *
     * @param pColumns 列信息  key为列英文 value为别名或中文 Map中key的顺序为excel导出数据的顺序（作为excel表头和json数据匹配中）
     * @param pJson
     *
     * @throws APPErrorException
     */
    void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson) throws APPErrorException;

    /**
     * 通过json 追加数据到默认"sheet1", 0行表头，1行开始写入数据） 中
     *
     * @param pColumns        列信息  key为列英文 value为别名或中文 Map中key的顺序为excel导出数据的顺序（作为excel表头和json数据匹配中）
     * @param pJson
     * @param pIsExportTitle 是否写入表头
     *
     * @throws APPErrorException
     */
    void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle) throws APPErrorException;

    /**
     * 向指定excel追加数据
     *
     * @param pColumns        列信息  key为列英文 value为别名或中文 Map中key的顺序为excel导出数据的顺序（作为excel表头和json数据匹配中）
     * @param pJson          需要追加的数据
     * @param pIsExportTitle 是否写入表头
     * @param pSheetName     excel sheet 名称（可null）
     * @param pSheetIndex    excel 新建sheet时使用的索引
     * @param pDataIndex     （从第几行增加数据,默认是1 即 标题行的下一行）
     *
     * @throws APPErrorException
     */
    void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle, String pSheetName, int pSheetIndex, int pDataIndex) throws APPErrorException;

    /**
     * 获取excel数据 同时 释放excel对象
     *
     * @return 获取 excel 所有内容
     *
     * @throws APPErrorException
     */
    byte[] getBytes() throws APPErrorException;

    /**
     * 获取 Excel 流对象 使用完 请调用 close 方法释放
     *
     * @return 获取 excel流 对象
     *
     * @throws APPErrorException
     */
    OutputStream getOutputStream() throws APPErrorException;

    /**
     * 获取 Excel 对象  使用完 请调用 close 方法释放
     *
     * @return
     */
    Object getExcel();
}
