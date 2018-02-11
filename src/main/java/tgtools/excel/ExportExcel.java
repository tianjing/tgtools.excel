package tgtools.excel;


import com.fasterxml.jackson.databind.node.ArrayNode;
import tgtools.excel.Listener.ExportListener;

import java.io.Closeable;
import java.io.File;
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
     * @throws Exception
     */
    void init(String pVersion) throws Exception;

    /**
     * 初始化参数
     *
     * @param pFile 文件对象 根据文件的扩展名 判断excel 版本
     *
     * @throws Exception
     */
    void init(File pFile) throws Exception;

    /**
     * 初始化参数
     *
     * @param pVersion      Excel版本 参考 ExportExcel中 VERSION_EXCEL2003 VERSION_EXCEL2007
     * @param pOutputStream excel写入时保存的流对象
     *
     * @throws Exception
     */
    void init(String pVersion, OutputStream pOutputStream) throws Exception;

    /**
     * 设置自定义监听 可进行功能扩展
     *
     * @param pLitener
     *
     * @throws Exception
     */
    void setLisener(ExportListener pLitener) throws Exception;

    /**
     * 通过json 追加数据到默认"sheet1", 0行表头，1行开始写入数据） 中
     *
     * @param pColumns 列信息  key为列英文 value为别名或中文 Map中key的顺序为excel导出数据的顺序（作为excel表头和json数据匹配中）
     * @param pJson
     *
     * @throws Exception
     */
    void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson) throws Exception;

    /**
     * 通过json 追加数据到默认"sheet1", 0行表头，1行开始写入数据） 中
     *
     * @param pColumns        列信息  key为列英文 value为别名或中文 Map中key的顺序为excel导出数据的顺序（作为excel表头和json数据匹配中）
     * @param pJson
     * @param pIsExportTitle 是否写入表头
     */
    void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle) throws Exception;

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
     * @throws Exception
     */
    void appendData(LinkedHashMap<String, String> pColumns, ArrayNode pJson, boolean pIsExportTitle, String pSheetName, int pSheetIndex, int pDataIndex) throws Exception;

    /**
     * 获取excel数据 同时 释放excel对象
     *
     * @return 获取 excel 所有内容
     *
     * @throws Exception
     */
    byte[] getBytes() throws Exception;

    /**
     * 获取 Excel 流对象 使用完 请调用 close 方法释放
     *
     * @return 获取 excel流 对象
     *
     * @throws Exception
     */
    OutputStream getOutputStream() throws Exception;

    /**
     * 获取 Excel 对象  使用完 请调用 close 方法释放
     *
     * @return
     */
    Object getExcel();
}
