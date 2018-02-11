package tgtools.excel;


import com.fasterxml.jackson.databind.node.ObjectNode;
import tgtools.excel.Listener.ImportListener;

import java.io.Closeable;
import java.io.File;
import java.io.InputStream;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：21:47
 */
public interface ImportExcel extends Closeable {
    /**
     * 设置生成excel时监听
     * @param pLitener
     */
    void setListener(ImportListener pLitener) ;

    /**
     * 初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable) ;

    /**
     *初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     * @param p_BeginDataRow 起始行
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable, int p_BeginDataRow);


    /**
     * 初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     * @param p_BeginTitleRow 标题行
     * @param p_BeginDataRow 数据起始行
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable,int p_BeginTitleRow,int p_BeginDataRow);

    /**
     * 导入
     * @param pFile
     * @throws Exception
     */
    void importExcel(File pFile) throws Exception;

    /**
     * 导入
     * @param pDatas
     * @throws Exception
     */
    void importExcel(byte[] pDatas) throws Exception;
    /**
     * 导入
     * @param pInputStream
     * @throws Exception
     */
    void importExcel(InputStream pInputStream) throws Exception;

    /**
     * 获取结果字符串 如 ：{"success":30,"error":10,"count":40}
     * @return
     */
    ObjectNode getImportResult();


}
