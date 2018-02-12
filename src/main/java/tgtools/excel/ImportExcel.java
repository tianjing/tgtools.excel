package tgtools.excel;


import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import tgtools.excel.listener.ImportListener;
import tgtools.exceptions.APPErrorException;

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
    void setListener(ImportListener pLitener);
    /**
     * 初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     * @throws APPErrorException
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable) throws APPErrorException;

    /**
     * 初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     * @param pDatabaseType 数据库类型 如果 空或null 则不做sql处理
     * @throws APPErrorException
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable,String pDatabaseType) throws APPErrorException;

    /**
     *初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     * @param pDatabaseType 数据库类型 如果 空或null 则不做sql处理
     * @param pBeginDataRow 起始行
     * @throws APPErrorException
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable,String pDatabaseType, int pBeginDataRow)throws APPErrorException;


    /**
     * 初始化
     * @param pColumns 表列名
     * @param pSheetTable sheet 名称与表对应的名称 如：put("sheet1","MQ_SYS.ACT_ID_USER")
     * @param pDatabaseType 数据库类型 如果 空或null 则不做sql处理
     * @param pBeginTitleRow 标题行
     * @param pBeginDataRow 数据起始行
     * @throws APPErrorException
     */
    void init(LinkedHashMap<String,String> pColumns, Map<String,String> pSheetTable,String pDatabaseType,int pBeginTitleRow,int pBeginDataRow) throws APPErrorException;

    /**
     * 导入
     * @param pFile
     * @throws APPErrorException
     */
    void importExcel(File pFile) throws APPErrorException;

    /**
     * 导入
     * @param pDatas
     * @param pVersion excel 版本
     * @throws APPErrorException
     */
    void importExcel(byte[] pDatas,String pVersion) throws APPErrorException;
    /**
     * 导入
     * @param pInputStream
     * @param pVersion excel 版本
     * @throws APPErrorException
     */
    void importExcel(InputStream pInputStream,String pVersion) throws APPErrorException;

    /**
     * 获取入库结果字符串 如 ：{"success":30,"error":10,"count":40}
     * @return
     * @throws APPErrorException
     */
    ObjectNode getImportResult()throws APPErrorException;
    /**
     * 获取转换后的结果的副本
     * @return 每一个ArrayNode 就是一个sheet 的数据
     * @throws APPErrorException
     */
    LinkedHashMap<String,ArrayNode> getParseData()throws APPErrorException;

}
