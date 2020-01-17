package tgtools.excel.jxl;


import jxl.Workbook;
import jxl.write.WritableWorkbook;
import tgtools.exceptions.APPErrorException;

import java.io.InputStream;


/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：17:02
 */
public class WorkbookFactory {

    public final static String EXCEL_TYPE_XLS = "xls";

    /**
     * 创建excel对象
     *
     * @param pInputStream
     * @return
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(InputStream pInputStream, String pVersion) throws APPErrorException {
        try {
            return Workbook.getWorkbook(pInputStream);
        } catch (Exception e) {
            throw new APPErrorException("加载EXCEL数据失败；原因：" + e.getMessage(), e);
        }
    }

}