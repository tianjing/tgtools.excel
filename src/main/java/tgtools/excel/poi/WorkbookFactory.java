package tgtools.excel.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tgtools.exceptions.APPErrorException;
import tgtools.util.FileUtil;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;


/**
 * 名  称：
 * 编写者：田径
 * 功  能：
 * 时  间：17:02
 */
public class WorkbookFactory {

    public final static String EXCEL_TYPE_XLS="xls";
    public final static String EXCEL_TYPE_XLSX="xlsx";
    /**
     * 创建excel对象
     * @param p_Type
     * @return
     */
    public static Workbook createWorkbook(String p_Type) {
        if (p_Type.toLowerCase().contains(EXCEL_TYPE_XLSX)) {
            return new XSSFWorkbook();
        }
        else if (p_Type.toLowerCase().contains(EXCEL_TYPE_XLS)) {
            return  new HSSFWorkbook();
        }
        return null;
    }

    /**
     * 创建excel对象
     * @param p_File
     * @return
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(File p_File) throws APPErrorException {
        String type = FileUtil.getExtensionName(p_File.getName());
        try {
            if ("xls".equals(type.toLowerCase())) {
                return new HSSFWorkbook(new FileInputStream(p_File));
            } else if ("xlsx".equals(type.toLowerCase())) {
                return new XSSFWorkbook(p_File);
            }
        } catch (Exception e) {
            throw new APPErrorException("加载文件失败；文件路径：" + p_File.getAbsolutePath() + ";原因：" + e.getMessage(), e);
        }

        return null;
    }

    /**
     * 创建excel对象
     * @param p_InputStream
     * @return
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(InputStream p_InputStream) throws APPErrorException {
        try {
            return new HSSFWorkbook(p_InputStream);
        }catch (Exception e)
        {
            try {
                p_InputStream.reset();
                return  new XSSFWorkbook(p_InputStream);
            } catch (Exception e1) {
                throw new APPErrorException("加载EXCEL数据失败；原因：" + e.getMessage(), e);
            }
        }
    }

    /**
     * 创建excel对象
     * @param p_Datas
     * @return
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(byte[] p_Datas) throws APPErrorException {
        ByteArrayInputStream dd=new ByteArrayInputStream(p_Datas);
        return createWorkbook(dd);
    }
}