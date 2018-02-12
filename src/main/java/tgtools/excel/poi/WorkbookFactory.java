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

    public final static String EXCEL_TYPE_XLS = "xls";
    public final static String EXCEL_TYPE_XLSX = "xlsx";

    /**
     * 创建excel对象
     *
     * @param pType
     *
     * @return
     */
    public static Workbook createWorkbook(String pType) {
        if (pType.toLowerCase().contains(EXCEL_TYPE_XLSX)) {
            return new XSSFWorkbook();
        } else if (pType.toLowerCase().contains(EXCEL_TYPE_XLS)) {
            return new HSSFWorkbook();
        }
        return null;
    }

    /**
     * 创建excel对象
     *
     * @param pFile
     *
     * @return
     *
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(File pFile) throws APPErrorException {
        String type = FileUtil.getExtensionName(pFile.getName());
        try {
            if ("xls".equals(type.toLowerCase())) {
                return new HSSFWorkbook(new FileInputStream(pFile));
            } else if ("xlsx".equals(type.toLowerCase())) {
                return new XSSFWorkbook(pFile);
            }
        } catch (Exception e) {
            throw new APPErrorException("加载文件失败；文件路径：" + pFile.getAbsolutePath() + ";原因：" + e.getMessage(), e);
        }

        return null;
    }

    /**
     * 创建excel对象
     *
     * @param pInputStream
     *
     * @return
     *
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(InputStream pInputStream, String pVersion) throws APPErrorException {
        if (EXCEL_TYPE_XLS.equals(pVersion) || ("." + EXCEL_TYPE_XLS).equals(pVersion)) {
            try {
                return new HSSFWorkbook(pInputStream);
            } catch (Exception e) {
                throw new APPErrorException("加载EXCEL数据失败；原因：" + e.getMessage(), e);
            }
        }
        if (EXCEL_TYPE_XLSX.equals(pVersion) || ("." + EXCEL_TYPE_XLSX).equals(pVersion)) {
            try {
                return new XSSFWorkbook(pInputStream);
            } catch (Exception e) {
                throw new APPErrorException("加载EXCEL数据失败；原因：" + e.getMessage(), e);
            }
        }
        throw new APPErrorException("无法识别的excel 版本；pVersion：" + (null==pVersion?"null":pVersion));
    }



    /**
     * 创建excel对象
     *
     * @param pDatas
     *
     * @return
     *
     * @throws APPErrorException
     */
    public static Workbook createWorkbook(byte[] pDatas, String pVersion) throws APPErrorException {
        ByteArrayInputStream dd = new ByteArrayInputStream(pDatas);
        return createWorkbook(dd,pVersion);
    }
}