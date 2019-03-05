package tgtools.excel.poi;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.junit.Test;
import tgtools.util.FileUtil;

import java.io.File;
import java.math.BigDecimal;
import java.util.LinkedHashMap;

import static org.junit.Assert.*;

/**
 * @author 田径
 * @Title
 * @Description
 * @date 8:38
 */
public class ExportExcelImplTest {


    @Test
    public void ExportXlsToByteTest()
    {
        try {
            String filepath = "C:\\Users\\tian_\\Desktop\\221.xls";
            String outfilepath = "C:\\Users\\tian_\\Desktop\\222.xls";
            ExportExcelImpl export = new ExportExcelImpl();
            export.init(new File(filepath));
            LinkedHashMap<String,String> columns=new LinkedHashMap<String,String>();
            columns.put("ID","主键");
            columns.put("NAME","名称");
            columns.put("SEX","性别");
            columns.put("BIR","生日");

            ObjectMapper mapper =new ObjectMapper();
            ArrayNode array= mapper.createArrayNode();
            ObjectNode json= mapper.createObjectNode();
            json.put("ID",new Integer(1));
            json.put("ID1",new Double(1232.321d));
            json.put("ID2",new BigDecimal(12321.012321321f));

            json.put("NAME","田径1");
            json.put("SEX","男");
            json.put("BIR","2013-12:12 12:00:00");
            array.add(json);
            ObjectNode json1= mapper.createObjectNode();
            json1.put("ID",2);
            json1.put("NAME","田径2");
            json1.put("SEX","女");
            json1.put("BIR","2014-12:12 12:00:00");
            array.add(json1);


            export.appendData(columns,array, false, "sheet1", 0, 3);
            byte[] data= export.getBytes();
            export.close();
            FileUtil.writeFile(outfilepath,data);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    @Test
    public void ExportXlsToFileTest()
    {
        try {
            String filepath = "C:\\tianjing\\Desktop\\222.xls";
            ExportExcelImpl export = new ExportExcelImpl();
            export.init(WorkbookFactory.EXCEL_TYPE_XLS);
            LinkedHashMap<String,String> columns=new LinkedHashMap<String,String>();
            columns.put("ID","主键");
            columns.put("NAME","名称");
            columns.put("SEX","性别");
            columns.put("BIR","生日");

            ObjectMapper mapper =new ObjectMapper();
            ArrayNode array= mapper.createArrayNode();
            ObjectNode json= mapper.createObjectNode();
            json.put("ID",1);
            json.put("NAME","田径1");
            json.put("SEX","男");
            json.put("BIR","2013-12:12 12:00:00");
            array.add(json);
            ObjectNode json1= mapper.createObjectNode();
            json1.put("ID",2);
            json1.put("NAME","田径2");
            json1.put("SEX","女");
            json1.put("BIR","2014-12:12 12:00:00");
            array.add(json1);
            export.appendData(columns,array);
            export.appendData(columns,array,false,"sheet1",0,2);

            byte[] data= export.getBytes();
            export.close();
            FileUtil.writeFile(filepath,data);
            System.out.println("结束");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
    @Test
    public void ExportXlsxToByteTest()
    {
        try {
            String filepath = "C:\\tianjing\\Desktop\\221.xlsx";
            ExportExcelImpl export = new ExportExcelImpl();
            export.init(WorkbookFactory.EXCEL_TYPE_XLSX);
            LinkedHashMap<String,String> columns=new LinkedHashMap<String,String>();
            columns.put("ID","主键");
            columns.put("NAME","名称");
            columns.put("SEX","性别");
            columns.put("BIR","生日");

            ObjectMapper mapper =new ObjectMapper();
            ArrayNode array= mapper.createArrayNode();
            ObjectNode json= mapper.createObjectNode();
            json.put("ID",1);
            json.put("NAME","田径1");
            json.put("SEX","男");
            json.put("BIR","2013-12:12 12:00:00");
            array.add(json);
            ObjectNode json1= mapper.createObjectNode();
            json1.put("ID",2);
            json1.put("NAME","田径2");
            json1.put("SEX","女");
            json1.put("BIR","2014-12:12 12:00:00");
            array.add(json1);


            export.appendData(columns,array);
            byte[] data= export.getBytes();
            export.close();
            FileUtil.writeFile(filepath,data);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    @Test
    public void ExportXlsxToFileTest()
    {
        try {
            String filepath = "C:\\tianjing\\Desktop\\222.xlsx";
            ExportExcelImpl export = new ExportExcelImpl();
            export.init(WorkbookFactory.EXCEL_TYPE_XLSX);
            LinkedHashMap<String,String> columns=new LinkedHashMap<String,String>();
            columns.put("ID","主键");
            columns.put("NAME","名称");
            columns.put("SEX","性别");
            columns.put("BIR","生日");

            ObjectMapper mapper =new ObjectMapper();
            ArrayNode array= mapper.createArrayNode();
            ObjectNode json= mapper.createObjectNode();
            json.put("ID",1);
            json.put("NAME","田径1");
            json.put("SEX","男");
            json.put("BIR","2013-12:12 12:00:00");
            array.add(json);
            ObjectNode json1= mapper.createObjectNode();
            json1.put("ID",2);
            json1.put("NAME","田径2");
            json1.put("SEX","女");
            json1.put("BIR","2014-12:12 12:00:00");
            array.add(json1);
            export.appendData(columns,array);
            export.appendData(columns,array,false,"sheet1",0,2);

            byte[] data= export.getBytes();
            export.close();
            FileUtil.writeFile(filepath,data);
            System.out.println("结束");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}