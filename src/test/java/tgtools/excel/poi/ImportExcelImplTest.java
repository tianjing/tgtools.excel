package tgtools.excel.poi;

import com.fasterxml.jackson.databind.node.ArrayNode;
import org.junit.Test;
import tgtools.excel.listener.ImportListener;
import tgtools.excel.listener.event.*;

import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author 田径
 * @Title
 * @Description
 * @date 9:39
 */
public class ImportExcelImplTest {

    @Test
    public void importExcel_File_Test() {
        String filepath = "C:\\Users\\tian_\\Desktop\\报修表.xls";
        ImportExcelImpl importExcel =new ImportExcelImpl();
        LinkedHashMap<String,String> columns=new LinkedHashMap<String,String>();
        columns.put("ID","地址");
        columns.put("NAME","序号");
        columns.put("SEX","间隔");
        columns.put("BIR","设备");
        columns.put("BIR","遥信名称");
        columns.put("BIR","信息分类");
        columns.put("BIR","硬接点");
        columns.put("BIR","备注");

        HashMap<String,String> table=new HashMap<String,String>();
        table.put("遥信","MQ_SYS.ACT_ID_USER");
        importExcel.init(columns,table);
        try {
            importExcel.setListener(new ImportListener(){
                @Override
                public void onCreateWorkbook(CreateWorkbookEvent p_Event) {

                }

                @Override
                public void onCompleted(ExcelCompletedEvent p_Event) {

                }

                @Override
                public void onLoadFilter(ImportEvent p_Event) {

                }

                @Override
                public void onGetAtted(ImportEvent p_Event) {

                }

                @Override
                public void onGetValue(ImportEvent p_Event) {

                }

                @Override
                public void onExcuteSQL(ImportEvent p_Event) {
                    try {
                        System.out.println("sql:"+p_Event.getSql());
                    }catch (Exception ex)
                    {}
                    p_Event.setIsExcute(false);
                }

                @Override
                public void onReadSheet(ReadSheetEvent p_Event) {
                    p_Event.setCancel(!"遥信".equals(p_Event.getSheetName()));
                }

                @Override
                public void onSheetParsed(SheetParsedEvent p_Event) {

                }
            });
            importExcel.importExcel(new File(filepath));
            Map<String,ArrayNode> ds= importExcel.getParseData();
            importExcel.close();
            System.out.println(ds);
        }catch (Exception e)
        {
            e.printStackTrace();
        }


    }
}