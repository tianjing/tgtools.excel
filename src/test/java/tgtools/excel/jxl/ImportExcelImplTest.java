package tgtools.excel.jxl;

import jxl.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Test;
import tgtools.excel.listener.ImportListener;
import tgtools.excel.listener.event.*;

import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashMap;

import static org.junit.Assert.*;

public class ImportExcelImplTest {


    @Test
    public void importExcel_File_Test() {
        String filepath = "C:\\Users\\tian_\\Desktop\\vcda.xls";
        ImportExcelImpl importExcel = new ImportExcelImpl();
        LinkedHashMap<String, String> columns = new LinkedHashMap<String, String>();
        columns.put("A", "A");
        columns.put("B", "B");
        columns.put("C", "C");
        columns.put("D", "D");
        columns.put("E", "E");

        HashMap<String, String> table = new HashMap<String, String>();
        table.put("Sheet1", "MQ_SYS.ACT_ID_USER");
        //默认不做数据库操作 之转换成json

        importExcel.init(columns, table,"dm",0,1);
        importExcel.setListener(new ImportListener(){

            @Override
            public void onCreateWorkbook(CreateWorkbookEvent pEvent) {
                System.out.println("");
            }

            @Override
            public void onCompleted(ExcelCompletedEvent pEvent) {
                System.out.println("");
            }

            @Override
            public void onLoadFilter(ImportEvent pEvent) {
                System.out.println("没有实现");
            }

            @Override
            public void onGetAtted(ImportEvent pEvent) {
                System.out.println("没有实现");
            }

            @Override
            public void onGetValue(ImportEvent pEvent) {
                if (((Sheet) pEvent.getSheet()).getRowView(pEvent.getRowIndex()).isHidden()) {
                    System.out.println(pEvent.getRowIndex() + "行被隐藏");
                }

                if (((Sheet) pEvent.getSheet()).getColumnView(pEvent.getColumnIndex()).isHidden()) {
                    System.out.println(pEvent.getColumnIndex() + "列被隐藏");
                }

                System.out.println();
            }

            @Override
            public void onExcuteSQL(ImportEvent pEvent) {
                pEvent.setIsExcute(false);
            }

            @Override
            public void onReadSheet(ReadSheetEvent pEvent) {
                System.out.println("");
            }

            @Override
            public void onSheetParsed(SheetParsedEvent pEvent) {
                System.out.println("");
            }
        });
        //设置数据库类型后进行sql 操作
        //importExcel.init(columns, table,"dm");
        try {
            importExcel.importExcel(new File(filepath));
            // Map<String, ArrayNode> ds = importExcel.getParseData();
            //importExcel.parseExcel();
            importExcel.close();
            System.out.println("");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}