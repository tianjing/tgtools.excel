package tgtools.excel.poi;

import com.fasterxml.jackson.databind.node.ArrayNode;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.Test;
import tgtools.excel.listener.ImportListener;
import tgtools.excel.listener.event.*;
import tgtools.excel.util.SheetUtils;

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
        String filepath = "C:\\Users\\tian_\\Desktop\\vcda.xlsx";
        ImportExcelImpl importExcel = new ImportExcelImpl();
        LinkedHashMap<String, String> columns = new LinkedHashMap<String, String>();
        columns.put("A", "A");
        columns.put("B", "B");
        columns.put("C", "C");
        columns.put("D", "D");
        columns.put("E", "E");

        HashMap<String, String> table = new HashMap<String, String>();
        table.put("Sheet1", "MQ_SYS.ACT_ID_USER");
        importExcel.init(columns, table);
        try {
            importExcel.setListener(new ImportListener() {
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
                    System.out.println();
                }

                @Override
                public void onGetValue(ImportEvent p_Event) {
                    if (SheetUtils.isRowHidden(p_Event.getSheet(),p_Event.getRowIndex())) {
                        System.out.println(p_Event.getRowIndex() + "行被隐藏");
                    }

                    if (SheetUtils.isRowHidden(p_Event.getSheet(),p_Event.getColumnIndex())) {
                        System.out.println(p_Event.getColumnIndex() + "列被隐藏");
                    }

                    System.out.println();
                }

                @Override
                public void onExcuteSQL(ImportEvent p_Event) {
                    try {
                        System.out.println("sql:" + p_Event.getSql());
                    } catch (Exception ex) {
                    }
                    p_Event.setIsExcute(false);
                }

                @Override
                public void onReadSheet(ReadSheetEvent p_Event) {
                    // p_Event.setCancel(!"遥信".equals(p_Event.getSheetName()));
                }

                @Override
                public void onSheetParsed(SheetParsedEvent p_Event) {

                }
            });
            importExcel.importExcel(new File(filepath));
            Map<String, ArrayNode> ds = importExcel.getParseData();
            importExcel.close();
            System.out.println(ds);
        } catch (Exception e) {
            e.printStackTrace();
        }


    }
}