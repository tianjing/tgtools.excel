package tgtools.excel.util;

/**
 * @author 田径
 * @date 2020-04-13 11:40
 * @desc
 **/
public class SheetUtils {

    public static boolean isRowHidden(Object pSheet, int pRowIndex) {
        if (null == pSheet || pRowIndex < 0) {
            return false;
        }
        if (pSheet instanceof org.apache.poi.ss.usermodel.Sheet) {
            return ((org.apache.poi.ss.usermodel.Sheet) pSheet).getRow(pRowIndex).getZeroHeight();
        }
        if (pSheet instanceof jxl.Sheet) {
            return ((jxl.Sheet) pSheet).getRowView(pRowIndex).isHidden();
        }
        return false;
    }

    public static boolean isColumnHidden(Object pSheet, int pColumnIndex) {
        if (pSheet instanceof org.apache.poi.ss.usermodel.Sheet) {
            return ((org.apache.poi.ss.usermodel.Sheet) pSheet).getColumnWidth(pColumnIndex) < 1;
        }
        if (pSheet instanceof jxl.Sheet) {
            return ((jxl.Sheet) pSheet).getColumnView(pColumnIndex).isHidden();
        }

        return false;
    }
}
