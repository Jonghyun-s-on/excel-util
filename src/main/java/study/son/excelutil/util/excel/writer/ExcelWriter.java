package study.son.excelutil.util.excel.writer;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public class ExcelWriter {
    public <T> ExcelWriter(Class<T> type) {

    }

    public <T> XSSFWorkbook writeExcel(XSSFWorkbook workbook, List<T> data) {
        return null;
    }
}
