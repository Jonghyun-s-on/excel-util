package study.son.excelutil.util.excel.writer;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import study.son.excelutil.util.excel.ExcelMetaData;

import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class ExcelWriter {
    private String NEW_SHEET_NAME = "DATA";
    private ExcelMetaData excelMetaData;
    private List<String> fieldNames = new ArrayList<>();
    private DateTimeFormatter DATE_TIME_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public <T> ExcelWriter(Class<T> type) {
        this.excelMetaData = new ExcelMetaData(type);
    }

    public <T> XSSFWorkbook writeExcel(XSSFWorkbook workbook, List<T> data) {
        this.fieldNames = excelMetaData.getFieldNames();
        XSSFSheet sheet = workbook.createSheet(NEW_SHEET_NAME);
        sheet = renderHeader(sheet);

        for (T d : data) {
            XSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
            for (int i = 0; i < this.fieldNames.size(); i++) {
                XSSFCell cell = row.createCell(i);
                try {
                    Optional<Field> field = this.excelMetaData.getField(this.fieldNames.get(i));
                    if (field.isPresent()) {
                        Field f = field.get();
                        f.setAccessible(true);
                        setValueIntoCell(cell, d, f);
                    }
                } catch (NullPointerException ignore) {
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        return workbook;
    }

    private <T> XSSFCell setValueIntoCell(XSSFCell cell, T d, Field f) throws IllegalAccessException {
        if (d != null) {
            if (f.getType() == int.class) {
                cell.setCellValue(f.getInt(d));
            } else if (f.getType() == LocalDateTime.class) {
                LocalDateTime dateTime = (LocalDateTime) f.get(d);
                cell.setCellValue(dateTime.format(DATE_TIME_FORMAT));
            } else {
                cell.setCellValue((String) f.get(d));
            }
        }
        return cell;
    }

    private XSSFSheet renderHeader(XSSFSheet sheet) {
        XSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
        for (int i = 0; i < this.fieldNames.size(); i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellValue(this.fieldNames.get(i));
        }
        return sheet;
    }

}
