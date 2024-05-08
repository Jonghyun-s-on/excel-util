package study.son.excelutil.util.excel.writer;

import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import study.son.excelutil.util.excel.ExcelMetaData;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class ExcelSXSSWriter {
    private String NEW_SHEET_NAME = "DATA";
    private ExcelMetaData excelMetaData;
    private List<String> fieldNames = new ArrayList<>();
    private DateTimeFormatter DATE_TIME_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public <T> ExcelSXSSWriter(Class<T> type) {
        this.excelMetaData = new ExcelMetaData(type);
    }

    public <T> void writeExcel(FileOutputStream fos, List<T> data) {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        workbook.setCompressTempFiles(true);

        if (workbook.getNumberOfSheets() == 0) {
            this.fieldNames = this.excelMetaData.getFieldNames();
        }

        SXSSFSheet sheet = workbook.createSheet(NEW_SHEET_NAME);
        sheet = renderHeader(sheet);

        sheet.setRandomAccessWindowSize(100);

        for (T d : data) {
            SXSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
            for (int i = 0; i < this.fieldNames.size(); i++) {
                SXSSFCell cell = row.createCell(i);
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

        try {
            workbook.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
                workbook.dispose();
                if (fos != null) {
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private <T> SXSSFCell setValueIntoCell(SXSSFCell cell, T d, Field f) throws IllegalAccessException {
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

    private SXSSFSheet renderHeader(SXSSFSheet sheet) {
        SXSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
        for (int i = 0; i < this.fieldNames.size(); i++) {
            SXSSFCell cell = row.createCell(i);
            cell.setCellValue(this.fieldNames.get(i));
        }
        return sheet;
    }
}
