package study.son.excelutil.sample.service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import study.son.excelutil.sample.dto.ExcelDownloadDto;
import study.son.excelutil.util.excel.writer.ExcelWriter;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

@Service
public class SampleDownloadService {
    public ByteArrayOutputStream downloadExcel() {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        List<ExcelDownloadDto> data = getSampleData();
        ExcelWriter writer = new ExcelWriter(ExcelDownloadDto.class);
        XSSFWorkbook workbook = writer.writeExcel(new XSSFWorkbook(), data);
        try {
            workbook.write(byteArrayOutputStream);
            byteArrayOutputStream.flush();
            byteArrayOutputStream.close();
        } catch (IOException e) {
            e.getLocalizedMessage();
        }
        return byteArrayOutputStream;
    }

    private List<ExcelDownloadDto> getSampleData() {
        List<ExcelDownloadDto> dataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            dataList.add(ExcelDownloadDto.builder()
                    .number(i)
                    .content("content")
                    .createdDateTime(LocalDateTime.now())
                    .build()
            );
        }
        return dataList;
    }
}
