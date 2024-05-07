package study.son.excelutil.sample.dto;

import lombok.Builder;
import lombok.Getter;
import study.son.excelutil.util.excel.ExcelColumn;
import java.time.LocalDateTime;

@Getter
@Builder
public class ExcelDownloadDto {
    @ExcelColumn(fieldName = "number")
    private int number;
    @ExcelColumn(fieldName = "content")
    private String content;
    @ExcelColumn(fieldName = "createdDateTime")
    private LocalDateTime createdDateTime;
}
