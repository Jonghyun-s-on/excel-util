package study.son.excelutil.sample.dto;

import lombok.Builder;
import lombok.Data;

import java.time.LocalDateTime;

@Data
@Builder
public class ExcelDownloadDto {
    private int number;
    private String content;
    private LocalDateTime createdDateTime;
}
