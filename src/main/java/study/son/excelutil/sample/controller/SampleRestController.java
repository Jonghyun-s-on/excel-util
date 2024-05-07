package study.son.excelutil.sample.controller;

import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import study.son.excelutil.sample.service.SampleDownloadService;

import java.io.ByteArrayOutputStream;

@RestController
public class SampleRestController {
    private final SampleDownloadService sampleDownloadService;

    public SampleRestController(SampleDownloadService sampleDownloadService) {
        this.sampleDownloadService = sampleDownloadService;
    }

    @PostMapping("/samples/download")
    public ResponseEntity<byte[]> downloadExcel() {
        ByteArrayOutputStream byteArrayOutputStream = sampleDownloadService.downloadExcel();
        return ResponseEntity.ok().headers(createHeaders()).body(byteArrayOutputStream.toByteArray());
    }

    private HttpHeaders createHeaders() {
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDispositionFormData("attachment", "example.xlsx");
        return headers;
    }
}
