package study.son.excelutil.sample.controller;

import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import study.son.excelutil.sample.service.SampleDownloadService;
import java.io.ByteArrayOutputStream;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

@RestController
public class SampleRestController {
    private final SampleDownloadService sampleDownloadService;

    public SampleRestController(SampleDownloadService sampleDownloadService) {
        this.sampleDownloadService = sampleDownloadService;
    }

    @PostMapping("/samples/download")
    public ResponseEntity<ByteArrayOutputStream> downloadExcel() {
        ByteArrayOutputStream byteArrayOutputStream = sampleDownloadService.downloadExcel();
        return ResponseEntity.ok().headers(createHeaders()).body(byteArrayOutputStream);
    }

    private HttpHeaders createHeaders() {
        HttpHeaders headers = new HttpHeaders();
        ContentDisposition contentDisposition = ContentDisposition.builder("attachment")
                .filename(URLEncoder.encode("downloadFilename", StandardCharsets.UTF_8))
                .build()
                ;
        headers.setContentDisposition(contentDisposition);
        return headers;
    }
}
