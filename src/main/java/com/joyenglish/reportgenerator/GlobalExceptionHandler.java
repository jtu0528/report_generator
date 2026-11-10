package com.joyenglish.reportgenerator;

import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.multipart.MaxUploadSizeExceededException;
import org.springframework.http.ResponseEntity;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;

@ControllerAdvice
public class GlobalExceptionHandler {

    @ExceptionHandler(MaxUploadSizeExceededException.class)
    public ResponseEntity<byte[]> handleMaxSizeException(MaxUploadSizeExceededException exc) {
        String errorHtml = "<meta charset='UTF-8'>" +
                           "<div style='font-family: sans-serif; padding: 30px;'>" +
                           "<h2 style='color:#e74c3c;'>⚠️ 上傳失敗：Excel 檔案太大了！</h2>" +
                           "<p>您上傳的檔案超過了系統預設的 1MB 限制 (115年統計明細通常有 1.2MB)。</p>" +
                           "<p>請確定您已經在 <b>src/main/resources/application.properties</b> 檔案中加入了以下設定：</p>" +
                           "<pre style='background:#eee; padding:15px; border-radius:5px;'>" +
                           "spring.servlet.multipart.max-file-size=10MB\n" +
                           "spring.servlet.multipart.max-request-size=30MB</pre>" +
                           "<p style='color:blue; font-weight:bold;'>👉 加入後，請務必重啟系統 (Ctrl+C 關閉，再重新 mvn clean package) 才會生效！</p>" +
                           "</div>";
        
        return ResponseEntity.status(HttpStatus.PAYLOAD_TOO_LARGE)
                .contentType(MediaType.TEXT_HTML)
                .body(errorHtml.getBytes());
    }
}