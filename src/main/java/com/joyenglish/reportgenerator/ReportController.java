package com.joyenglish.reportgenerator;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/api/report")
public class ReportController {

    @Autowired
    private ReportService reportService;

    @PostMapping(value = "/generate", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> generate(
            @RequestParam("targetMonth") String targetMonth,
            @RequestParam(value = "projectsJson", required = false) String projectsJson,
            @RequestParam("excelMain") MultipartFile excelMain,
            @RequestParam("excelEnrollment") MultipartFile excelEnrollment,
            @RequestParam("excelHistory") MultipartFile excelHistory
    ) {
        try {
            // 防呆處理：沒傳、傳 null、傳空字串皆預設為空 List，絕不中斷
            List<ProjectReportRequest> projects = new ArrayList<>();
            if (projectsJson != null && !projectsJson.trim().isEmpty() && !projectsJson.trim().equals("null")) {
                ObjectMapper mapper = new ObjectMapper();
                projects = mapper.readValue(
                    projectsJson, new TypeReference<List<ProjectReportRequest>>(){}
                );
            }

            byte[] pdfBytes = reportService.generateReportPdf(
                targetMonth, projects, excelMain, excelEnrollment, excelHistory
            );

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_PDF);
            headers.setContentDispositionFormData("filename", "MonthlyReport_" + targetMonth + ".pdf");

            return ResponseEntity.ok().headers(headers).body(pdfBytes);

        } catch (Exception e) {
            e.printStackTrace();
            String errorHtml = "<!DOCTYPE html><html lang='zh-TW'><head><meta charset='UTF-8'></head>" +
                               "<body style='font-family: sans-serif; padding: 30px; line-height: 1.6;'>" +
                               "<h2 style='color:#e74c3c;'>⚠️ 系統發生錯誤 (HTTP 500)</h2>" +
                               "<h3 style='color:#333; background:#f8d7da; padding:10px; border-radius:5px;'>錯誤原因：" + e.getMessage() + "</h3>" +
                               "</body></html>";
            
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .header(HttpHeaders.CONTENT_TYPE, "text/html; charset=UTF-8")
                    .body(errorHtml.getBytes(java.nio.charset.StandardCharsets.UTF_8));
        }
    }
}