package com.joyenglish.reportgenerator;

import com.openhtmltopdf.pdfboxout.PdfRendererBuilder;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.*;

@Service
public class ReportService {

    @Autowired
    private TemplateEngine templateEngine;

    public byte[] generateReportPdf(String targetMonth, List<ProjectReportRequest> projects,
                                    MultipartFile excelMain, MultipartFile excelEnrollment,
                                    MultipartFile excelHistory) throws Exception {

        Map<String, Object> dataModel = new HashMap<>();
        dataModel.put("reportTitle", "龍安/榮富/新昌隆 校務部工作月報告 115 年 " + targetMonth);
        dataModel.put("branchName", "新昌隆");

        // 0. 過濾專案報告：剔除未填寫、空字串或僅填寫「無」的項目
        List<ProjectReportRequest> validProjects = new ArrayList<>();
        if (projects != null) {
            for (ProjectReportRequest proj : projects) {
                if (proj != null && proj.getName() != null && !proj.getName().trim().isEmpty()) {
                    List<String> validItems = new ArrayList<>();
                    if (proj.getItems() != null) {
                        for (String item : proj.getItems()) {
                            if (item != null && !item.trim().isEmpty() && !"無".equals(item.trim())) {
                                validItems.add(item.trim());
                            }
                        }
                    }
                    if (!validItems.isEmpty()) {
                        proj.setItems(validItems);
                        validProjects.add(proj);
                    }
                }
            }
        }
        dataModel.put("projects", validProjects);

        // 學生人數報告指標（步驟 1 建立，步驟 2 再補上國高中/小六人數）
        Map<String, String> studentStats = new HashMap<>();

        // 1. 解析【綜合統計明細檔】(學生人數報告 + 流失原因明細)
        try (InputStream is = excelMain.getInputStream()) {
            Workbook wb = WorkbookFactory.create(is);

            // 流失名冊 (自動過濾標題列與空白列)
            String churnSheetName = targetMonth + "學生流失原因";
            Sheet churnSheet = wb.getSheet(churnSheetName);
            if (churnSheet == null) {
                throw new RuntimeException("匯出失敗：在統計檔中找不到「" + churnSheetName + "」工作表！");
            }
            List<Map<String, String>> churnList = readChurnList(churnSheet);
            dataModel.put("churnList", churnList);

            // 學生人數報告指標 (自動動態計算 TJ, MJ, ERE, BM)
            studentStats = readStudentStatsDynamically(wb, targetMonth, churnList.size());
        }

        // 2. 解析【招生人數 + 各年級統計】(100% 動態欄位架構，支援未來班級代號變更)
        try (InputStream is = excelEnrollment.getInputStream()) {
            Workbook wb = WorkbookFactory.create(is);

            // 2-1. 各年級統計：國高中人數、國小生人數、小六人數
            studentStats.putAll(readGradeStats(wb, targetMonth));

            // 2-2. 每月招生人數
            Map<String, Object> enrollmentData = readMonthlyEnrollmentDynamic(wb);
            dataModel.put("enrollmentHeaders", enrollmentData.get("headers"));
            dataModel.put("monthlyEnrollmentRows", enrollmentData.get("rows"));
            dataModel.put("enrollmentTotalRow", enrollmentData.get("totalRow"));
        }

        dataModel.put("studentStats", studentStats);

        // 3. 解析【歷年統計】(整合 109~113年 與 111~115年，完整涵蓋 110~115 年)
        try (InputStream is = excelHistory.getInputStream()) {
            Workbook wb = WorkbookFactory.create(is);
            Map<String, Object> historyData = readHistoricalStatsComplete(wb);
            dataModel.put("historicalStats", historyData.get("list"));
            dataModel.put("historicalAvg", historyData.get("avg"));
        }

        // 4. Thymeleaf 渲染 HTML
        Context context = new Context();
        context.setVariables(dataModel);
        String htmlContent = templateEngine.process("report_template", context);

        // 5. OpenHTMLToPDF 轉存 PDF
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        PdfRendererBuilder builder = new PdfRendererBuilder();
        builder.useFastMode();

        File fontFile = new File("C:/Windows/Fonts/msjh.ttc");
        if (fontFile.exists()) {
            builder.useFont(fontFile, "msjh");
        }

        builder.withHtmlContent(htmlContent, "/");
        builder.toStream(os);
        builder.run();

        return os.toByteArray();
    }

    // =========================================================================
    // 1. 動態讀取學生人數報告
    // =========================================================================
    private Map<String, String> readStudentStatsDynamically(Workbook wb, String targetMonth, int actualChurnCount) {
        Map<String, String> stats = new HashMap<>();
        stats.put("month", targetMonth);
        stats.put("totalStudents", "0");
        stats.put("avgStudents", "0");
        stats.put("classCount", "0");
        stats.put("TJ", "0");
        stats.put("MJ", "0");
        stats.put("ERE", "0");
        stats.put("BM", "0");
        stats.put("juniorSeniorHigh", "");
        stats.put("gradeJuniorHigh", "");
        stats.put("gradePrimary", "");
        stats.put("churnCount", String.valueOf(actualChurnCount));
        stats.put("churnRate", "0.00%");

        // A. 讀取月報合計
        Sheet monthSheet = wb.getSheet(targetMonth + "學生流動月報");
        double lastMonthTotal = 0;
        if (monthSheet != null) {
            for (int r = 0; r <= monthSheet.getLastRowNum(); r++) {
                Row row = monthSheet.getRow(r);
                if (row == null) continue;

                String c0 = ExcelHelper.getCellValue(row.getCell(0));
                String c4 = ExcelHelper.getCellValue(row.getCell(4));

                if (c4.contains("合") && c4.contains("計")) {
                    try {
                        lastMonthTotal = Double.parseDouble(ExcelHelper.getCellValue(row.getCell(5)));
                    } catch (Exception ignored) {}
                    stats.put("totalStudents", ExcelHelper.getCellValue(row.getCell(11)));
                }

                if (c0.contains("班級數")) {
                    stats.put("classCount", ExcelHelper.getCellValue(row.getCell(2)));
                } else if (c0.contains("平均人數")) {
                    try {
                        double avg = Double.parseDouble(ExcelHelper.getCellValue(row.getCell(2)));
                        stats.put("avgStudents", new DecimalFormat("#.0").format(avg));
                    } catch (Exception e) {
                        stats.put("avgStudents", ExcelHelper.getCellValue(row.getCell(2)));
                    }
                }
            }
        }

        // B. 計算流失率
        if (lastMonthTotal > 0 && actualChurnCount > 0) {
            double rate = (actualChurnCount / lastMonthTotal) * 100.0;
            stats.put("churnRate", new DecimalFormat("0.00").format(rate) + "%");
        }

        // C. 加總明細班型（動態掃描前 5 列以精準定位 TJ/MJ/菁英/BM 欄位）
        Sheet detailSheet = wb.getSheet(targetMonth + "學生流動明細");
        if (detailSheet != null) {
            int headerRowIdx = -1;
            int tjCol = -1, mjCol = -1, ereCol = -1, bmCol = -1;

            for (int r = 0; r <= Math.min(5, detailSheet.getLastRowNum()); r++) {
                Row row = detailSheet.getRow(r);
                if (row == null) continue;
                for (int c = 0; c < row.getLastCellNum(); c++) {
                    String h = ExcelHelper.getCellValue(row.getCell(c));
                    if ("TJ".equalsIgnoreCase(h)) tjCol = c;
                    else if ("MJ".equalsIgnoreCase(h)) mjCol = c;
                    else if ("菁英".equalsIgnoreCase(h) || h.contains("菁英")) ereCol = c;
                    else if ("BM".equalsIgnoreCase(h)) bmCol = c;
                }
                if (tjCol != -1) {
                    headerRowIdx = r;
                    break;
                }
            }

            if (headerRowIdx != -1) {
                int tjSum = 0, mjSum = 0, ereSum = 0, bmSum = 0;
                for (int r = headerRowIdx + 1; r <= detailSheet.getLastRowNum(); r++) {
                    Row row = detailSheet.getRow(r);
                    if (row == null) continue;

                    String className = ExcelHelper.getCellValue(row.getCell(2));
                    if (className.isEmpty() || className.contains("班別") || className.contains("合")) continue;

                    tjSum += parseCellInt(row, tjCol);
                    mjSum += parseCellInt(row, mjCol);
                    ereSum += parseCellInt(row, ereCol);
                    bmSum += parseCellInt(row, bmCol);
                }

                stats.put("TJ", String.valueOf(tjSum));
                stats.put("MJ", String.valueOf(mjSum));
                stats.put("ERE", String.valueOf(ereSum));
                stats.put("BM", String.valueOf(bmSum));
            }
        }

        return stats;
    }

    private int parseCellInt(Row row, int colIdx) {
        if (colIdx < 0) return 0;
        try {
            String val = ExcelHelper.getCellValue(row.getCell(colIdx));
            if (val.isEmpty()) return 0;
            return (int) Double.parseDouble(val);
        } catch (Exception e) {
            return 0;
        }
    }

    // =========================================================================
    // 2. 智慧讀取每月招生人數（完全動態表頭與欄位，不受未來班級更名影響）
    // =========================================================================
    private Map<String, Object> readMonthlyEnrollmentDynamic(Workbook wb) {
        Map<String, Object> result = new HashMap<>();
        List<String> headers = new ArrayList<>();
        List<List<String>> rows = new ArrayList<>();
        List<String> totalRow = null;

        // 新版工作表名稱為「招生人數」，舊版為「學生人數目標及成果」，都找不到就用第一個工作表
        Sheet sheet = wb.getSheet("招生人數");
        if (sheet == null) sheet = wb.getSheet("學生人數目標及成果");
        if (sheet == null) sheet = wb.getSheetAt(0);

        int startRow = -1;
        int monthCol = -1;

        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                String val = ExcelHelper.getCellValue(row.getCell(c));
                if ("招生人數".equals(val) || val.contains("招生人數")) {
                    startRow = r;
                    monthCol = c;
                    break;
                }
            }
            if (startRow != -1) break;
        }

        if (startRow != -1 && monthCol != -1) {
            Row headerRow = sheet.getRow(startRow);
            int colCount = 0;

            for (int c = monthCol; c < headerRow.getLastCellNum(); c++) {
                String h = ExcelHelper.getCellValue(headerRow.getCell(c));
                if (h.isEmpty()) {
                    if (c > monthCol) break;
                }
                headers.add(h);
                colCount++;
                if (c > monthCol && (h.contains("總") || h.contains("計"))) {
                    break;
                }
            }

            for (int r = startRow + 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String monthVal = ExcelHelper.getCellValue(row.getCell(monthCol));
                if (monthVal.isEmpty()) continue;

                List<String> rowData = new ArrayList<>();
                for (int i = 0; i < colCount; i++) {
                    rowData.add(ExcelHelper.getCellValue(row.getCell(monthCol + i)));
                }

                if (monthVal.contains("總計") || monthVal.contains("合計")) {
                    totalRow = rowData;
                    break;
                } else {
                    rows.add(rowData);
                }
            }
        }

        result.put("headers", headers);
        result.put("rows", rows);
        result.put("totalRow", totalRow);
        return result;
    }

    // =========================================================================
    // 2-1. 讀取「各年級統計」：國高中人數、國小生人數、小六人數
    //      表頭與欄位位置皆動態尋找，月份列以 A 欄等於 targetMonth（例如「10月」）比對
    // =========================================================================
    private Map<String, String> readGradeStats(Workbook wb, String targetMonth) {
        Map<String, String> result = new HashMap<>();
        result.put("juniorSeniorHigh", "");   // 國、高中人數
        result.put("gradePrimary", "");       // 國小/小六生人數，格式：289/61
        result.put("primaryTotal", "");       // 國小生人數
        result.put("gradeSix", "");           // 小六人數

        // A. 找工作表：優先用名稱，找不到就掃描哪個工作表有「國高中人數」表頭
        Sheet sheet = wb.getSheet("各年級統計");
        if (sheet == null) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                if (findHeaderRow(wb.getSheetAt(i)) != -1) {
                    sheet = wb.getSheetAt(i);
                    break;
                }
            }
        }
        if (sheet == null) return result;

        int headerRowIdx = findHeaderRow(sheet);
        if (headerRowIdx == -1) return result;

        // B. 動態定位欄位
        int highCol = -1, primaryCol = -1, sixCol = -1;
        Row headerRow = sheet.getRow(headerRowIdx);
        for (int c = 0; c < headerRow.getLastCellNum(); c++) {
            String h = normalizeHeader(ExcelHelper.getCellValue(headerRow.getCell(c)));
            if (h.equals("國高中人數")) highCol = c;
            else if (h.equals("國小生人數")) primaryCol = c;
            else if (h.equals("小六")) sixCol = c;
        }

        // C. 找到目標月份那一列
        String month = targetMonth == null ? "" : targetMonth.trim();
        for (int r = headerRowIdx + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            if (!month.equals(ExcelHelper.getCellValue(row.getCell(0)).trim())) continue;

            String high = toIntString(row, highCol);
            String primary = toIntString(row, primaryCol);
            String six = toIntString(row, sixCol);

            result.put("juniorSeniorHigh", high);
            result.put("gradePrimary", six);
            break;
        }
        return result;
    }

    // 在前 5 列內找含有「國高中人數」的表頭列，找不到回傳 -1
    private int findHeaderRow(Sheet sheet) {
        if (sheet == null) return -1;
        for (int r = 0; r <= Math.min(5, sheet.getLastRowNum()); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < row.getLastCellNum(); c++) {
                if (normalizeHeader(ExcelHelper.getCellValue(row.getCell(c))).equals("國高中人數")) {
                    return r;
                }
            }
        }
        return -1;
    }

    // 去除空白、換行與頓號，避免「國、高中 人數」這類寫法比對失敗
    private String normalizeHeader(String s) {
        return s == null ? "" : s.replaceAll("[\\s、，,]", "");
    }

    // 讀整數字串；空白或 0 一律回傳空字串（月份尚無資料時不顯示 0）
    private String toIntString(Row row, int colIdx) {
        if (colIdx < 0) return "";
        try {
            String val = ExcelHelper.getCellValue(row.getCell(colIdx));
            if (val == null || val.trim().isEmpty()) return "";
            long n = Math.round(Double.parseDouble(val.trim()));
            return n == 0 ? "" : String.valueOf(n);
        } catch (Exception e) {
            return "";
        }
    }

    // =========================================================================
    // 3. 讀取歷年月底人數（整合 110~115 年共 6 年）
    // =========================================================================
    private Map<String, Object> readHistoricalStatsComplete(Workbook wb) {
        Map<String, Object> result = new HashMap<>();
        List<Map<String, String>> list = new ArrayList<>();
        Map<String, String> avgRow = new HashMap<>();

        Sheet sheet111 = wb.getSheet("111~115年");
        if (sheet111 == null) sheet111 = wb.getSheetAt(0);
        Sheet sheet109 = wb.getSheet("109~113年");

        for (int m = 1; m <= 12; m++) {
            int rowIdx = m + 2;
            Row r111 = sheet111.getRow(rowIdx);
            Row r109 = (sheet109 != null) ? sheet109.getRow(rowIdx) : null;

            Map<String, String> data = new HashMap<>();
            data.put("month", m + "月");
            data.put("y110", r109 != null ? formatHistory(r109, 32, 27, 22) : "");
            data.put("y111", formatHistory(r111, 31, 26, 21));
            data.put("y112", formatHistory(r111, 32, 27, 22));
            data.put("y113", formatHistory(r111, 33, 28, 23));
            data.put("y114", formatHistory(r111, 34, 29, 24));
            data.put("y115", formatHistory(r111, 35, 30, 25));

            list.add(data);
        }

        Row avgRow111 = sheet111.getRow(16);
        Row avgRow109 = (sheet109 != null) ? sheet109.getRow(16) : null;

        avgRow.put("month", "平均");
        avgRow.put("y110", avgRow109 != null ? formatHistory(avgRow109, 32, 27, 22) : "");
        avgRow.put("y111", formatHistory(avgRow111, 31, 26, 21));
        avgRow.put("y112", formatHistory(avgRow111, 32, 27, 22));
        avgRow.put("y113", formatHistory(avgRow111, 33, 28, 23));
        avgRow.put("y114", formatHistory(avgRow111, 34, 29, 24));
        avgRow.put("y115", formatHistory(avgRow111, 35, 30, 25));

        result.put("list", list);
        result.put("avg", avgRow);
        return result;
    }

    // 核心修復：嚴謹處理缺漏欄位，不再產生 352()/ 這種破圖格式
    private String formatHistory(Row row, int totalIdx, int classIdx, int avgIdx) {
        if (row == null) return "";

        String totalStr = ExcelHelper.getCellValue(row.getCell(totalIdx));
        String classStr = ExcelHelper.getCellValue(row.getCell(classIdx));
        String avgStr = ExcelHelper.getCellValue(row.getCell(avgIdx));

        if (totalStr.isEmpty() || totalStr.equals("0")) return "";

        try {
            double tot = Double.parseDouble(totalStr);
            if (tot <= 0) return "";
            long totRound = Math.round(tot);

            // 若班級數有值
            if (!classStr.isEmpty()) {
                double cls = Double.parseDouble(classStr);
                long clsRound = Math.round(cls);

                // 若班平均也有值，直接採用
                if (!avgStr.isEmpty()) {
                    double avg = Double.parseDouble(avgStr);
                    String avgFmt = (avg == Math.floor(avg)) ? String.valueOf((long) avg) : new DecimalFormat("#.0").format(avg);
                    return totRound + "(" + clsRound + ")/" + avgFmt;
                } else if (clsRound > 0) {
                    // 若班平均為空或公式報錯，由程式自動除法補算
                    double autoAvg = tot / (double) clsRound;
                    return totRound + "(" + clsRound + ")/" + new DecimalFormat("#.0").format(autoAvg);
                }
                return totRound + "(" + clsRound + ")";
            }

            // 若完全沒填班級數，僅回傳總人數，不帶括號與斜線
            return String.valueOf(totRound);

        } catch (Exception e) {
            // 例外防呆：只在兩者皆存在時才拼裝，絕不輸出空括號
            if (!classStr.isEmpty() && !avgStr.isEmpty()) {
                return totalStr + "(" + classStr + ")/" + avgStr;
            } else if (!classStr.isEmpty()) {
                return totalStr + "(" + classStr + ")";
            }
            return totalStr;
        }
    }

    // =========================================================================
    // 4. 讀取流失名冊（過濾標題列與空白列，依真實名冊動態增減列數）
    // =========================================================================
    private List<Map<String, String>> readChurnList(Sheet sheet) {
        List<Map<String, String>> list = new ArrayList<>();
        if (sheet == null) return list;

        int serialNo = 1;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String studentName = ExcelHelper.getCellValue(row.getCell(4));
            if (studentName.isEmpty() || "學生姓名".equals(studentName) || studentName.contains("姓名")) {
                continue;
            }

            Map<String, String> data = new HashMap<>();
            data.put("no", String.valueOf(serialNo++));
            data.put("date", ExcelHelper.getCellValue(row.getCell(1)));
            data.put("teacher", ExcelHelper.getCellValue(row.getCell(2)));
            data.put("classType", ExcelHelper.getCellValue(row.getCell(3)));
            data.put("studentName", studentName);
            data.put("reason", ExcelHelper.getCellValue(row.getCell(7)));
            list.add(data);
        }
        return list;
    }
}