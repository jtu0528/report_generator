package com.joyenglish.reportgenerator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import java.text.SimpleDateFormat;

public class ExcelHelper {
    public static String getCellValue(Cell cell) {
        if (cell == null) return "";
        
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return new SimpleDateFormat("M/d").format(cell.getDateCellValue());
            }
            double val = cell.getNumericCellValue();
            if (val == Math.floor(val)) return String.valueOf((long) val);
            return String.valueOf(val);
        } else if (cell.getCellType() == CellType.FORMULA) {
            try {
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val)) return String.valueOf((long) val);
                return String.valueOf(val);
            } catch (Exception e) {
                try {
                    return cell.getStringCellValue().trim();
                } catch (Exception ex) {
                    return "";
                }
            }
        }
        return "";
    }
}