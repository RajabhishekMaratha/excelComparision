package excel;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelSheetComparator {

    public static void main(String[] args) {
        ZipSecureFile.setMinInflateRatio(0.001);

        String filepath = "C:\\Users\\2320611\\eclipse-workspace\\Selenium\\excel\\src\\main\\resources\\comparision.xlsx";
        String sheet1Name = "HFM3";
        String sheet2Name = "FCCS2";

        try (FileInputStream fis = new FileInputStream(filepath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet1 = workbook.getSheet(sheet1Name);
            Sheet sheet2 = workbook.getSheet(sheet2Name);

            int lastRow1 = sheet1.getLastRowNum();
            int lastRow2 = sheet2.getLastRowNum();

            // *** COLUMN MAPPING SECTION ***
            // Create header maps for both sheets (header name -> column index)
            Map<String, Integer> sheet1HeaderMap = getHeaderMap(sheet1);
            Map<String, Integer> sheet2HeaderMap = getHeaderMap(sheet2);
            // *** END OF COLUMN MAPPING SECTION ***

            // *** ROW MAPPING SECTION ***
            // Create ID maps for both sheets (ID -> row index)
            Map<String, Integer> sheet1IdMap = getIdMap(sheet1);
            Map<String, Integer> sheet2IdMap = getIdMap(sheet2);
            // *** END OF ROW MAPPING SECTION ***

            for (Map.Entry<String, Integer> entry : sheet1IdMap.entrySet()) {
                String id1 = entry.getKey();
                int row1 = entry.getValue();

                if (sheet2IdMap.containsKey(id1)) {
                    int row2 = sheet2IdMap.get(id1);
                    compareRow(sheet1, sheet2, row1, row2, sheet1HeaderMap, sheet2HeaderMap, workbook);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(filepath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // Method to create a header map (header name -> column index) for a sheet
    private static Map<String, Integer> getHeaderMap(Sheet sheet) {
        Map<String, Integer> headerMap = new HashMap<>();
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    headerMap.put(cell.getStringCellValue().replaceAll(" ", ""), i);
                }
            }
        }
        return headerMap;
    }

    // Method to create an ID map (ID -> row index) for a sheet
    private static Map<String, Integer> getIdMap(Sheet sheet) {
        Map<String, Integer> idMap = new HashMap<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Cell idCell = sheet.getRow(i).getCell(0);
            String id = getCellValueAsString(idCell).replaceAll(" ", "");
            if (!id.isEmpty()) {
                idMap.put(id, i);
            }
        }
        return idMap;
    }

    // Method to compare two rows based on header mappings and apply cell styles
    private static void compareRow(Sheet sheet1, Sheet sheet2, int row1, int row2, Map<String, Integer> sheet1HeaderMap, Map<String, Integer> sheet2HeaderMap, Workbook workbook) {
        CellStyle greenStyle = workbook.createCellStyle();
        greenStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle redStyle = workbook.createCellStyle();
        redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (Map.Entry<String, Integer> headerEntry : sheet1HeaderMap.entrySet()) {
            String header = headerEntry.getKey();
            int col1 = headerEntry.getValue();

            if (sheet2HeaderMap.containsKey(header)) {
                int col2 = sheet2HeaderMap.get(header);

                String val1 = getCellValueAsString(sheet1.getRow(row1).getCell(col1)).replaceAll(" ", "");
                String val2 = getCellValueAsString(sheet2.getRow(row2).getCell(col2)).replaceAll(" ", "");

                if ((val1.equals("-") && val2.isEmpty()) || (val1.isEmpty() && val2.equals("-")) || val1.equals(val2)) {
                    setCellStyle(sheet1, row1, col1, greenStyle);
                    setCellStyle(sheet2, row2, col2, greenStyle);
                } else {
                    setCellStyle(sheet1, row1, col1, redStyle);
                    setCellStyle(sheet2, row2, col2, redStyle);
                }
            }
        }
    }

    private static void setCellStyle(Sheet sheet, int row, int col, CellStyle style) {
        Row r = sheet.getRow(row);
        if(r == null){
            r = sheet.createRow(row);
        }
        Cell c = r.getCell(col);
        if(c == null){
            c = r.createCell(col);
        }
        c.setCellStyle(style);
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    return String.valueOf((int) cell.getNumericCellValue());
                case BOOLEAN:  
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return cell.getCellFormula();
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }
}