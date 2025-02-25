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
            int lastCol1 = sheet1.getRow(0).getLastCellNum();
            int lastRow2 = sheet2.getLastRowNum();
            int lastCol2 = sheet2.getRow(0).getLastCellNum();

            // Column Mapping
            int[] colMap = new int[lastCol1];
            for (int j = 0; j < lastCol1; j++) {
                String header1 = "";
                if (sheet1.getRow(0).getCell(j) != null) {
                    header1 = sheet1.getRow(0).getCell(j).getStringCellValue().replaceAll(" ", "");
                }
                colMap[j] = -1;
                for (int k = 0; k < lastCol2; k++) {
                    String header2 = "";
                    if (sheet2.getRow(0).getCell(k) != null) {
                        header2 = sheet2.getRow(0).getCell(k).getStringCellValue().replaceAll(" ", "");
                    }
                    if (header1.equals(header2)) {
                        colMap[j] = k;
                        break;
                    }
                }
            }

            // Row Mapping
            Map<String, Integer> sheet1IdMap = new HashMap<>();
            Map<String, Integer> sheet2IdMap = new HashMap<>();

            for (int i = 1; i <= lastRow1; i++) {
                String id1 = getCellValueAsString(sheet1.getRow(i).getCell(0)).replaceAll(" ", "");
                sheet1IdMap.put(id1, i);
            }

            for (int i = 1; i <= lastRow2; i++) {
                String id2 = getCellValueAsString(sheet2.getRow(i).getCell(0)).replaceAll(" ", "");
                sheet2IdMap.put(id2, i);
            }

            // Compare Rows and Columns
            for (Map.Entry<String, Integer> entry : sheet1IdMap.entrySet()) {
                String id1 = entry.getKey();
                int row1 = entry.getValue();

                if (sheet2IdMap.containsKey(id1)) {
                    int row2 = sheet2IdMap.get(id1);

                    for (int j = 0; j < lastCol1; j++) {
                        if (colMap[j] != -1) {
                            int k = colMap[j];
                            String val1 = getCellValueAsString(sheet1.getRow(row1).getCell(j)).replaceAll(" ", "");
                            String val2 = getCellValueAsString(sheet2.getRow(row2).getCell(k)).replaceAll(" ", "");

                            CellStyle greenStyle = workbook.createCellStyle();
                            greenStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                            greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                            CellStyle redStyle = workbook.createCellStyle();
                            redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                            redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                            if ((val1.equals("-") && val2.isEmpty()) || (val1.isEmpty() && val2.equals("-")) || val1.equals(val2)) {
                                if (sheet1.getRow(row1).getCell(j) == null) {
                                    sheet1.getRow(row1).createCell(j).setCellStyle(greenStyle);
                                } else {
                                    sheet1.getRow(row1).getCell(j).setCellStyle(greenStyle);
                                }

                                if (sheet2.getRow(row2).getCell(k) == null) {
                                    sheet2.getRow(row2).createCell(k).setCellStyle(greenStyle);
                                } else {
                                    sheet2.getRow(row2).getCell(k).setCellStyle(greenStyle);
                                }

                            } else {
                                if (sheet1.getRow(row1).getCell(j) == null) {
                                    sheet1.getRow(row1).createCell(j).setCellStyle(redStyle);
                                } else {
                                    sheet1.getRow(row1).getCell(j).setCellStyle(redStyle);
                                }

                                if (sheet2.getRow(row2).getCell(k) == null) {
                                    sheet2.getRow(row2).createCell(k).setCellStyle(redStyle);
                                } else {
                                    sheet2.getRow(row2).getCell(k).setCellStyle(redStyle);
                                }
                            }
                        }
                    }
                } else {
                    System.out.println("ID: " + id1 + " not found in Sheet2");
                }
            }

            try (FileOutputStream fos = new FileOutputStream(filepath)) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
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
            e.printStackTrace();
            return "";
        }
    }
}