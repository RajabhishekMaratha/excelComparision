package excel;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class sectionComparision {

    public static void main(String[] args) {
        try {
            // Adjust the minimum inflate ratio
            ZipSecureFile.setMinInflateRatio(0.001);

            FileInputStream file = new FileInputStream("C:\\Users\\2320611\\eclipse-workspace\\Selenium\\excel\\src\\main\\resources\\comparision.xlsx");

            Workbook workbook = new XSSFWorkbook(file);

            Sheet sheet1 = workbook.getSheet("HFM2");
            Sheet sheet2 = workbook.getSheet("FCCS2");

            int lastRow1 = sheet1.getLastRowNum();
            int lastRow2 = sheet2.getLastRowNum();
            int lastCol1 = sheet1.getRow(0).getLastCellNum();
            int lastCol2 = sheet2.getRow(0).getLastCellNum();

            // Initialize the column mapping array
            Map<Integer, Integer> colMap = new HashMap<>();
            for (int j = 0; j < lastCol1; j++) {
                String header1 = getCellValueAsString(sheet1.getRow(0).getCell(j)).replace(" ", "");
                for (int k = 0; k < lastCol2; k++) {
                    String header2 = getCellValueAsString(sheet2.getRow(0).getCell(k)).replace(" ", "");
                    if (header1.equals(header2)) {
                        colMap.put(j, k);
                        break;
                    }
                }
            }

            int startRow1 = 1; // Assuming headers are in the first row
            int startRow2 = 1; // Assuming headers are in the first row

            while (startRow1 <= lastRow1 && startRow2 <= lastRow2) {
                // Find the end of the current section
                int endRow1 = startRow1;
                while (endRow1 <= lastRow1 && !getCellValueAsString(sheet1.getRow(endRow1).getCell(0)).equals("0")) {
                    endRow1++;
                }
                endRow1--;

                int endRow2 = startRow2;
                while (endRow2 <= lastRow2 && !getCellValueAsString(sheet2.getRow(endRow2).getCell(0)).equals("0")) {
                    endRow2++;
                }
                endRow2--;

                // Compare rows within the current section
                for (int i = startRow1; i <= endRow1; i++) {
                    String id1 = getCellValueAsString(sheet1.getRow(i).getCell(0)).replace(" ", "");

                    int row2 = -1;
                    for (int j = startRow2; j <= endRow2; j++) {
                        String id2 = getCellValueAsString(sheet2.getRow(j).getCell(0)).replace(" ", "");
                        if (id1.equals(id2)) {
                            row2 = j;
                            break;
                        }
                    }

                    // If a matching row is found, compare the rows
                    if (row2 != -1) {
                        for (int j = 0; j < lastCol1; j++) {
                            if (colMap.containsKey(j)) {
                                int k = colMap.get(j);
                                Cell cellVal1 = sheet1.getRow(i).getCell(j);
                                if (cellVal1 == null) {
                                    cellVal1 = sheet1.getRow(i).createCell(j);
                                }
                                String val1 = getCellValueAsString(cellVal1).replace(" ", "");

                                Cell cellVal2 = sheet2.getRow(row2).getCell(k);
                                if (cellVal2 == null) {
                                    cellVal2 = sheet2.getRow(row2).createCell(k);
                                }
                                String val2 = getCellValueAsString(cellVal2).replace(" ", "");

                                CellStyle greenStyle = workbook.createCellStyle();
                                greenStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
                                greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                                CellStyle redStyle = workbook.createCellStyle();
                                redStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                                redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                                if ((val1.equals("-") && val2.equals("")) || (val1.equals("") && val2.equals("-")) || val1.equals(val2)) {
                                    cellVal1.setCellStyle(greenStyle);
                                    cellVal2.setCellStyle(greenStyle);
                                } else {
                                    cellVal1.setCellStyle(redStyle);
                                    cellVal2.setCellStyle(redStyle);
                                }
                            }
                        }
                    }
                }

                // Move to the next section
                startRow1 = endRow1 + 2;
                startRow2 = endRow2 + 2;
            }

            file.close();

            FileOutputStream outFile = new FileOutputStream("C:\\Users\\2320611\\eclipse-workspace\\Selenium\\excel\\src\\main\\resources\\comparision.xlsx");
            workbook.write(outFile);
            outFile.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
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
                    return String.valueOf(cell.getNumericCellValue());
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