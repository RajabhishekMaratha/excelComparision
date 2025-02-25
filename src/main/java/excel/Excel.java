package excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Excel {

    public static void main(String[] args) throws IOException {
    	
    	
        FileInputStream file = new FileInputStream("C:\\Users\\2320611\\eclipse-workspace\\Selenium\\excel\\src\\main\\resources\\Book3.xlsx");
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheet("Sheet1");

        int startRow = 5; // C6 is the 6th row (0-indexed)
        int endRow = 20; // H21 is the 21st row (0-indexed)
        int startCol1 = 2; // C is the 3rd column (0-indexed)
        int endCol1 = 7; // H is the 8th column (0-indexed)
        int startCol2 = 11; // L is the 12th column (0-indexed)
        int endCol2 = 16; // Q is the 17th column (0-indexed)

        for (int i = startRow; i <= endRow; i++) {
            boolean isDifferent = false;

            for (int j = startCol1, k = startCol2; j <= endCol1 && k <= endCol2; j++, k++) {
                Cell cell1 = sheet.getRow(i).getCell(j);
                Cell cell2 = sheet.getRow(i).getCell(k);

                if (compareCells(cell1, cell2)) {
                    isDifferent = true;
                    break;
                }
            }

            if (isDifferent) {
                highlightRow(sheet, i, startCol1, endCol1, IndexedColors.RED);
                highlightRow(sheet, i, startCol2, endCol2, IndexedColors.RED);
            } else {
                highlightRow(sheet, i, startCol1, endCol1, IndexedColors.GREEN);
                highlightRow(sheet, i, startCol2, endCol2, IndexedColors.GREEN);
            }
        }

        file.close();
        FileOutputStream outFile = new FileOutputStream("C:\\Users\\2320611\\eclipse-workspace\\Selenium\\excel\\src\\main\\resources\\Book3.xlsx");
        workbook.write(outFile);
        outFile.close();
        workbook.close();
    }

    private static boolean compareCells(Cell cell1, Cell cell2) {
        if (cell1 == null && cell2 == null) {
            return false; // Both cells are null, so they are considered equal
        }
        if (cell1 == null || cell2 == null) {
            return true; // One cell is null, the other is not
        }
        if (cell1.getCellType() != cell2.getCellType()) {
            return true; // Different types
        }
        switch (cell1.getCellType()) {
            case NUMERIC:
                return cell1.getNumericCellValue() != cell2.getNumericCellValue();
            case STRING:
                return !cell1.getStringCellValue().equals(cell2.getStringCellValue());
            case BOOLEAN:
                return cell1.getBooleanCellValue() != cell2.getBooleanCellValue();
            case BLANK:
                return false; // Both are blank
            default:
                return false; // returns true if the values are different
        }
    }

    private static void highlightRow(Sheet sheet, int rowIndex, int startCol, int endCol, IndexedColors color) {
        Row row = sheet.getRow(rowIndex);
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setFillForegroundColor(color.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (int j = startCol; j <= endCol; j++) {
            Cell cell = row.getCell(j);
            if (cell == null) {
                cell = row.createCell(j);
            }
            cell.setCellStyle(style);
        }
    }
}