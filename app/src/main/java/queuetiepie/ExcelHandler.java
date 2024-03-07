package queuetiepie;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelHandler {

private static final int TARGET_COLUMN_INDEX = 1;


    public void processExcel(String filePath) {

        try {

            Workbook workbook = readInWorkbook(filePath);
            addColumn(workbook, TARGET_COLUMN_INDEX);







        } catch (IOException e) {
            e.printStackTrace();
        }
        


    }


private static Workbook readInWorkbook(String filePath) throws IOException {


    try (FileInputStream inputStream = new FileInputStream(filePath)) {

        return WorkbookFactory.create(inputStream);
    }
        
}


private static void calculateBreaks(Workbook workbook, Employee employee) {



}


private static void saveWorkbook(Workbook workbook, String filepath) throws IOException{ 
    try (FileOutputStream outputStream = new FileOutputStream(filepath)) {
        workbook.write(outputStream);
    }
}


private static void addColumn(Workbook workbook, int columnIndex) {
    
    Sheet sheet = workbook.getSheetAt(0); // Assuming first sheet
        
    int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue; // Skip if row is null
            }
            // Shift existing cells to the right to make space for the new column
            for (int j = row.getLastCellNum(); j > columnIndex; j--) {
                Cell cell = row.getCell(j - 1);
                if (cell != null) {
                    row.createCell(j);
                    row.getCell(j).setCellValue(cell.getStringCellValue());
                }
            }

            // Add the new cell at the desired column index
            row.createCell(columnIndex);
            row.getCell(columnIndex).setCellValue(""); 
        }
}

    


    
    
}
