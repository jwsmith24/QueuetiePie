package queuetiepie;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.time.LocalDateTime;
import java.time.Duration;


public class ExcelHandler {

private static final int TARGET_COLUMN_INDEX = 1;


    public void processExcel(String filePath) {

        try {

            Workbook workbook = readInWorkbook(filePath);
            addColumn(workbook, TARGET_COLUMN_INDEX);
            calculateBreaks(workbook);
            saveWorkbook(workbook, filePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
        


    }


private static Workbook readInWorkbook(String filePath) throws IOException {


    try (FileInputStream inputStream = new FileInputStream(filePath)) {

        return WorkbookFactory.create(inputStream);
    }
        
}


private static void calculateBreaks(Workbook workbook) {

    Sheet sheet = workbook.getSheetAt(0);

    // format the date/time parser
    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/dd/yyyy HH:mm");



    for (Row row : sheet) {
        Cell cell1 = row.getCell(0); // Assuming datetime values are in the first column
        
        if (cell1 != null && cell1.getCellType() == CellType.STRING) {
            int rowIndex = row.getRowNum();
            String datetimeStr1 = cell1.getStringCellValue();
            LocalDateTime dateTime1 = LocalDateTime.parse(datetimeStr1, formatter);

            if (rowIndex > 0) {
                Row prevRow = sheet.getRow(rowIndex - 1);
                Cell prevCell = prevRow.getCell(0);
               
                if (prevCell != null && prevCell.getCellType() == CellType.STRING) {
                    String datetimeStr2 = prevCell.getStringCellValue();
                    LocalDateTime dateTime2 = LocalDateTime.parse(datetimeStr2, formatter);

                    // Calculate the difference in minutes
                    long minutesDifference = Duration.between(dateTime2, dateTime1).toMinutes();

                    // Output the difference in the next column
                    Cell diffCell = row.createCell(1);
                    diffCell.setCellValue(minutesDifference);
                }
            }
        }
    }

    

}


private static void saveWorkbook(Workbook workbook, String filepath) throws IOException{ 
    try (FileOutputStream outputStream = new FileOutputStream(filepath)) {
        workbook.write(outputStream);
        workbook.close();
        
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
