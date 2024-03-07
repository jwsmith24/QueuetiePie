package queuetiepie;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
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
    DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy H:mm");
    Row prevRow = null;

    for (Row currentRow : sheet) {
        if (currentRow.getRowNum() <= 1) {
            continue; // Skip the header and first row (employee clocking in)
        }

        Cell currentCell = currentRow.getCell(0); // Time stamp values are in the first column

        if (currentCell != null && currentCell.getCellType() == CellType.STRING) {
            String currentTimestampStr = currentCell.getStringCellValue();

            try {
                LocalDateTime currentDateTime = LocalDateTime.parse(currentTimestampStr, formatter);

                if (prevRow != null) {
                    Cell prevCell = prevRow.getCell(0);

                    if (prevCell != null && prevCell.getCellType() == CellType.STRING) {
                        String prevTimestampStr = prevCell.getStringCellValue();
                        LocalDateTime prevDateTime = LocalDateTime.parse(prevTimestampStr, formatter);

                        // Calculate the difference in minutes
                        long minutesDifference = Duration.between(prevDateTime, currentDateTime).toMinutes();

                        // Output the difference in the next column
                        Cell diffCell = currentRow.createCell(1);
                        diffCell.setCellValue(minutesDifference);
                    }
                }

                prevRow = currentRow; // Update prevRow for next iteration
            } catch (DateTimeParseException e) {
                // Handle invalid timestamp format
                System.err.println("Invalid timestamp format at row " + (currentRow.getRowNum() + 1) + ": " + currentTimestampStr);
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
