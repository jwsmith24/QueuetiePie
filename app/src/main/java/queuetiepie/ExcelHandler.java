package queuetiepie;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;


public class ExcelHandler {

    private static final int TARGET_COLUMN_INDEX = 1;

    /**
     * Reads in target xls/xlsx file in to be processed.
     *
     * @param filePath target filepath
     * @return workbook object
     * @throws IOException if filepath is unreachable
     */
    private static Workbook readInWorkbook(String filePath) throws IOException {


        try (FileInputStream inputStream = new FileInputStream(filePath)) {

            return WorkbookFactory.create(inputStream);
        }

    }

    /**
     * Format cell to ensure it's always a string.
     *
     * @param cell current cell ref to format
     * @return formatted cell
     */
    private static Cell getFormattedDateCell(Cell cell) {
        DataFormatter dataFormatter = new DataFormatter();
        String stringValue = dataFormatter.formatCellValue(cell);

        cell.setCellValue(stringValue);

        return cell;
    }

    /**
     * Adds conditional formatting rule to the target column.
     *
     * @param sheet ref to active sheet
     */
    private static void applyConditionalFormatting(Sheet sheet) {
        // Create conditional formatting rule
        ConditionalFormattingRule conditionalFormattingRule = sheet.getSheetConditionalFormatting()
                .createConditionalFormattingRule(
                        ComparisonOperator.GT,
                        "4.9",
                        null
                );

        // Describe the style to apply if condition is met
        PatternFormatting patternFormatting = conditionalFormattingRule.createPatternFormatting();
        patternFormatting.setFillBackgroundColor(IndexedColors.RED.getIndex());

        // Define the region the conditional formatting will apply to
        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("B1:B1000")
        };

        // Add the rule to the cell range
        sheet.getSheetConditionalFormatting()
                .addConditionalFormatting(regions, conditionalFormattingRule);

    }


    /**
     * Executes adding the calculated processing times to the target cell.
     *
     * @param workbook ref to current workbook
     */
    private static void calculateProcessingTime(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        Row prevRow = null;

        for (Row currentRow : sheet) {
            if (currentRow.getRowNum() <= 1) {
                continue; // Skip the header and first row (employee clocking in)
            }

            // Format cell to STRING
            Cell currentCell = getFormattedDateCell(currentRow.getCell(0));

            // Ensure cell was properly formatted to STRING
            if (currentCell.getCellType() == CellType.STRING) {
                // Grab the time stamp to work with
                String currentTimestampStr = currentCell.getStringCellValue();
                // Perform calculation and insert result to target cell
                calculateTimeDifference(currentTimestampStr, prevRow, currentRow);
                // Update prevRow for next iteration
                prevRow = currentRow;
            }

        }
    }

    /**
     * Parse timestamps for current and previous rows, calculate the difference, and add to the target column.
     */
    private static void calculateTimeDifference(String timestamp, Row prevRow, Row currentRow) {
        // Set date format
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy H:mm");

        try {
            LocalDateTime currentDateTime = LocalDateTime.parse(timestamp, formatter);

            if (prevRow != null) {
                Cell prevCell = prevRow.getCell(0);

                if (prevCell != null) {
                    String prevTimestampStr = prevCell.getStringCellValue();
                    LocalDateTime prevDateTime = LocalDateTime.parse(prevTimestampStr, formatter);

                    // Calculate the difference in minutes
                    long minutesDifference = Duration.between(prevDateTime, currentDateTime).toMinutes();

                    // Output the difference in the target column
                    Cell diffCell = currentRow.createCell(TARGET_COLUMN_INDEX);
                    diffCell.setCellValue(minutesDifference);
                }
            }

        } catch (DateTimeParseException e) {
            // Handle invalid timestamp format
            System.err.println("Invalid timestamp format at row " + (currentRow.getRowNum() + 1)
                    + ": " + timestamp);
        }
    }

    /**
     * Writes workbook back to .xlsx in the target directory.
     *
     * @param workbook workbook ref
     * @param filepath target filepath
     * @throws IOException handled by Main
     */
    private static void saveWorkbook(Workbook workbook, String filepath) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(filepath)) {
            workbook.write(outputStream);
        }
    }

    /**
     * Helper method to add column to the right of the timestamp column. Uses constant TARGET_COLUMN_INDEX to
     * easily change location if necessary.
     *
     * @param workbook workbook ref
     */
    private static void addColumn(Workbook workbook) {

        Sheet sheet = workbook.getSheetAt(0); // Assuming first sheet

        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue; // Skip if row is null
            }

            // Shift existing cells to the right to make space for the new column
            for (int j = row.getLastCellNum(); j > ExcelHandler.TARGET_COLUMN_INDEX; j--) {
                Cell cell = row.getCell(j - 1);
                if (cell != null) {
                    row.createCell(j);
                    row.getCell(j).setCellValue(cell.getStringCellValue());
                }
            }

            // Add the new cell at the desired column index
            row.createCell(ExcelHandler.TARGET_COLUMN_INDEX);
            row.getCell(ExcelHandler.TARGET_COLUMN_INDEX).setCellValue("");
        }
    }

    /**
     * Process the given spreadsheet to add a column with the calculated times between completed tasks.
     *
     * @param filePath filepath of the target spreadsheet
     */
    public void processExcel(String filePath) {

        try {
            Workbook workbook = readInWorkbook(filePath);
            addColumn(workbook);
            calculateProcessingTime(workbook);
            applyConditionalFormatting(workbook.getSheetAt(0));
            saveWorkbook(workbook, filePath);

        } catch (IOException e) {
            System.out.println("Workbook could not be modified");

        }
    }

}
