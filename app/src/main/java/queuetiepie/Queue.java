package queuetiepie;

import java.io.File; 
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException; 

import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;







public class Queue {

    String excelFilePath = "output.xlsx";

    public void readInFile() throws IOException {
        
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firsSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firsSheet.iterator();

        while(iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellType()) {
                    
                    case STRING:
                    System.out.print(cell.getStringCellValue());
                    break;
                    case BOOLEAN:
                    System.out.print(cell.getBooleanCellValue());
                    break;
                    case NUMERIC :
                    System.out.print(cell.getNumericCellValue());
                    break;
                    default:
                    System.out.print("Unsupported Cell Type");
                }
                System.out.print(" - ");
            }
            System.out.println();
        }

        workbook.close();
        inputStream.close();


    }




    public void writeToFile() throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet();

        Object[][] cellData = {
            {"Employee", "Total Break Time (Min)"},
            {"Employee 1", 5},
            {"Employee 2", 12},
            {"Employee 3", 2}
        };

        int rowCount = 0;

        for (Object[] rowData: cellData) {
            
            Row row = sheet.createRow(++rowCount);

            int columnCount = 0;

            for (Object cellValue : rowData) {
                Cell cell = row.createCell(++columnCount);

                if (cellValue instanceof String) {

                    cell.setCellValue((String) cellValue);

                } else if (cellValue instanceof Integer) {
                    cell.setCellValue((Integer) cellValue);
                }

            }
        }

        // write content to file
        try (FileOutputStream outputStream = new FileOutputStream("output.xlsx")) {

            workbook.write(outputStream);
    

        }

        workbook.close();

        }
    }



