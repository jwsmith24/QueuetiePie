package queuetiepie;

import java.io.File; 
import java.io.FileInputStream; 
import java.io.FileOutputStream;
import java.io.IOException; 
import java.util.TreeMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell; 
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.io.FileBackedOutputStream;




public class Queue {

    

    public void writeToWorkbook() {
        
        
    // Make a blank workbook
    XSSFWorkbook workbook = new XSSFWorkbook();
        // Add a new blank sheet
         XSSFSheet sheet = workbook.createSheet("test");

        // Empty tree map to store data to input
            Map<String, Object[]> data = new TreeMap<String, Object[]>();

        // write data into the map

            data.put("1", new Object[] {"ID", "FIRST NAME", "LAST NAME"});
            data.put("2", new Object[] { 1, "Pankaj", "Kumar" }); 
            data.put("3", new Object[] { 2, "Prakashni", "Yadav" }); 
            data.put("4", new Object[] { 3, "Ayan", "Mondal" }); 
            data.put("5", new Object[] { 4, "Virat", "kohli" }); 

            Set<String> keyset = data.keySet();

            int rowNum = 0;

            for (String key: keyset) {

                Row row = sheet.createRow(rowNum++);

                Object[] objArray = data.get(key);

                int cellNum = 0;

                for (Object obj : objArray) {
                    Cell cell = row.createCell(cellNum++);


                    if (obj instanceof String) {
                        cell.setCellValue((String)obj);
                    } else if (obj instanceof Integer) {
                        cell.setCellValue((Integer)obj);
                    }
                }
            }

            // writing to the workbook
            try (FileOutputStream output = new FileOutputStream(new File("test.xlsx"))) {

                workbook.write(output);

                System.out.println("test.xlsx written succesfully");

    } catch (IOException e) {

        e.printStackTrace();
    }

    }

}


