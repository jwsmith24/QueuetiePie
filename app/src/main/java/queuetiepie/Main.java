package queuetiepie;

public class Main {

    public static void main(String[] args) {
        System.out.println("Get ready to queue!");

        ExcelHandler excelHandler = new ExcelHandler();

        excelHandler.processExcel("/home/jake/repos/QueuetiePie/app/workQueue.xlsx");


    }

}
