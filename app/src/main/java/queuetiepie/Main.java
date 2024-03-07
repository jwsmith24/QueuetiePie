package queuetiepie;

public class Main {

    public static void main(String[] args) {
        System.out.println("Get ready to queue!");


        Queue queueOutput = new Queue();

        try {
            // queueOutput.readInFile();

            queueOutput.writeToFile();


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
}
