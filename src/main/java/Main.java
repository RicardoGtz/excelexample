import java.io.*;

public class Main{
    public static void main(String args[])throws IOException {
        try {
            //New object ExcelReader
            ExcelReader er = new ExcelReader(".\\sample.xlsx");
            //New fos to save the formatted Excel file
            FileOutputStream fos = new FileOutputStream(new File("salida.xlsx"));
            er.clearExcel(fos, 0);
        }catch (ExcelReaderException e){
            System.out.println(e.getMessage());
        }

    }
}