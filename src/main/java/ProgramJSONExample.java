import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class ProgramJSONExample {

    public static void main(String[] args) throws IOException {
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        String fileData = new String(Files.readAllBytes(Paths
                .get("C:\\Users\\HP\\IdeaProjects\\JSONProject\\JSONInfo.json")));
        JSONInfo jsonInfo = gson.fromJson(fileData, JSONInfo.class);
        if (jsonInfo != null) {
            writeExcel(jsonInfo); // Method to write data in excel
        } else {
            System.out.println("No data to write in excel, json is null or empty.");
        }
    }



    private static void writeExcel(JSONInfo jsonInfo) {

        HSSFWorkbook hssfWorkbook = null;
        HSSFRow row = null;
        HSSFSheet hssfSheet = null;
        FileOutputStream fileOutputStream = null;
        Properties properties = null;
        try {
            String filename = "C:\\Users\\HP\\IdeaProjects\\JSONProject\\JSONInfo.xls";
            hssfWorkbook = new HSSFWorkbook();
            hssfSheet = hssfWorkbook.createSheet("new sheet");

            HSSFRow rowhead = hssfSheet.createRow((short) 0); // Header
            rowhead.createCell((short) 0).setCellValue("SNo");
            rowhead.createCell((short) 1).setCellValue("name");
            rowhead.createCell((short) 2).setCellValue("age");
            rowhead.createCell((short) 3).setCellValue("marks");


            int counter = 1;
            for (ProgramInfo programInfo : jsonInfo.getRecords()) {
                properties = programInfo.getProperties();
                row = hssfSheet.createRow((short) counter);
                row.createCell((short) 0).setCellValue(counter);
                row.createCell((short) 1).setCellValue(
                        properties.getName());
                row.createCell((short) 2).setCellValue(
                        properties.getAge());
                row.createCell((short) 3).setCellValue(
                        properties.getMarks());

                counter++;
            }


            fileOutputStream = new FileOutputStream(filename);
            hssfWorkbook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("JSON data successfully exported to excel!");
        } catch (Throwable throwable) {
            System.out.println("Exception in writting data to excel : "
                    + throwable);
        }
    }
}