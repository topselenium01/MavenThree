import org.openqa.selenium.WebDriver;

import java.io.IOException;

public class MavenThree {

    public static void main(String[] args) {

        String path = "src/main/resources/data/TestData.xlsx";
        try {
            ExcelReader.setExcelFile(path,"Sites");
            String strTestData = ExcelReader.fetchData("testFlow1", "Product Name");
            ExcelReader.writeExcel(path,"Sites" ,"testFlow1", "Test Result", "Passed");
            //TODO: Update writeExcel method to handle creation of multiple columns in same method call
            ExcelReader.writeExcel(path,"Sites" ,"testFlow2", "Test Result-##-Screenshot Required", "Not Run-##-Not Necessary");
            System.out.println("Test Data is: " + strTestData);

        } catch (IOException e) {
            System.out.println("File not found");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
