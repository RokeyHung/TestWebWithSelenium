package StartTest;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static StartTest.ExcelReader.*;

public class RemoveFilm {
    private WebDriver driver;
    private final String filePath = "DataTest.xlsx";
    private final String SheetName = "Test Case";
    int startRow = 21;
    int totalRowData = 1;
    int startWriteRowExcel = startRow;
    int rowNum = 0;

    public RemoveFilm() throws IOException {
    }

    @BeforeTest
    public void setUp() throws InterruptedException {
        String pathDriver = "chromedriver.exe";
        System.setProperty("webdriver.chrome.driver", pathDriver);
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver(options);
        driver.get("https://vuighe.net/");

        WebElement avt = driver.findElement(By.className("navbar-avatar"));
        avt.click();

        WebElement usernameField = driver.findElement(By.name("username"));
        WebElement passwordField = driver.findElement(By.name("password"));
        WebElement loginButton = driver.findElement(By.id("login"));

        usernameField.sendKeys("huyhy03");
        passwordField.sendKeys("123456789");
        loginButton.click();
        Thread.sleep(1500);
    }

    @DataProvider(name = "removeFilm")
    public Object[][] data() throws Exception {
        return getDataCellInColl(filePath, SheetName, startRow, 3, totalRowData);
    }

    @Test(dataProvider = "removeFilm")
    public boolean TestRemoveFilm(String idFilm) throws InterruptedException {
        WebElement avt1 = driver.findElement(By.className("navbar-avatar"));
        Thread.sleep(1500);
        avt1.click();
        WebElement element = driver.findElement(By.xpath("//div[@class='user-item']//a[span[text()='Phim đã xem']]"));
        Thread.sleep(1500);
        element.click();

        try {
            WebElement findElementFilm = driver.findElement(By.cssSelector("div.film-delete[data-id='" + idFilm + "']"));
            Thread.sleep(1500);
            findElementFilm.click();
        } catch (Exception e) {
            System.out.println("Không tìm được film từ findElement");
            return false;
        }

        WebElement okBtn = driver.findElement(By.className("ok"));
        Thread.sleep(1000);
        okBtn.click();
        Thread.sleep(2000);

        List<WebElement> elements = driver.findElements(By.cssSelector("div.film-delete[data-id='" + idFilm + "']"));
        if (!elements.isEmpty()) {
            System.out.println("Có film: " + idFilm);
            return false;
        } else {
            System.out.println("Không còn film");
            return true;
        }
    }

    List<String> actualLoginList = new ArrayList<>();
    List<Boolean> expectedLoginInExcel = readExcelColumnAsBoolean(filePath, SheetName, startRow, 4, totalRowData);
    int index1 = 0;
    int index2 = 0;
    int rowNumber = 0;

    @Test(dataProvider = "removeFilm")
    public void testMultiRegister(String idFilm) throws InterruptedException, IOException {
        boolean actualResult = TestRemoveFilm(idFilm);
        try {
            if (actualResult == expectedLoginInExcel.get(index1++)) {
                actualLoginList.add("PASS");
            } else {
                actualLoginList.add("FAIL");
                List<String> rowValues = getRowValues(filePath, SheetName, startWriteRowExcel, 0, 6);
                FileInputStream file = new FileInputStream(new File(filePath));
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheet("Test Defect");
                int lastRowNum = sheet.getLastRowNum();
                Row row = sheet.createRow(lastRowNum + 1);
                if (lastRowNum > 0) {
                    Row previousRow = sheet.getRow(lastRowNum);
                    boolean hasValue = false;
                    for (Cell cell : previousRow) {
                        if (cell.getCellType() != CellType.BLANK) {
                            hasValue = true;
                            break;
                        }
                    }
                    if (hasValue) {
                        row = sheet.createRow(lastRowNum + 1);
                    }
                }
                for (int i = 0; i < rowValues.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(rowValues.get(i));
                }
                FileOutputStream outputStream = new FileOutputStream(filePath);
                workbook.write(outputStream);
                workbook.close();
                outputStream.close();
                startWriteRowExcel++;
            }
            Assert.assertEquals(expectedLoginInExcel.get(index2++), actualResult);
        } catch (AssertionError ae) {
            FileInputStream file1 = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook1 = new XSSFWorkbook(file1);
            XSSFSheet sheet1 = workbook1.getSheet("Test Defect");

            Row row = sheet1.getRow(rowNum);
            if (row == null) {
                row = sheet1.createRow(rowNum);
            }
            int colNum = 6;
            Cell cell = row.getCell(colNum);
            while (cell != null && !cell.getStringCellValue().isEmpty()) {
                rowNum++;
                row = sheet1.getRow(rowNum);
                if (row == null) {
                    row = sheet1.createRow(rowNum);
                }
                cell = row.getCell(colNum);
            }
            cell = row.getCell(colNum);
            if (cell == null) {
                cell = row.createCell(colNum);
            }
            cell.setCellValue(ae.getMessage().trim());
            try {
                FileOutputStream outputStream = new FileOutputStream(filePath);
                workbook1.write(outputStream);
                workbook1.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    @AfterTest
    public void tearDown() throws InterruptedException {
        System.out.println(actualLoginList.get(0));
        try {
            FileInputStream file = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet(SheetName);
            for (int i = startRow; i < startRow + actualLoginList.size(); i++) {
                Cell resultCell = sheet.getRow(i).getCell(5);
                resultCell.setCellValue(actualLoginList.get(i - startRow));
            }
            FileOutputStream outFile = new FileOutputStream(new File(filePath));
            workbook.write(outFile);
            outFile.close();
        } catch (FileNotFoundException fnfe) {
            fnfe.printStackTrace();
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
//        Thread.sleep(4000);
        driver.close();
        driver.quit();
    }
}
