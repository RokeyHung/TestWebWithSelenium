package StartTest;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
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
import static StartTest.ExcelReader.getStackTraceAsString;

public class ChangePasswordTest {
    private WebDriver driver;
    private final String filePath = "DataTest.xlsx";
    private final String SheetName = "Test Case";
    int startRow = 19;
    int totalRowData = 1;

    public ChangePasswordTest() throws IOException {
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
    }

    @DataProvider(name = "register")
    public Object[][] data() throws Exception {
        return getDataCellInColl(filePath, SheetName, startRow, 3, totalRowData);
    }

    @Test(dataProvider = "register")
    public boolean TestRegister(String old_password, String password, String password_confirmation) throws InterruptedException {
        WebElement avt = driver.findElement(By.className("navbar-avatar"));
        avt.click();

        WebElement usernameField = driver.findElement(By.name("username"));
        WebElement passwordField = driver.findElement(By.name("password"));
        WebElement loginButton = driver.findElement(By.id("login"));

        usernameField.sendKeys("huyhy03");
        passwordField.sendKeys(old_password);
        loginButton.click();
        Thread.sleep(1700);

        WebElement avt1 = driver.findElement(By.className("navbar-avatar"));
        avt1.click();
        Thread.sleep(1700);

        WebElement element = driver.findElement(By.xpath("//div[@class='user-item']//a[span[text()='Đổi mật khẩu']]"));
        element.click();

        WebElement change_old_passwordField = driver.findElement(By.name("old_password"));
        WebElement change_passwordField = driver.findElement(By.name("password"));
        WebElement change_password_confirmationField = driver.findElement(By.name("password_confirmation"));

        change_old_passwordField.sendKeys(old_password);
        change_passwordField.sendKeys(password);
        change_password_confirmationField.sendKeys(password_confirmation);

        WebElement changeButton = driver.findElement(By.className("navbar-form-group"));
        Thread.sleep(1700);
        changeButton.click();

        WebElement avt2 = driver.findElement(By.className("navbar-avatar"));
        avt2.click();

        WebElement logout = driver.findElement(By.className("logout"));
        Thread.sleep(1700);
        logout.click();

        Thread.sleep(1000);
        WebElement avt3 = driver.findElement(By.className("navbar-avatar"));
        Thread.sleep(1700);
        avt3.click();

        WebElement usernameField1 = driver.findElement(By.name("username"));
        WebElement passwordField1 = driver.findElement(By.name("password"));
        WebElement loginButton1 = driver.findElement(By.id("login"));
        usernameField1.sendKeys("huyhy03");
        passwordField1.sendKeys(password);
        loginButton1.click();
        Thread.sleep(1700);

        return driver.manage().getCookieNamed("remember_web_59ba36addc2b2f9401580f014c7f58ea4e30989d") != null;
    }

    List<String> actualLoginList = new ArrayList<>();
    List<Boolean> expectedLoginInExcel = readExcelColumnAsBoolean(filePath, SheetName, startRow, 4, totalRowData);
    int index1 = 0;
    int index2 = 0;
    int rowNumber = 0;

    @Test(dataProvider = "register")
    public void testMultiRegister(String old_password, String password, String password_confirmation) throws InterruptedException, IOException {
        boolean actualResult = TestRegister(old_password, password, password_confirmation);
        try {
            if (actualResult == expectedLoginInExcel.get(index1++)) {
                actualLoginList.add("PASS");
            } else {
                actualLoginList.add("FAIL");
            }
            Assert.assertEquals(expectedLoginInExcel.get(index2++), actualResult);
        } catch (Exception e) {
            FileInputStream file = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet("Test Defect");
            while (sheet.getRow(rowNumber) != null) {
                rowNumber++;
            }
            Row row = sheet.createRow(rowNumber);
            Cell cell = row.createCell(0);
            cell.setCellValue(getStackTraceAsString(e));
            System.out.println(getStackTraceAsString(e) + "123");

            try {
                FileOutputStream outputStream = new FileOutputStream(filePath);
                workbook.write(outputStream);
                workbook.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    @AfterTest
    public void tearDown() throws InterruptedException {
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
