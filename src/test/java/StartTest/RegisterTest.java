package StartTest;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import static StartTest.ExcelReader.*;

public class RegisterTest {
    private WebDriver driver;
    private final String filePath = "DataTest.xlsx";
    private final String SheetName = "Test Case";
    int startRow = 9;
    int totalRowData = 9;
    int startWriteRowExcel = startRow;
    int rowNum = 0;

    public RegisterTest() throws IOException {
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
        Thread.sleep(1000);
        WebElement signupNavbar = driver.findElement(By.className("navbar-tab-signup"));
        signupNavbar.click();
    }

    @DataProvider(name = "register")
    public Object[][] data() throws Exception {
        return getDataCellInColl(filePath, SheetName, startRow, 3, totalRowData);
    }

    @Test(dataProvider = "register")
    public boolean TestRegister(String username, String password, String confirm, String full_name, String email) throws InterruptedException {
        WebElement usernameInput = driver.findElement(By.className("tab-signup")).findElement(By.name("username"));
        WebElement passwordInput = driver.findElement(By.className("tab-signup")).findElement(By.name("password"));
        WebElement password_confirmInput = driver.findElement(By.className("tab-signup")).findElement(By.name("password_confirm"));
        WebElement full_nameInput = driver.findElement(By.className("tab-signup")).findElement(By.name("full_name"));
        WebElement emailInput = driver.findElement(By.className("tab-signup")).findElement(By.name("email"));

        usernameInput.sendKeys(username);
        passwordInput.sendKeys(password);
        password_confirmInput.sendKeys(confirm);
        full_nameInput.sendKeys(full_name);
        emailInput.sendKeys(email);

        WebElement element = driver.findElement(By.id("signup"));
        Thread.sleep(1500);
        element.click();

        WebElement element1 = driver.findElement(By.id("signup"));
        Thread.sleep(1500);
        element1.click();

        if (indexExpected != 9) {
            Actions actions = new Actions(driver);
            usernameInput.sendKeys("");
            actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform();
            passwordInput.sendKeys("");
            actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform();
            password_confirmInput.sendKeys("");
            actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform();
            full_nameInput.sendKeys("");
            actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform();
            emailInput.sendKeys("");
            actions.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).sendKeys(Keys.DELETE).perform();
            indexExpected++;
        } else {
            Thread.sleep(13000);
        }
        return driver.manage().getCookieNamed("remember_web_59ba36addc2b2f9401580f014c7f58ea4e30989d") != null;
    }

    List<String> actualList = new ArrayList<>();
    List<Boolean> expectedLoginInExcel = readExcelColumnAsBoolean(filePath, SheetName, startRow, 4, totalRowData);
    int index1 = 0;
    int index2 = 0;
    int indexExpected = 1;

    @Test(dataProvider = "register")
    public void testMultiRegister(String username, String password, String confirm, String full_name, String email) throws InterruptedException, IOException {
        boolean actualResult = TestRegister(username, password, confirm, full_name, email);
        try {
            if (actualResult == expectedLoginInExcel.get(index1++)) {
                actualList.add("PASS");
            } else {
                actualList.add("FAIL");
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
    public void tearDown() {
        try {
            FileInputStream file = new FileInputStream(new File(filePath));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheet(SheetName);
            for (int i = startRow; i < startRow + actualList.size(); i++) {
                Cell resultCell = sheet.getRow(i).getCell(5);
                resultCell.setCellValue(actualList.get(i - startRow));
            }
            FileOutputStream outFile = new FileOutputStream(new File(filePath));
            workbook.write(outFile);
            outFile.close();
        } catch (FileNotFoundException fnfe) {
            fnfe.printStackTrace();
        } catch (IOException ioe) {
            ioe.printStackTrace();
        }
        driver.close();
        driver.quit();
    }
}
