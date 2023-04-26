package StartTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelReader {

    public static void main(String[] args) throws IOException {
        String filePath = "DataTest.xlsx";
        String SheetName = "Test Case";

        Object[][] data = getDataCellInColl(filePath, SheetName, 2, 3, 6);

        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[i].length; j++) {
                System.out.print(data[i][j] + " ");
            }
            System.out.println();
        }
//        for (boolean value : readExcelColumnAsBoolean(filePath, SheetName, 1, 2, 6)) {
//            System.out.println(value);
//        }
    }

    /**
     * Trả về dấu vết ngăn xếp của một ngoại lệ dưới dạng một chuỗi.
     *
     * @param throwable ném ngoại lệ
     * @return dấu vết ngăn xếp dưới dạng chuỗi
     */
    public static String getStackTraceAsString(Throwable throwable) {
        StringBuilder stringBuilder = new StringBuilder();

        for (StackTraceElement stackTraceElement : throwable.getStackTrace()) {
            stringBuilder.append(stackTraceElement.toString());
            stringBuilder.append("\n");
        }

        return stringBuilder.toString();
    }

    public static Object[][] getDataTable(String filePath, String SheetName, int startRow, int startCol, int totalRows, int totalCols) throws IOException {
        FileInputStream file = new FileInputStream(new File(filePath));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(SheetName);

        String[][] data = new String[totalRows][totalCols];

        // Duyệt từng hàng và cột để lấy dữ liệu
        for (int i = startRow; i < startRow + totalRows; i++) {
            Row row = sheet.getRow(i);
            for (int j = startCol; j < startCol + totalCols; j++) {
                Cell cell = row.getCell(j);
                if (cell.getCellType() == CellType.NUMERIC) {
                    // Xử lý giá trị số ở đây
                    double numericValue = cell.getNumericCellValue();
                    data[i - startRow][j - startCol] = String.valueOf(numericValue);
                } else {
                    // Xử lý giá trị chuỗi ở đây
                    String stringValue = cell.getStringCellValue();
                    data[i - startRow][j - startCol] = stringValue;
                }
            }
        }
        workbook.close();
        file.close();
        return data;
    }

    public static Object[][] getDataBooleanTable(String fileName, String SheetName, int startRow, int startCol, int numRows, int numCols, int booleanColIndex) throws IOException {
        FileInputStream inputStream = new FileInputStream(new File(fileName));
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(SheetName);

        Object[][] data = new Object[numRows][numCols];

        int rowIndex = 0;
        for (int i = startRow; i < startRow + numRows; i++) {
            Row row = sheet.getRow(i);
            int colIndex = 0;
            if (row != null) {
                // Lặp qua các cột của hàng để lấy dữ liệu
                for (int j = startCol; j < startCol + numCols; j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        // Ép kiểu giá trị cell sang kiểu String
                        DataFormatter formatter = new DataFormatter();
                        String value = formatter.formatCellValue(cell);

                        if (j == booleanColIndex) {
                            data[rowIndex][colIndex] = Boolean.parseBoolean(value);
                        } else {
                            data[rowIndex][colIndex] = value;
                        }
                    } else {
                        data[rowIndex][colIndex] = "";
                    }
                    colIndex++;
                }
            }
            rowIndex++;
        }
        workbook.close();
        inputStream.close();
        return data;
    }

    public static Object[][] getDataCellInColl(String filePath, String SheetName, int startRow, int startCol, int totalRows) throws IOException {
        Object[][] dataNoSplit = getDataTable(filePath, SheetName, startRow, startCol, totalRows, 1);

        Object[][] data = new Object[dataNoSplit.length][];
        for (int i = 0; i < dataNoSplit.length; i++) {
            String s = dataNoSplit[i][0].toString();
            String[] parts = s.split("'");
            data[i] = new Object[parts.length / 2]; // tạo mảng con với độ dài bằng một nửa số phần tử của mảng parts
            for (int j = 0; j < parts.length; j += 2) {
                data[i][j / 2] = parts[j + 1];
            }
        }
        return data;
    }

    public static Object[][] readExcelColumn(String filePath, String SheetName, int startRow, int startCol, int totalRows) throws IOException {
        return getDataTable(filePath, SheetName, startRow, startCol, totalRows, 1);
    }

    public static List<Boolean> readExcelColumnAsBoolean(String filePath, String SheetName, int startRow, int startColumn, int totalRows) throws IOException {
        List<Boolean> columnData = new ArrayList<>();
        try (Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath))) {
            Sheet sheet = workbook.getSheet(SheetName);
            for (int i = startRow; i < startRow + totalRows; i++) {
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(startColumn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                boolean value = cell.getBooleanCellValue();
                columnData.add(value);
            }
        }
        return columnData;
    }

    public static List<List<String>> getDataFromExcel(String filePath, String sheetName, int startRow, int startCol, int totalRows, int totalCols) throws Exception {
        List<List<String>> data = new ArrayList<>();
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);

        for (int i = startRow; i < startRow + totalRows; i++) {
            Row row = sheet.getRow(i);
            List<String> rowData = new ArrayList<>();
            for (int j = startCol; j < startCol + totalCols; j++) {
                Cell cell = row.getCell(j);
                rowData.add(cell.getStringCellValue());
            }
            data.add(rowData);
        }

        workbook.close();
        inputStream.close();
        return data;
    }

    public static void writeDataToExcel(String filePath, String sheetName, List<List<String>> data) throws Exception {
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);

        // Kiểm tra nếu hàng đầu tiên không rỗng, tìm hàng trống tiếp theo để ghi dữ liệu
        int rowNum = 0;
        Row firstRow = sheet.getRow(rowNum);
        Cell firstCell = firstRow.getCell(0);
        while (firstCell != null && !firstCell.getStringCellValue().isEmpty()) {
            rowNum++;
            firstRow = sheet.getRow(rowNum);
            firstCell = firstRow.getCell(0);
        }

        // Ghi dữ liệu vào hàng tiếp theo
        for (List<String> rowData : data) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (String cellData : rowData) {
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(cellData);
            }
        }

        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream(filePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    public static List<String> getRowValues(String filePath, String sheetName, int startRow, int startCol, int totalCols) throws IOException {
        List<String> rowValues = new ArrayList<>();
        FileInputStream inputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);
        Iterator<Row> rowIterator = sheet.iterator();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row.getRowNum() < startRow) {
                continue;
            }
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if (cell.getColumnIndex() < startCol) {
                    continue;
                }
                if (cell.getColumnIndex() >= startCol + totalCols) {
                    break;
                }
                rowValues.add(cell.toString());
            }
            break;
        }
        workbook.close();
        return rowValues;
    }


    public void ObjectTest() {
        Object[][] dataNoSplit = {
                {"Tài khoản: 'huyhy03' \n Mật khẩu: '123asd123'"},
                {"Tài khoản: 'huyhy03' \n Mật khẩu: 123456"}
        };
        Object[][] data = {
                {"huyhy03", "123asd132"},
                {"huyhy03", 123456}
        };
        Object[][] dataTrueFalseNoSplit = {
                {"Tài khoản: 'huyhy03' Mật khẩu: '6B' Email: 'hhungnm@gmail.com'", "FALSE"},
                {"Tài khoản: 'huyhy03' Mật khẩu: '123asd123'", "TRUE"}
        };
        Object[][] dataTrueFalse = {
                {"huyhy03", "6B", "hhungnm@gmail.com", "FALSE"},
                {"huyhy03", "123asd123", "FALSE"},
        };
    }
}
