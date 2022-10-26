package com.example.uploadexcelfileservice.helper;

import com.example.uploadexcelfileservice.model.Employee;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class EmployeeHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] HEADERs = { "Id", "First Name", "Last Name", "Gender", "Country", "Age" };
    static String SHEET = "Employees";

    public static boolean hasExcelFormat(MultipartFile file){
        if (!TYPE.equals(file.getContentType())){
            return false;
        }
        return true;
    }

    public static List<Employee> uploadFileExcel(InputStream request){
        try {
            Workbook workbook = WorkbookFactory.create(request);
            Sheet sheet = workbook.getSheetAt(0);

            Iterator<Row> rows = sheet.iterator();

            List<Employee> employees = new ArrayList<Employee>();

            int rowNumber = 0;
            while (rows.hasNext()){
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();
                Employee employee = new Employee();

                int cellId = 0;
                while (cellsInRow.hasNext()){
                    Cell currentCell = cellsInRow.next();

                    switch (cellId){
                        case 0:
                            employee.setId((long) currentCell.getNumericCellValue());
                            break;
                        case 1:
                            employee.setFirstName(currentCell.getStringCellValue());
                            break;
                        case 2:
                            employee.setLastName(currentCell.getStringCellValue());
                            break;
                        case 3:
                            employee.setGender(currentCell.getStringCellValue());
                            break;
                        case 4:
                            employee.setCountry(currentCell.getStringCellValue());
                            break;
                        case 5:
                            employee.setAge((long) currentCell.getNumericCellValue());
                            break;
                        default:
                            break;
                    }
                    cellId++;
                }
                employees.add(employee);
            }
            workbook.close();
            return employees;
        }catch (IOException e){
            throw new RuntimeException("fail to parse Excel file: "+ e.getMessage());
        }
    }

    public static ByteArrayInputStream downloadFileExcel(List<Employee> employees){
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream output = new ByteArrayOutputStream();){
            Sheet sheet = workbook.createSheet(SHEET);

            // Header
            Row headerRow = sheet.createRow(0);

            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }


            int rowIdx = 1;
            for (Employee employee : employees) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(employee.getId());
                row.createCell(1).setCellValue(employee.getFirstName());
                row.createCell(2).setCellValue(employee.getLastName());
                row.createCell(3).setCellValue(employee.getGender());
                row.createCell(4).setCellValue(employee.getCountry());
                row.createCell(5).setCellValue(employee.getAge());
            }

            workbook.write(output);
            return new ByteArrayInputStream(output.toByteArray());

        }catch (IOException e){
            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
        }
    }
}
