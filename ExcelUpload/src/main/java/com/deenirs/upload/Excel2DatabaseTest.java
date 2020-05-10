package com.deenirs.upload;

import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
 
/**
 * Sample Java program that imports data from an Excel file to MySQL database.
 *
 * @author Nam Ha Minh - https://www.codejava.net
 *
 */
public class Excel2DatabaseTest {
 
    public static void main(String[] args) {
      //  String jdbcURL = "jdbc:mysql://localhost:3306/sales";
        String jdbcURL = "jdbc:mysql://localhost:3306/generic?createDatabaseIfNotExist=true";
        String username = "root";
        String password = "*****@**";
 
        String excelFilePath = "F:/Email data.xls";
 
        int batchSize = 30;
 
        Connection connection = null;
 
        try {
            long start = System.currentTimeMillis();
             
            FileInputStream inputStream = new FileInputStream(excelFilePath);
 
            HSSFWorkbook   workbook = new HSSFWorkbook(inputStream);
 
            HSSFSheet  firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
 
            connection = DriverManager.getConnection(jdbcURL, username, password);
            connection.setAutoCommit(false);
  
            String sql = "INSERT INTO Emp_details (companyname, name, job,email) VALUES (?, ?, ?,?)";
            PreparedStatement statement = connection.prepareStatement(sql);    
             
            int count = 0;
             
            rowIterator.next(); // skip the header row
             
            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
 
                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();
 
                    int columnIndex = nextCell.getColumnIndex();
 
                    switch (columnIndex) {
                    case 0:
                        String companyname = nextCell.getStringCellValue();
                        statement.setString(1, companyname);
                        break;
                    case 1:
                    	String name = nextCell.getStringCellValue();
                        statement.setString(2, name);
                       // break;
                    case 2:
                        String job =  nextCell.getStringCellValue();
                        statement.setString(3, job);
                      //  break;
                    case 3:
                        String email =  nextCell.getStringCellValue();
                        statement.setString(4, email);
                       // break;
                    }
 
                }
                 
                statement.addBatch();
                 
                if (count % batchSize == 0) {
                    statement.executeBatch();
                }              
 
            }
 
            workbook.close();
             
            // execute the remaining queries
            statement.executeBatch();
  
            connection.commit();
            connection.close();
             
            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));
             
        } catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }
 
    }
}
