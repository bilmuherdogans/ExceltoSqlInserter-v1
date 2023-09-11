package com.github.bilmuherdogan.exceltosqlinserter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class ExceltoSqlInserter {

    public static void main(String[] args) {
        File excelFilePath = new File("D:\\Dev\\sts-workbench\\ExcelToSQLInserter\\src\\main\\java\\com\\github\\bilmuherdogan\\exceltosqlinserter\\examplesData.xlsx");
        
        StringBuilder insertQueries = new StringBuilder();
        FileInputStream excelFile = null;
        Workbook workbook = null;

        try {
            excelFile = new FileInputStream(excelFilePath);
            workbook = new XSSFWorkbook(excelFile);
            
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                Cell keyCell = row.getCell(0);
                Cell valueCell = row.getCell(1);

                if (isValidRow(keyCell, valueCell)) {
                    double numericKey = keyCell.getNumericCellValue();
                    String stringValue = valueCell.getStringCellValue();

                    String insertSQL = formatInsertQuery((int) numericKey, stringValue);
                    insertQueries.append(insertSQL).append("\n");
                }
            }

            saveQueriesToFile(insertQueries.toString(), excelFilePath.getParent());
            System.out.println("Querys were saved.");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (excelFile != null) {
                    excelFile.close();
                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    private static boolean isValidRow(Cell keyCell, Cell valueCell) {
        return keyCell != null && valueCell != null
                && keyCell.getCellType() == CellType.NUMERIC
                && valueCell.getCellType() == CellType.STRING;
    }

    private static String formatInsertQuery(int numericKey, String stringValue) {
        return String.format(
                "INSERT INTO ABCCONF.XYZ (ID, CODE, KEY, VALUE) VALUES (ABCCONF.SEQ_XYZ.nextval, 'ABC_REASON', '%d', '%s');",
                numericKey,
                stringValue.replace("'", "''"));
    }

    private static void saveQueriesToFile(String queries, String directory) {
        File outputFile = new File(directory, "queries.sql");
        PrintWriter out = null;
        try {
            out = new PrintWriter(new FileOutputStream(outputFile));
            out.println(queries);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                out.close();
            }
        }
    }
}