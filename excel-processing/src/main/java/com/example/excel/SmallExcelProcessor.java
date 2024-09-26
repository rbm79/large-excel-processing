package com.example.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class SmallExcelProcessor {
    
    public void processSmallFile(String filename) {
        try (FileInputStream fis = new FileInputStream(filename);
             Workbook workbook = WorkbookFactory.create(fis)) {

            // Obtendo a primeira planilha
            Sheet sheet = workbook.getSheetAt(0);

            // Iterando sobre as linhas
            for (Row row : sheet) {
                processRow(row);
            }

            System.out.println("Processamento concluído com sucesso!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void processRow(Row row) {
        // Processa cada célula de cada linha
        for (Cell cell : row) {
            switch (cell.getCellType()) {
                case STRING:
                    System.out.print(cell.getStringCellValue() + "\t");
                    break;
                case NUMERIC:
                    System.out.print(cell.getNumericCellValue() + "\t");
                    break;
                case BOOLEAN:
                    System.out.print(cell.getBooleanCellValue() + "\t");
                    break;
                default:
                    System.out.print("Unknown" + "\t");
            }
        }
        System.out.println(); // Nova linha após cada row
    }
}
