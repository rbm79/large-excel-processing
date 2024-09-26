package com.example.excel;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class LargeExcelProcessor {

    public void processLargeFile(String filename) {

        ZipSecureFile.setMinInflateRatio(0);
        // Definindo o limite de leitura para dados binários maiores
        IOUtils.setByteArrayMaxOverride(100000000); // Aumenta para 200 MB

        try (FileInputStream fis = new FileInputStream(filename);
             XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
             SXSSFWorkbook workbook = new SXSSFWorkbook(xssfWorkbook, 100)) {

            // Configuração para economizar espaço em disco
            workbook.setCompressTempFiles(true);

            // Obtendo a primeira planilha
            Sheet sheet = workbook.getSheetAt(0);

            // Iterando sobre as linhas
            for (Row row : sheet) {
                processRow(row);
            }

            // Libera recursos temporários
            workbook.dispose();

            System.out.println("Processamento de arquivo grande concluído com sucesso!");

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
