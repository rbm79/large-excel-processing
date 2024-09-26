package com.example.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.commons.lang3.time.StopWatch;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.dhatim.fastexcel.reader.Cell;

public class FastExcelReaderThread {

    private static final int BATCH_SIZE = 1000;

    public void processFile(String filePath) throws IOException {
        try (InputStream is = new FileInputStream(filePath);
             ReadableWorkbook wb = new ReadableWorkbook(is)) {

            StopWatch watch = new StopWatch();
            watch.start();

            wb.getSheets().forEach(sheet -> {
                try {
                    processSheet(sheet);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });

            watch.stop();
            System.out.println("Processing time :: " + watch.getTime(TimeUnit.SECONDS));
        }
    }

    private void processSheet(Sheet sheet) throws IOException {
        List<Row> currentBatch = new ArrayList<>();
        
        for (Row row : sheet.read()) {
            if (row.getRowNum() == 0) continue; // Pula a primeira linha (cabeçalho)
            currentBatch.add(row);
            if (currentBatch.size() == BATCH_SIZE) {
                processBatch(currentBatch);
                currentBatch.clear();
            }
        }
        // Processa o último lote se não estiver vazio
        if (!currentBatch.isEmpty()) {
            processBatch(currentBatch);
        }
    }

    private void processBatch(List<Row> batch) {
        StringBuilder output = new StringBuilder();
        for (Row r : batch) {
            String field1 = getCellValueAsString(r, 0);
            String field2 = getCellValueAsString(r, 1);
            output.append(String.format("Row %d: Field 1 = %s, Field 2 = %s%n", 
                                        r.getRowNum(), field1, field2));
        }
        // Imprime todas as linhas do lote de uma vez
        System.out.print(output);
    }

    private String getCellValueAsString(Row row, int columnIndex) {
        if (row == null || columnIndex < 0 || columnIndex >= row.getCellCount()) {
            return "";
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return "";
        }
        switch (cell.getType()) {
            case NUMBER:
                return cell.asNumber().toString();
            case STRING:
                return cell.asString();
            case BOOLEAN:
                return Boolean.toString(cell.asBoolean());
            case FORMULA:
                // Para fórmulas, você pode querer avaliar o resultado ou retornar a fórmula
                return cell.getFormula();
            default:
                return "";
        }
    }
}
