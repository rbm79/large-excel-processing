package com.example.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.TimeUnit;
import java.util.function.Consumer;
import java.util.concurrent.atomic.AtomicLong;
import org.apache.commons.lang3.time.StopWatch;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.dhatim.fastexcel.reader.Cell;

public class FastExcelReaderLowMemory {

    private static final int LOG_INTERVAL = 10000;

    // Método que aceita apenas o caminho do arquivo
    public void processFile(String filePath) throws IOException {
        processFile(filePath, this::defaultRowProcessor);
    }

    // Método que aceita o caminho do arquivo e um Consumer personalizado
    public void processFile(String filePath, Consumer<String> rowProcessor) throws IOException {
        System.out.println("Iniciando processamento do arquivo: " + filePath);
        
        StopWatch watch = new StopWatch();
        watch.start();

        AtomicLong totalRowsProcessed = new AtomicLong(0);
        AtomicLong sheetCount = new AtomicLong(0);

        try (InputStream is = new FileInputStream(filePath);
             ReadableWorkbook wb = new ReadableWorkbook(is)) {

            wb.getSheets().forEach(sheet -> {
                sheetCount.incrementAndGet();
                System.out.println("Processando planilha: " + sheet.getName());
                
                try {
                    long rowsInSheet = processSheet(sheet, rowProcessor);
                    totalRowsProcessed.addAndGet(rowsInSheet);
                    
                    System.out.println("Planilha processada: " + sheet.getName() + ", Linhas: " + rowsInSheet);
                } catch (IOException e) {
                    throw new RuntimeException("Erro ao processar planilha: " + sheet.getName(), e);
                }
            });
        }

        watch.stop();
        System.out.println("Processamento concluído.");
        System.out.println("Total de planilhas processadas: " + sheetCount.get());
        System.out.println("Total de linhas processadas: " + totalRowsProcessed.get());
        System.out.println("Tempo de processamento: " + watch.getTime(TimeUnit.SECONDS) + " segundos");
    }

    private long processSheet(Sheet sheet, Consumer<String> rowProcessor) throws IOException {
        long rowCount = 0;
        
        for (Row row : sheet.read()) {
            if (row.getRowNum() == 0) continue; // Pula a primeira linha (cabeçalho)

            processRow(row, rowProcessor);
            rowCount++;

            if (rowCount % LOG_INTERVAL == 0) {
                System.out.println("Processadas " + rowCount + " linhas na planilha " + sheet.getName());
            }
        }
        
        return rowCount;
    }

    private void processRow(Row row, Consumer<String> rowProcessor) {
        StringBuilder rowOutput = new StringBuilder(128);
        rowOutput.append("Row ").append(row.getRowNum()).append(": ");
        
        for (int i = 0; i < row.getCellCount(); i++) {
            if (i > 0) rowOutput.append(", ");
            rowOutput.append("Field ").append(i + 1).append(" = ")
                     .append(getCellValueAsString(row.getCell(i)));
        }
        
        rowProcessor.accept(rowOutput.toString());
    }

    private void defaultRowProcessor(String rowOutput) {
        // Este é o processador padrão. Você pode querer ajustá-lo conforme necessário.
        System.out.println(rowOutput);
    }

    private String getCellValueAsString(Cell cell) {
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
                return cell.getFormula();
            default:
                return "";
        }
    }
}
