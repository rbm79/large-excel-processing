package com.example.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.apache.commons.lang3.time.StopWatch;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;

public class FastExcelReaderBatch {

    // Tamanho do lote para processamento. Pode ser ajustado para otimizar o desempenho.
    private static final int BATCH_SIZE = 1000;

    public void processFile(String filePath) throws IOException {
        // Abre o arquivo Excel usando FileInputStream
        try (InputStream is = new FileInputStream(filePath);
             ReadableWorkbook wb = new ReadableWorkbook(is)) {

            // Inicia o cronômetro para medir o tempo de processamento
            StopWatch watch = new StopWatch();
            watch.start();

            // Processa cada planilha no arquivo Excel
            wb.getSheets().forEach(sheet -> {
                try {
                    processSheet(sheet);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });

            // Para o cronômetro e imprime o tempo total de processamento
            watch.stop();
            System.out.println("Processing time :: " + watch.getTime(TimeUnit.MINUTES));
        }
    }

    private void processSheet(Sheet sheet) throws IOException {
        List<List<Row>> batches = new ArrayList<>();
        List<Row> currentBatch = new ArrayList<>();
        
        // Lê as linhas da planilha e as agrupa em lotes
        for (Row row : sheet.read()) {
            if (row.getRowNum() == 0) continue; // Pula a primeira linha (cabeçalho)
            currentBatch.add(row);
            if (currentBatch.size() == BATCH_SIZE) {
                batches.add(currentBatch);
                currentBatch = new ArrayList<>();
            }
        }
        // Adiciona o último lote se não estiver vazio
        if (!currentBatch.isEmpty()) {
            batches.add(currentBatch);
        }

        // Processa os lotes em paralelo para melhor desempenho
        batches.parallelStream().forEach(this::processBatch);
    }

    private void processBatch(List<Row> batch) {
        List<String> output = new ArrayList<>();
        for (Row r : batch) {
            // Extrai os dados de cada célula
            BigDecimal field1 = r.getCellAsNumber(0).orElse(null);
            String field2 = r.getCellAsString(1).orElse(null);
            // Formata a saída para cada linha
            output.add(String.format("Row %d: Field 1 = %s, Field 2 = %s", 
                                     r.getRowNum(), field1, field2));
        }
        // Imprime todas as linhas do lote de uma vez, reduzindo as operações de I/O
        System.out.println(String.join("\n", output));
    }
}
