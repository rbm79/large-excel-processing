package com.example.excel;

import java.io.IOException;

public class App {
    public static void main(String[] args) {
        String smallFile = "C:\\repositorios\\excel-processing\\teste.xlsx";
        String largeFile = "C:\\repositorios\\excel-processing\\amostra800k-linhas-sucesso-para-boleto-40k1.xlsx";

        FastExcelReaderLowMemory readerThread = new FastExcelReaderLowMemory();
        try {
            readerThread.processFile(largeFile);
        } catch (IOException e) {
            e.printStackTrace();
        }
         

        /*FastExcelReaderBatchSeq reader = new FastExcelReaderBatchSeq();
        try {
            reader.processFile(largeFile);
        } catch (IOException e) {
            e.printStackTrace();
        } */


        // Processando arquivo pequeno
        //SmallExcelProcessor smallProcessor = new SmallExcelProcessor();
        //smallProcessor.processSmallFile(smallFile);

        // Processando arquivo grande
        //LargeExcelProcessor largeProcessor = new LargeExcelProcessor();
        //largeProcessor.processLargeFile(largeFile);
    }
}
