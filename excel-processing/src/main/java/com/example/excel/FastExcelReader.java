package com.example.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.concurrent.TimeUnit;
import java.util.stream.Stream;
import org.apache.commons.lang3.time.StopWatch;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;

public class FastExcelReader {

    public void processFile(String filePath) throws IOException {
            
        try (InputStream is = new FileInputStream(filePath);
        ReadableWorkbook wb = new ReadableWorkbook(is)) {

        StopWatch watch = new StopWatch();
        watch.start();
        
        wb.getSheets().forEach(sheet ->
        {
            try (Stream<Row> rows = sheet.openStream()) {
            
                System.out.println("discarding first row as this is just the header");
                rows.skip(1).forEach(r -> {
                    
                    BigDecimal field1 = r.getCellAsNumber(0).orElse(null);
                    String field2 = r.getCellAsString(1).orElse(null);
                    
                    System.out.println("Field 1 str value :: " + field1);
                    System.out.println("Field 2 str value :: " + field2);
                    System.out.println("Row toString - " + r.toString());
                    System.out.println("Row getRowNum - "+r.getRowNum()); 
                    System.out.println("====================================");
                    
                });

            } catch (Exception e) {
                e.printStackTrace();
            }

            watch.stop();
            System.out.println("Processing time :: " + watch.getTime(TimeUnit.MINUTES));
        });

        }

    }

}