package org.eugenewyj.boot.poi.sample;

import org.eugenewyj.boot.poi.SimpleExcelExporter;
import org.eugenewyj.boot.poi.annotation.ExcelColumn;
import org.eugenewyj.boot.poi.annotation.ExcelSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * SimpleExcelExporterSample
 *
 * @author eugene
 * @date 2020/7/29
 */

class SimpleExcelExporterSample {
    private static final Logger logger = LoggerFactory.getLogger(SimpleExcelExporterSample.class);

    /**
     * 测试导出
     * @param args
     * @throws IOException
     * @throws IllegalAccessException
     */
    public static void main(String[] args) throws IOException, IllegalAccessException {
        List<TestRecord> records = new ArrayList<>();
        records.add(new TestRecord(1, "测试1"));
        records.add(new TestRecord(2, "测试2"));
        try (FileOutputStream out = new FileOutputStream(new File("test.xlsx"))) {
            SimpleExcelExporter exporter = new SimpleExcelExporter(TestRecord.class);
            exporter.export(out, records);
        }

    }

    @ExcelSheet("测试sheet")
    static class TestRecord {
        @ExcelColumn
        private int a;
        @ExcelColumn(order = 2, value = "列2")
        private String b;
        @ExcelColumn(order = 3, value = "列3")
        private Integer c = 0;
        @ExcelColumn
        private double d = 0.001;

        public TestRecord(int a, String b) {
            this.a = a;
            this.b = b;
        }
    }
}
