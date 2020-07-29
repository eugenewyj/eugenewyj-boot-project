package org.eugenewyj.boot.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * ExcelUtilTest
 *
 * @author eugene
 * @date 2020/7/29
 */

class ExcelUtilTests {

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
            ExcelUtil.<TestRecord>export(out, records);
        }
    }

    @ExcelSheet("测试sheet")
    static class TestRecord {
        @ExcelColumn
        private int a;
        @ExcelColumn(order = 2, value = "列2")
        private String b;

        public TestRecord(int a, String b) {
            this.a = a;
            this.b = b;
        }
    }
}
