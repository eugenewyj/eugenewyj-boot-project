package org.eugenewyj.boot.poi;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.eugenewyj.boot.poi.annotation.ExcelColumn;
import org.eugenewyj.boot.poi.annotation.ExcelSheet;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;

/**
 * SimpleExcelExporterTest
 *
 * @author eugene
 * @date 2020/8/2
 */
public class SimpleExcelExporterTest {
    private static final String sheetName = "导出测试sheet";
    /**
     * 测试成功导出Workbook。
     */
    @Test
    void testSuccessExport() throws IllegalAccessException {
        SimpleExcelExporter exporter = new SimpleExcelExporter(SuccessRecord.class);
        List<SuccessRecord> records = Arrays.asList(
                new SuccessRecord(1, BigDecimal.valueOf(0.01), "行1"),
                new SuccessRecord(2, BigDecimal.valueOf(2.22), "行2")
        );
        final SXSSFWorkbook workBook = exporter.createWorkBook(records);
        Assertions.assertNotNull(workBook.getSheet(sheetName));
    }

    @ExcelSheet(sheetName)
    static class SuccessRecord {
        @ExcelColumn(order = 1)
        private int column1;
        @ExcelColumn(order = 3, value = "小数列")
        private BigDecimal column2;
        @ExcelColumn(order = 2, value = "字符串列")
        private String column3;

        public SuccessRecord(int column1, BigDecimal column2, String column3) {
            this.column1 = column1;
            this.column2 = column2;
            this.column3 = column3;
        }
    }
}
