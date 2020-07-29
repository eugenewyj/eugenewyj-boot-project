package org.eugenewyj.boot.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * ExcelUtil
 *
 * @author eugene
 * @date 2020/7/24
 */
public final class ExcelUtil {
    static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 私有构造函数，防止其他类继承和实例化此工具类。
     */
    private ExcelUtil() {
    }

    /**
     * 将数据写出到输出流中。
     * @param out
     * @param records
     */
    public static void export(OutputStream out, List records) throws IOException, IllegalAccessException {
        logger.info("导出数据到Excel开始，记录数={}", records.size());
        if (Objects.isNull(records) || records.isEmpty()) {
            return;
        }
        Class clazz = records.get(0).getClass();
        Field[] fields = clazz.getDeclaredFields();
        List<ExportField> exportFields = Arrays.stream(fields)
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .map(field -> new ExportField(field))
                .sorted(Comparator.comparing(ExportField::getOrder))
                .collect(Collectors.toList());
        String sheetName = ExcelSheet.DEFAULT_VALUE;
        boolean enableColumnTitle = ExcelSheet.DEFAULT_ENABLE_COLUMN_TITLE;
        if (clazz.isAnnotationPresent(ExcelSheet.class)) {
            ExcelSheet sheetAnnotation = (ExcelSheet) clazz.getAnnotation(ExcelSheet.class);
            sheetName = sheetAnnotation.value();
            enableColumnTitle = sheetAnnotation.enableColumnTitle();
        }
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        SXSSFSheet sheet = sxssfWorkbook.createSheet(sheetName);
        int rowNum = 0;
        if (enableColumnTitle) {
            Row row = sheet.createRow(rowNum++);
            int i = 0;
            for (ExportField exportField : exportFields) {
                Cell cell = row.createCell(i++);
                cell.setCellValue(exportField.getColumnTitle());
            }
        }
        for (Object record : records) {
            Row row = sheet.createRow(rowNum++);
            int i = 0;
            for (ExportField exportField : exportFields) {
                Cell cell = row.createCell(i++);
                cell.setCellValue(exportField.getField().get(record).toString());
            }
        }
        sxssfWorkbook.write(out);
        sxssfWorkbook.dispose();
        logger.info("导出数据到Excel结束");
    }

    static class ExportField {
        private ExcelColumn annotation;
        private Field field;

        /**
         * 构造函数
         * @param field
         */
        public ExportField(Field field) {
            this.field = field;
            this.field.setAccessible(true);
            this.annotation = field.getDeclaredAnnotation(ExcelColumn.class);
        }

        public int getOrder() {
            return annotation.order();
        }

        /**
         * 如果filed注解未指定列名，则采用field名字。
         * @return
         */
        public String getColumnTitle() {
            return "".equals(annotation.value()) ? field.getName() : annotation.value();
        }

        public Field getField() {
            return field;
        }
    }
}
