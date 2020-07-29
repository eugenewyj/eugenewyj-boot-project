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
     * @param recordClass
     * @throws IOException
     * @throws IllegalAccessException
     */
    public static void export(OutputStream out, List records, Class recordClass) throws IOException, IllegalAccessException {
        logger.info("导出数据到Excel开始，记录数={}", records.size());
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        SXSSFSheet sheet = createSheet(recordClass, sxssfWorkbook);
        List<ExportField> exportFields = getExportFields(recordClass);
        int rowNum = createTitleRow(recordClass, sheet, exportFields);
        exportDataRow(records, sheet, exportFields, rowNum);
        sxssfWorkbook.write(out);
        sxssfWorkbook.dispose();
        logger.info("导出数据到Excel结束");
    }

    /**
     * 导出数据行。
     * @param records
     * @param sheet
     * @param exportFields
     * @param startRowNum
     * @throws IllegalAccessException
     */
    private static void exportDataRow(List records, SXSSFSheet sheet, List<ExportField> exportFields, int startRowNum) throws IllegalAccessException {
        for (Object record : records) {
            Row row = sheet.createRow(startRowNum++);
            int i = 0;
            for (ExportField exportField : exportFields) {
                Cell cell = row.createCell(i++);
                cell.setCellValue(exportField.getField().get(record).toString());
            }
        }
    }

    /**
     * 创建列标题行
     * @param recordClass
     * @param sheet
     * @param exportFields
     * @return
     */
    private static int createTitleRow(Class recordClass, SXSSFSheet sheet, List<ExportField> exportFields) {
        int rowNum = 0;
        boolean enableColumnTitle = ExcelSheet.DEFAULT_ENABLE_COLUMN_TITLE;
        if (recordClass.isAnnotationPresent(ExcelSheet.class)) {
            enableColumnTitle = ((ExcelSheet) recordClass.getAnnotation(ExcelSheet.class)).enableColumnTitle();
        }
        if (enableColumnTitle) {
            Row row = sheet.createRow(rowNum++);
            int i = 0;
            for (ExportField exportField : exportFields) {
                Cell cell = row.createCell(i++);
                cell.setCellValue(exportField.getColumnTitle());
            }
        }
        return rowNum;
    }

    /**
     * 创建sheet。
     * @param recordClass
     * @param sxssfWorkbook
     * @return
     */
    private static SXSSFSheet createSheet(Class recordClass, SXSSFWorkbook sxssfWorkbook) {
        String sheetName = ExcelSheet.DEFAULT_VALUE;
        if (recordClass.isAnnotationPresent(ExcelSheet.class)) {
            sheetName = ((ExcelSheet) recordClass.getAnnotation(ExcelSheet.class)).value();
        }
        return sxssfWorkbook.createSheet(sheetName);
    }

    /**
     * 根据类上的注解获取导出的列及顺序。
     * @param clazz
     * @return
     */
    private static List<ExportField> getExportFields(Class clazz) {
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields)
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .map(field -> new ExportField(field))
                .sorted(Comparator.comparing(ExportField::getOrder))
                .collect(Collectors.toList());
    }

    /**
     * 导出字段的信息
     */
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
         * 对应的列头。
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
