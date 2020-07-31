package org.eugenewyj.boot.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.eugenewyj.boot.poi.annotation.ExcelColumn;
import org.eugenewyj.boot.poi.annotation.ExcelSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * SimpleExcelExporter
 *
 * @author eugene
 * @date 2020/7/30
 */
public final class SimpleExcelExporter {
    private static final Logger logger = LoggerFactory.getLogger(SimpleExcelExporter.class);

    private Class recordClass;
    private ExcelSheet excelSheet;
    private List<SimpleExcelExporter.ExportField> exportFields;

    /**
     * 根据导出支持的记录类，构造导出器。
     * @param recordClass 支持的导出记录类型。
     */
    public SimpleExcelExporter(Class recordClass) {
        this.recordClass = recordClass;
        if (!recordClass.isAnnotationPresent(ExcelSheet.class)) {
            throw new IllegalArgumentException("导出Excel的记录类上必须有ExcelSheet注解。");
        }
        excelSheet = ((ExcelSheet) recordClass.getAnnotation(ExcelSheet.class));
        this.exportFields = Arrays.stream(recordClass.getDeclaredFields())
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .map(field -> new SimpleExcelExporter.ExportField(field))
                .sorted(Comparator.comparing(SimpleExcelExporter.ExportField::getOrder))
                .collect(Collectors.toList());
        if (this.exportFields.isEmpty()) {
            throw new IllegalArgumentException("导出Excel的记录类上未通过注解ExcelColumn指定导出列。");
        }
    }

    /**
     * 根据注解导出记录集合到Excel
     * @param out
     * @param records
     */
    public void export(OutputStream out, List records) throws IOException, IllegalAccessException {
        records = Optional.ofNullable(records).orElse(new ArrayList());
        logger.info("导出数据到Excel开始，记录数={}", records.size());
        checkRecordClass(records);
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook();
        SXSSFSheet sheet = sxssfWorkbook.createSheet(excelSheet.value());
        int rowNum = createTitleRow(sheet);
        exportData(records, sheet, rowNum);
        sxssfWorkbook.write(out);
        sxssfWorkbook.dispose();
        logger.info("导出数据到Excel结束");
    }

    /**
     * 校验集合中记录类是否符合要求。
     * @param records
     */
    private void checkRecordClass(List records) {
        if (records.isEmpty()) {
            return;
        }
        Class clazz = records.get(0).getClass();
        if (!recordClass.isAssignableFrom(clazz)) {
            throw new IllegalArgumentException(String.format("参数records中的记录类型并非继承自%s", recordClass.toString()));
        }
    }

    /**
     * 导出数据行。
     * @param records
     * @param sheet
     * @param startRowNum
     * @throws IllegalAccessException
     */
    private void exportData(List records, SXSSFSheet sheet, int startRowNum) throws IllegalAccessException {
        if (records.isEmpty()) {
            return;
        }
        CellStyle defaultCellStyle = defaultRecordStyle(sheet.getWorkbook());
        CellStyle numberCellStyle = sheet.getWorkbook().createCellStyle();
        numberCellStyle.cloneStyleFrom(defaultCellStyle);
        numberCellStyle.setAlignment(HorizontalAlignment.RIGHT);
        for (SimpleExcelExporter.ExportField exportField : exportFields) {
            switch (exportField.getExportStyleType()) {
                case EXPORT_NUMBER:
                    exportField.setCellStyle(numberCellStyle);
                    break;
                case EXPORT_TEXT:
                    exportField.setCellStyle(defaultCellStyle);
                    break;
            }
        }
        for (Object record : records) {
            Row row = sheet.createRow(startRowNum++);
            int i = 0;
            for (SimpleExcelExporter.ExportField exportField : exportFields) {
                Cell cell = row.createCell(i++);
                Object value = exportField.getField().get(record);
                cell.setCellValue(value.toString());
                cell.setCellStyle(Optional.ofNullable(exportField.getCellStyle()).orElse(defaultCellStyle));
            }
        }
    }

    /**
     * 创建列标题行
     * @param sheet
     * @return
     */
    private int createTitleRow(SXSSFSheet sheet) {
        int rowNum = 0;
        if (excelSheet.enableColumnTitle()) {
            Row row = sheet.createRow(rowNum++);
            int i = 0;
            CellStyle cellStyle = titleStyle(sheet.getWorkbook());
            for (SimpleExcelExporter.ExportField exportField : exportFields) {
                Cell cell = row.createCell(i++);
                cell.setCellValue(exportField.getColumnTitle());
                cell.setCellStyle(cellStyle);
            }
        }
        return rowNum;
    }

    /**
     * 默认标题单元格style。
     * @param workBook
     * @return
     */
    private CellStyle titleStyle(SXSSFWorkbook workBook) {
        CellStyle style = workBook.createCellStyle();
        style.setFont(createFont(workBook));
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /**
     * 默认数据单元格Style.
     * @param workBook
     * @return
     */
    private CellStyle defaultRecordStyle(SXSSFWorkbook workBook) {
        CellStyle style = workBook.createCellStyle();
        style.setFont(createFont(workBook));
        return style;
    }

    /**
     * 创建字体
     * @param workBook
     * @return
     */
    private Font createFont(SXSSFWorkbook workBook) {
        Font font = workBook.createFont();
        font.setFontName(excelSheet.fontName());
        font.setFontHeightInPoints(excelSheet.fontSize());
        return font;
    }

    /**
     * 导出字段的信息
     */
    static class ExportField {
        private ExcelColumn annotation;
        private Field field;
        private CellStyle cellStyle;

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

        public ExportStyleType getExportStyleType() {
            Class type = this.field.getType();
            List primitiveNumberTypes = Arrays.asList("int", "long", "short", "byte", "float", "double");
            boolean isPrimitiveNumber = type.isPrimitive() && primitiveNumberTypes.contains(type.getName());
            if (isPrimitiveNumber || Number.class.isAssignableFrom(type)) {
                return ExportStyleType.EXPORT_NUMBER;
            } else {
                return ExportStyleType.EXPORT_TEXT;
            }
        }

        public void setCellStyle(CellStyle cellStyle) {
            this.cellStyle = cellStyle;
        }

        public CellStyle getCellStyle() {
            return cellStyle;
        }
    }

    /**
     * 导出风格枚举
     */
    enum ExportStyleType {
        /**
         * 数值
         */
        EXPORT_NUMBER,
        /**
         * 文本
         */
        EXPORT_TEXT,
    }
}
