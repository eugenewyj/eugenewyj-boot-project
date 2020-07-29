package org.eugenewyj.boot.poi;

import java.lang.annotation.*;

/**
 * ExcelSheet
 *
 * @author eugene
 * @date 2020/7/28
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelSheet {
     String DEFAULT_VALUE = "Sheet1";
     boolean DEFAULT_ENABLE_COLUMN_TITLE = true;

    /**
     * 对应的Excel Sheet名称，默认是Sheet1。
     * @return
     */
    String value() default DEFAULT_VALUE;

    /**
     * 是否有列头。
     * @return
     */
    boolean enableColumnTitle() default DEFAULT_ENABLE_COLUMN_TITLE;
}
