package org.eugenewyj.boot.poi.annotation;

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
    /**
     * 对应的Excel Sheet名称，默认是Sheet1。
     * @return
     */
    String value() default "Sheet1";

    /**
     * 是否有列头。
     * @return
     */
    boolean enableColumnTitle() default true;

    /**
     * 字体名称
     * @return
     */
    String fontName() default "微软雅黑";

    /**
     * 字体大小
     * @return
     */
    short fontSize() default 9;
}
