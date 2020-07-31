package org.eugenewyj.boot.poi.annotation;

import java.lang.annotation.*;

/**
 * ExcelColumn
 *
 * @author eugene
 * @date 2020/7/24
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {
    /**
     * 在Excel中对应的列序号。
     * @return
     */
    int order() default 0;

    /**
     * Excel中对应的列头。
     * @return
     */
    String value() default "";
}
