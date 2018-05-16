package com.alibaba.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by jipengfei on 17/3/19.
 * Field column num at excel head
 *
 * @author jipengfei
 *
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelColumnNum {

    /**
     * col num
     * @return col num
     */
    int value();

    /**
     *
     * Default @see com.alibaba.TypeUtil
     * if default is not  meet you can set format
     *
     * @return format
     */
    String format() default "";

    /**
     * 枚举
     *
     * 例如：男_1,女_2
     *
     * @return
     */
    String replace() default "";
}
