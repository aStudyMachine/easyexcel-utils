package com.wukun.module.easyexcel.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface LocalDateTimeFormat {

    /**
     * 指定LocalDateTime转化的日期格式
     * 例如 : "yyyy-MM-dd HH:mm:ss",导入导出都会将LocalDateTime转化为对应格式字符串
     *
     * @return 日期格式
     */
    String value();
}
