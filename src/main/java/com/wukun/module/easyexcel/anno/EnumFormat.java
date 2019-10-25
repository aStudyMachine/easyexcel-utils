package com.wukun.module.easyexcel.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <pre>
 * '@EnumFormat 注解 :
 *  作用 : 用于自定义excel单元格中的内容,转换成对应的枚举值
 *  属性 :
 *      value : 要转换的枚举类型
 *      fromExcel : 指定excel中用户输入的枚举值,可以与toJavaEnum中指定的枚举值一一对应
 *                  例如 : excel 单元格中输入
 *                         '待支付' -> OrderStatusEnum.UNPAY
 *                         '已支付' -> OrderStatusEnum.PAYED
 *      toJavaEnum : 如上所述
 *  注意 : toJavaEnum 与 fromExcel 必须搭配使用
 * </pre>
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface EnumFormat {
    /**
     * 要转换的枚举类型
     *
     * @return enum class
     */
    Class value();

    /**
     * 要转换枚举的全部变量名数组集
     *
     * @return String[]
     */
    String[] toJavaEnum() default {};

    /**
     * 枚举导出excel时所展示的内容
     *
     * @return String[]
     */
    String[] fromExcel() default {};
}
