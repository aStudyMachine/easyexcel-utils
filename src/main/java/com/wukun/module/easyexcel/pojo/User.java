package com.wukun.module.easyexcel.pojo;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.wukun.module.easyexcel.anno.EnumFormat;
import com.wukun.module.easyexcel.anno.LocalDateTimeFormat;
import com.wukun.module.easyexcel.converter.EnumExcelConverter;
import com.wukun.module.easyexcel.converter.LocalDateTimeExcelConverter;
import com.wukun.module.easyexcel.envm.GenderEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

/**
 * @author WuKun
 * @since 2019/10/09
 */
@Data
@AllArgsConstructor
@NoArgsConstructor //必须要保证无参构造方法存在,否则会报初始化对象失败
public class User {

    /**
     * {@code @ExcelIgnore} 用于标识该字段不用做excel读写过程中的数据转换
     */
    @ExcelIgnore
    private Integer userId;

    /**
     * <pre>
     * {@code @ExcelIgnore} 中的属性 不建议 index 和 name 同时用
     *
     * 要么一个对象统一只用index表示列号,
     * 例如 : {@code @ExcelProperty(index = 0)}
     *
     * 要么一个对象统一只用value去匹配列名
     * 例如 : {@code @ExcelProperty("姓名")}
     *
     * 用名字去匹配,这里需要注意,如果名字重复,会导致只有一个字段读取到数据
     * </pre>
     */
    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("年龄")
    private Integer age;

    @ExcelProperty("地址")
    private String address;

    /**
     * <pre>
     * {@code @EnumFormat} 注解 :
     *  作用 : 与 {@code @ExcelProperty(converter = EnumExcelConverter.class)} 搭配使用
     *         转换java枚举与excel中指定的内容
     *  属性 :
     *   - value : 要转换的枚举类class对象
     *   - fromExcel : 指定excel中用户输入的枚举值的名字的字符串形式,与toJavaEnum中指定的枚举值一一对应
     *                 以下面的示例来说,fromExcel指定的 "男" 对应 toJavaEnum中的 "MAN" ,
     *                 当excel中该列读取到"男" 这个字符串时,会自动转化为枚举{@code GenderEnum.MAN},
     *                 同理在写excel时,如果该字段为{@code GenderEnum.MAN} 时, 写到excel时则转化为 "男"
     *   - toJavaEnum : 如上所述
     *
     *  注意 : fromExcel 与 toJavaEnum 这两个属性必须同时使用, 而且两个属性的字符串的数组长度必须相同,
     *        若两个属性都不指定 , 则默认 枚举值名字符串转化为对应的枚举 例如: "MAN" <--> {@code GenderEnum.MAN}
     * </pre>
     */
    @EnumFormat(value = GenderEnum.class,
            fromExcel = {"男", "女"},
            toJavaEnum = {"MAN", "WOMAN"}) // "男" <--> GenderEnum.MAN ; "女" <--> GenderEnum.WOMAN
    @ExcelProperty(value = "性别", converter = EnumExcelConverter.class)
    private GenderEnum gender;

    /**
     * <pre>
     * {@code @LocalDateTimeFormat} 注解
     *  作用: 与 {@code  @ExcelProperty(converter = LocalDateTimeExcelConverter.class)} 搭配使用,
     *        指定导入导出的时间格式.
     *  属性 :
     *   - value : 日期格式字符串
     * </pre>
     */
    @ExcelProperty(value = "生日", converter = LocalDateTimeExcelConverter.class)
    @LocalDateTimeFormat("yyyy-MM-dd HH:mm:ss")
    private LocalDateTime birthday;
}
