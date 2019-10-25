package com.luwei.module.easyexcel.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import com.luwei.module.easyexcel.anno.EnumFormat;
import com.luwei.module.easyexcel.anno.LocalDateTimeFormat;
import com.luwei.module.easyexcel.converter.EnumExcelConverter;
import com.luwei.module.easyexcel.converter.LocalDateTimeExcelConverter;
import com.luwei.module.easyexcel.envm.OrderStatusEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

import java.math.BigDecimal;
import java.time.LocalDateTime;

/**
 * @author WuKun
 * @since 2019/10/10
 */
@Data
@AllArgsConstructor
@Accessors(chain = true)
@NoArgsConstructor //必须要保证无参构造方法存在,否则会报初始化对象失败
public class Order {

    @ExcelProperty("订单id")
    private Integer orderId;

    @ExcelProperty("商品名称")
    private String goodsName;

    @ExcelProperty("数量")
    private Integer num;

    @ExcelProperty("总价")
    private BigDecimal price;

    /**
     * <pre>
     * {@code @EnumFormat} 注解 :
     *  作用 : 用于自定义excel单元格中的内容,转换成对应的枚举值
     *  属性 :
     *      value : 要转换的枚举类型
     *      fromExcel : 指定excel中用户输入的枚举值,可以与toJavaEnum中指定的枚举值一一对应
     *                  例如 : excel 单元格中输入
     *                         '待支付' -> OrderStatusEnum.UNPAY
     *                         '已支付' -> OrderStatusEnum.PAYED
     *      toJavaEnum : 如上所述
     * </pre>
     */
    @ExcelProperty(value = "订单状态", converter = EnumExcelConverter.class)
    @EnumFormat(value = OrderStatusEnum.class,
            fromExcel = {"待支付", "已支付", "确认收货"},
            toJavaEnum = {"UNPAY", "PAYED", "RECEIVED"})
    private OrderStatusEnum orderStatus;

    @ExcelProperty(value = "下单时间", converter = LocalDateTimeExcelConverter.class)
    @LocalDateTimeFormat("yyyy-MM-dd HH:mm:ss")
    private LocalDateTime createTime;

//    @ExcelProperty(value = "下单时间")
//    private String createTime;

}
