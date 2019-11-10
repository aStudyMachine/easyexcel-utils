import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.luwei.module.easyexcel.envm.OrderStatusEnum;
import com.luwei.module.easyexcel.pojo.Order;
import org.junit.Test;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * @author WuKun
 * @since 2019/10/12
 * <p>
 * 写excel demo
 */
public class Test02 {

    private static final String pathName1 = "C:/Users/studymachine/Desktop/";

    @Test
    public void test01() {
        EasyExcel.write(pathName1 + "/write.xlsx", Order.class)
                .sheet("模板")
                .doWrite(data(100));
    }

    @Test
    public void test02() {
        // 写法2
        // 这里 需要指定写用哪个class去读
        ExcelWriter excelWriter = EasyExcel.write(pathName1 + "write.xlsx", Order.class).build();
        WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
        excelWriter.write(data(1000), writeSheet);
        /// 千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }

    @Test
    public void test03() {
        EasyExcel.write(pathName1 + "write" + ExcelTypeEnum.XLS.getValue(), Order.class)
                .excelType(ExcelTypeEnum.XLS)
                .sheet("sheet1").doWrite(data(1000));

    }

    private List<Order> data(int row) {
        List<Order> list = new ArrayList<>();

        for (int i = 0; i < row; i++) {

            Order order = new Order();
            order.setCreateTime(LocalDateTime.now());
            order.setGoodsName("香蕉");
            order.setNum(1);
            order.setOrderId(i);
            order.setOrderStatus(OrderStatusEnum.PAYED);
            order.setPrice(BigDecimal.valueOf(11.16));
            list.add(order);
        }
        return list;
    }


}
