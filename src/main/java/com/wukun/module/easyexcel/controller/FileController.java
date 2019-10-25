package com.wukun.module.easyexcel.controller;

import com.wukun.module.easyexcel.envm.OrderStatusEnum;
import com.wukun.module.easyexcel.listener.UserListener;
import com.wukun.module.easyexcel.pojo.Order;
import com.wukun.module.easyexcel.pojo.User;
import com.wukun.module.easyexcel.service.UserService;
import com.wukun.module.easyexcel.utils.EasyExcelParams;
import com.wukun.module.easyexcel.utils.EasyExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author: WuKun
 * @Date: 2019/4/17 14:34
 * <p>
 * 演示demo
 */
@Slf4j
@RestController
public class FileController {
    /*---------------------------------------------- Fields ~ ----------------------------------------------*/
    //用于测试EasyExcelUtil
    private List<Order> data = new ArrayList<>();

    /*---------------------------------------------- Methods ~ ----------------------------------------------*/

    /**
     * 使用EasyExcelUtil导出Excel 2003
     *
     * @param response
     * @throws Exception
     */
    @GetMapping("/easy2003")
    public void easy2003(HttpServletResponse response) throws Exception {
        initData();

        //设置参数
        EasyExcelParams params = new EasyExcelParams().setResponse(response)
                .setExcelNameWithoutExt("Order(xlsx)")
                .setSheetName("第一张sheet")
                .setData(data)
                .setDataModelClazz(Order.class);

        long begin = System.currentTimeMillis();
        EasyExcelUtil.exportExcel2003(params);
        long end = System.currentTimeMillis();

        log.info("-----EasyExcelUtils : 导出成功,导出excel花费时间为 : " + ((end - begin) / 1000) + "秒");
    }

    /**
     * 使用EasyExcelUtils 导出Excel 2007
     *
     * @param response HttpServletResponse
     * @throws Exception exception
     */
    @GetMapping("/easy2007")
    public void easy2007(HttpServletResponse response) throws Exception {
        initData();

        //设置参数
        EasyExcelParams params = new EasyExcelParams().setResponse(response)
                .setExcelNameWithoutExt("Order(xlsx)")
                .setSheetName("第一张sheet")
                .setData(data)
                .setDataModelClazz(Order.class)
                .checkValid();

        long begin = System.currentTimeMillis();
        EasyExcelUtil.exportExcel2007(params);
        long end = System.currentTimeMillis();

        log.info("-----EasyExcelUtils : 导出成功,导出excel花费时间为 : " + ((end - begin) / 1000) + "秒");
    }

    private void initData() {
        if (CollectionUtils.isEmpty(data)) {
            for (int i = 0; i < 60000; i++) {
                data.add(new Order().setPrice(BigDecimal.valueOf(11.11))
                        .setCreateTime(LocalDateTime.now()).setGoodsName("香蕉")
                        .setOrderId(i)
                        .setNum(11)
                        .setOrderStatus(OrderStatusEnum.PAYED));
            }
        }
    }


    /**
     * 导出User相关模板
     */
    @GetMapping("/easy2007/template/user")
    public void easy2007Template4User(HttpServletResponse response) throws IOException {
        //设置参数
        EasyExcelParams params = new EasyExcelParams().setResponse(response)
                .setExcelNameWithoutExt("Order(xlsx)")
                .setSheetName("第一张sheet")
                .setData(new ArrayList<User>())
                .setDataModelClazz(User.class)
                .checkValid();

        long begin = System.currentTimeMillis();
        EasyExcelUtil.exportExcel2007(params);
        long end = System.currentTimeMillis();

        log.info("-----EasyExcelUtils : 导出成功,导出excel花费时间为 : " + ((end - begin) / 1000) + "秒");
    }

    @Autowired
    private UserService userService;

    /**
     * 读取测试
     *
     * @param excel excel文件
     */
    @PostMapping("/readExcel")
    public void readExcel(@RequestParam MultipartFile excel) {
        EasyExcelUtil.readExcel(excel, User.class, new UserListener(userService));
    }

}
