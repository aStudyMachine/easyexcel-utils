package com.wukun.module.easyexcel.listener;

import com.wukun.module.easyexcel.pojo.Order;
import lombok.extern.slf4j.Slf4j;

/**
 * @author WuKun
 * @since 2019/10/10
 */
@Slf4j
public class OrderListener extends BaseExcelListener<Order> {


    @Override
    void saveData() {
        log.info("/*------- 写入数据 -------*/");
    }
}
