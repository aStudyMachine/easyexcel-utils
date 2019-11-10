package com.luwei.module.easyexcel.listener.impl;

import com.luwei.module.easyexcel.listener.BaseExcelListener;
import com.luwei.module.easyexcel.pojo.Order;
import lombok.extern.slf4j.Slf4j;

/**
 * @author WuKun
 * @since 2019/10/10
 */
@Slf4j
public class OrderListener extends BaseExcelListener<Order> {

    @Override
    public boolean validateBeforeAddData(Order object) {
        return true;
    }

    @Override
    public void doService() {
        log.info("/*------- 写入数据 -------*/");
    }
}
