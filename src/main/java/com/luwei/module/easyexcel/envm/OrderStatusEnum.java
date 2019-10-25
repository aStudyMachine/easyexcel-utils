package com.luwei.module.easyexcel.envm;

/**
 * @author WuKun
 * @since 2019/10/10
 */
public enum OrderStatusEnum {

    UNPAY(0, "待支付"),
    PAYED(1, "已支付"),
    RECEIVED(2, "已确认收货");

    private Integer value;
    private String displayName;

    OrderStatusEnum(Integer value, String displayName) {
        this.value = value;
        this.displayName = displayName;
    }

    public Integer getValue() {
        return value;
    }

    public String getDisplayName() {
        return displayName;
    }
}
