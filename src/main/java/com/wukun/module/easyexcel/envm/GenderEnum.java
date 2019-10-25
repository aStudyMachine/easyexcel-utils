package com.wukun.module.easyexcel.envm;

/**
 * @author WuKun
 * @since 2019/10/14
 */
public enum GenderEnum {
    MAN(0, "男"),
    WOMAN(1, "女");

    private Integer value;
    private String displayName;

    GenderEnum(Integer value, String displayName) {
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
