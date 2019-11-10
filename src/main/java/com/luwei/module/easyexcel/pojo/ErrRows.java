package com.luwei.module.easyexcel.pojo;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.util.List;

/**
 * @author luwei
 * @since 2019/10/26
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class ErrRows implements Serializable {

    private static final long serialVersionUID = 1L;

    /**
     * sheetNum(从0开始)
     */
    private String sheetNo;

    /**
     * 读取发生错误的行号(不包括表头)
     */
    private List<Integer> errRows;
}
