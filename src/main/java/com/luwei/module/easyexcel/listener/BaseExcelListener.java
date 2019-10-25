package com.luwei.module.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.util.*;

/**
 * @author WuKun
 * @since 2019-10-10
 * <p>
 * 由于在实际中可能会根据不同的业务场景需要的读取到的不同的excel表的数据进行不同操作,
 * 所以这里将ExcelListener作为所有listener的基类,根据读取不同的java模型自定义一个listener类继承ExcelListener,
 * 根据不同的业务场景选择性对以下方法进行重写,具体如com.luwei.listener.OrderListener所示
 * </p>
 *
 * <p>如果默认实现的方法不满足业务,则直接自定义一个listener实现AnalysisEventListener,重写一遍方法即可.</p>
 */
@Slf4j
public abstract class BaseExcelListener<Model> extends AnalysisEventListener<Model> {

    /**
     * 每隔N条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 3000;

    /**
     * 自定义用于暂时存储data。
     * 可以通过实例获取该值
     * 可以指定AnalysisEventListener的泛型来确定List的存储类型
     */
    @Getter
    private List<Model> data = new ArrayList<>();

    /**
     * 读取时抛出异常是否继续读取,默认true,表示跳过错误行继续读取
     */
    @Setter
    private boolean continueAfterThrowing = true;


    /**
     * 读取过程中发生异常被跳过的行数记录
     * String 为 sheetNo
     * List<Integer> 为 错误的行数列表
     */
    @Getter
    private Map<String, List<Integer>> errRowsMap = new HashMap<>();

    /**
     * 每解析一行会回调invoke()方法。
     * 如果当前行无数据,该方法不会执行,
     * 也就是说如果导入的的excel表无数据,该方法不会执行,
     * 不需要对上传的Excel表进行数据非空判断
     *
     * @param object  当前读取到的行数据对应的java模型对象
     * @param context 定义了获取读取excel相关属性的方法
     */
    @Override
    public void invoke(Model object, AnalysisContext context) {
        log.info("解析到一条数据:{}", object);
        // 数据存储到list，供批量处理，或后续自己业务逻辑处理。
        data.add(object);

        //如果continueAfterThrowing 为false 时保证数据插入的一致性
        if (data.size() >= BATCH_COUNT && continueAfterThrowing) {
            saveData();
            data.clear();
        }
    }

    /**
     * 入库
     */
    abstract void saveData();
//    {
//        log.info("模拟写入数据库");
//        log.info("/*------- {} -------*/", JSON.toJSONString(data));
//        data.clear();
//    }

    /**
     * 解析监听器
     * 每个sheet解析结束会执行该方法
     *
     * @param context 定义了获取读取excel相关属性的方法
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        saveData();
        log.info("/*------- 当前sheet读取完毕,sheetNo : {} , 读取错误的行号列表 : {} -------*/",
                getCurrentSheetNo(context), JSON.toJSONString(errRowsMap));
        data.clear();//解析结束销毁不用的资源
    }

    /**
     * 在转换异常 获取其他异常下会调用本接口。抛出异常则停止读取。如果这里不抛出异常则 继续读取下一行。
     * 如果不重写该方法,默认抛出异常,停止读取
     *
     * @param exception exception
     * @param context   context
     */
    @Override
    public void onException(Exception exception, AnalysisContext context) {
        if (!continueAfterThrowing) {
            throw new IllegalArgumentException(exception);
        }

        Integer sheetNo = getCurrentSheetNo(context);
        Integer rowIndex = context.readRowHolder().getRowIndex();
        log.error("/*------- 读取发生错误! 错误SheetNo:{},错误行号:{} -------*/ ", sheetNo, rowIndex, exception);

        List<Integer> errRowNumList = errRowsMap.get(String.valueOf(sheetNo));
        if (Objects.isNull(errRowNumList)) {
            errRowNumList = new ArrayList<>();
            errRowNumList.add(rowIndex);
            errRowsMap.put(String.valueOf(sheetNo), errRowNumList);
        } else {
            errRowNumList.add(rowIndex);
        }
    }

    /**
     * 获取当前读取的sheet no
     *
     * @param context 定义了获取读取excel相关属性的方法
     * @return current sheet no
     */
    private Integer getCurrentSheetNo(AnalysisContext context) {
        return context.readSheetHolder().getSheetNo();
    }

}