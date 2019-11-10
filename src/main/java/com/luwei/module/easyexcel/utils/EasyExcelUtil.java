package com.luwei.module.easyexcel.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.luwei.module.easyexcel.listener.BaseExcelListener;
import com.luwei.module.easyexcel.pojo.ErrRows;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.util.Assert;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Optional;

/**
 * EasyExcel工具类
 *
 * @author luwei
 * @since 2019/10/15
 */
@Slf4j
public class EasyExcelUtil {

    /*---------------------------------------------- 导出excel相关 ----------------------------------------------*/

    /**
     * 下载EXCEL文件2007版本
     *
     * @param excelParams 使用EasyExcel导出文件需要设置的相关参数对象,下同
     */
    public static void exportExcel2007(EasyExcelParams excelParams) throws IOException {
        exportExcel(excelParams, ExcelTypeEnum.XLSX);
    }

    /**
     * 下载EXCEL文件2003版本
     *
     * @param excelParams 使用EasyExcel导出文件需要设置的相关参数对象
     */
    public static void exportExcel2003(EasyExcelParams excelParams) throws IOException {
        exportExcel(excelParams, ExcelTypeEnum.XLS);
    }

    /**
     * 根据参数和版本枚举导出excel文件
     *
     * @param excelParams 参数实体
     * @param excelType   excel类型枚举 03 or 07
     * @throws IOException IOException
     */
    private static void exportExcel(EasyExcelParams excelParams, ExcelTypeEnum excelType) throws IOException {
        HttpServletResponse response = excelParams.getResponse();

        ServletOutputStream out = response.getOutputStream();
        prepareResponds(response, excelParams.getExcelNameWithoutExt(), excelType);
        ExcelWriter writer = null;
        try {
            writer = EasyExcel.write(out, excelParams.getDataModelClazz()).excelType(excelType).build();
            WriteSheet writeSheet = EasyExcel.writerSheet(excelParams.getSheetName()).build();
            writer.write(excelParams.getData(), writeSheet);
        } finally {
            //必须保证写出结束后关闭IO
            Optional.ofNullable(writer).ifPresent(ExcelWriter::finish);
        }

    }

    /**
     * 设置response相关参数
     *
     * @param response response
     * @param fileName 文件名
     * @param typeEnum excel类型
     * @throws UnsupportedEncodingException e
     */
    private static void prepareResponds(HttpServletResponse response, String fileName, ExcelTypeEnum typeEnum) throws UnsupportedEncodingException {
        //contentType默认是.xls类型
        response.setContentType("application/vnd.ms-excel;charset=utf-8");

        if (".xlsx".equals(typeEnum.getValue())) {
            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
        }

        response.setHeader("Content-disposition", String.format("attachment; filename=%s", URLEncoder.encode(fileName, "UTF-8") + typeEnum.getValue()));
    }

    /*---------------------------------------------- 以下为读取相关 ----------------------------------------------*/

    /**
     * 读取 Excel(支持单个model的多个sheet)
     *
     * @param excel    文件
     * @param rowModel 实体类映射
     * @param listener 用于读取excel的listener
     * @return 错误的行号
     */
    @SuppressWarnings("unchecked")
    public static List<ErrRows> readExcel(MultipartFile excel, Class rowModel, BaseExcelListener listener) {
        ExcelReader reader = getReader(excel, rowModel, listener);
        try {
            Assert.notNull(reader, "导入Excel失败!");
            Integer totalSheetCount = reader.getSheets().size();
            for (Integer i = 0; i < totalSheetCount; i++) {
                reader.read(EasyExcel.readSheet(i).build());
            }
        } finally {
            // 这里千万别忘记关闭,读的时候会创建临时文件,到时磁盘会崩的
            Optional.ofNullable(reader).ifPresent(ExcelReader::finish);
        }
        return listener.getErrRowsList();
    }

    /**
     * 获取 ExcelReader
     *
     * @param excel    需要解析的 Excel 文件
     * @param listener new ExcelListener()
     * @return ExcelReader 或 null
     */
    private static ExcelReader getReader(MultipartFile excel, Class rowModel,
                                         BaseExcelListener listener) {
        String filename = excel.getOriginalFilename();
        if (StringUtils.isEmpty(filename) || (!filename.toLowerCase().endsWith(".xls") && !filename.toLowerCase().endsWith(".xlsx"))) {
            throw new IllegalArgumentException("文件格式错误！");
        }

        try {
            return EasyExcel.read(excel.getInputStream(), rowModel, listener).build();
        } catch (IOException e) {
            log.error("/*------- 读取Excel IO异常 -------*/", e);
        }
        return null;
    }


}
