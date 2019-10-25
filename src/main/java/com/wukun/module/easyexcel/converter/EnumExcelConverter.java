package com.wukun.module.easyexcel.converter;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.wukun.module.easyexcel.anno.EnumFormat;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.EnumUtils;
import org.springframework.util.Assert;

import java.util.Objects;

/**
 * @author WuKun
 * @since 2019/10/10
 */
public class EnumExcelConverter implements Converter<Enum> {

    @Override
    public Class supportJavaTypeKey() {
        return Enum.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public Enum convertToJavaData(CellData cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        String cellDataStr = cellData.getStringValue();

        EnumFormat annotation = contentProperty.getField().getAnnotation(EnumFormat.class);
        Class enumClazz = annotation.value();
        String[] fromExcel = annotation.fromExcel();
        String[] toJavaEnum = annotation.toJavaEnum();

        Enum anEnum = null;
        if (ArrayUtils.isNotEmpty(fromExcel) && ArrayUtils.isNotEmpty(toJavaEnum)) {
            Assert.isTrue(fromExcel.length == toJavaEnum.length, "fromExcel 与 toJavaEnum 的长度必须相同");
            for (int i = 0; i < fromExcel.length; i++) {
                if (Objects.equals(fromExcel[i], cellDataStr)) {
                    anEnum = EnumUtils.getEnum(enumClazz, toJavaEnum[i]);
                }
            }
        } else {
            anEnum = EnumUtils.getEnum(enumClazz, cellDataStr);
        }

        Assert.notNull(anEnum, "枚举值不合法");
        return anEnum;
    }

    @Override
    public CellData convertToExcelData(Enum value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {

        String enumName = value.name();

        EnumFormat annotation = contentProperty.getField().getAnnotation(EnumFormat.class);
        String[] fromExcel = annotation.fromExcel();
        String[] toJavaEnum = annotation.toJavaEnum();

        if (ArrayUtils.isNotEmpty(fromExcel) && ArrayUtils.isNotEmpty(toJavaEnum)) {
            Assert.isTrue(fromExcel.length == toJavaEnum.length, "fromExcel 与 toJavaEnum 的长度必须相同");
            for (int i = 0; i < toJavaEnum.length; i++) {
                if (Objects.equals(toJavaEnum[i], enumName)) {
                    return new CellData(fromExcel[i]);
                }
            }
        }
        return new CellData(enumName);
    }
}
