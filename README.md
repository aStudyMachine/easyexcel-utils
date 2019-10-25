# 1. 前言

最近阿里开源的excel读写项目**EasyExcel**又火了起来 . 原来是项目又开始维护了, 从1.x 更新到2.x 了 , 而且迭代迅速 , 目前已经更新到 [2.1.0-beta3](https://mvnrepository.com/artifact/com.alibaba/easyexcel/2.1.0-beta3) 版本. 

关于easyexcel , 其主要目的为**降低读取excel时的内存损耗** , **简化读写excel的api** .  

同时2.x版本提供了很多新功能 , 具体大家可以直接参考官方说明吧 , github文档上写一清二楚 , 这里就不传播一些没啥必要的二三手知识了 , 而且目前该项目还在不停地迭代 , 给contributor一个star也是很有必要的 . :smile:

官方地址 : [easyexcel仓库地址](https://github.com/alibaba/easyexcel)  , [easyexcel官网](https://alibaba-easyexcel.github.io/index.html)

<br>

这里我基于`easyexcel 2.0.5`版本简单封装了一个web读写excel的工具类 , 主要封装了如下功能 : 

- 通过注解自定义`LocalDateTime`的读写格式
- 通过注解自定义枚举类型的读写格式
- 自定义`BaseExcelListener`抽象类封装了常用的数据处理逻辑 , 以及补充读取excel过程中读取发生错误被跳过的行号记录 . 

- 封装了web的读写excel操作
- ...

下面列举主要功能以及相关示例 , 可以直接看源码 , 每个方法都有写完整的注释 , 如果觉得写得还凑合能看的话 , 给我这个刚毕业没多久的小菜鸡点个star呗 :smile:

附上源码地址 :  https://github.com/aStudyMachine/easyexcel-utils  



# 2. 主要功能

## 2.1 建立excel表每行数据与Java模型的映射

easyexcel读写excel可以基于java 模型的方式 , 也可以使用`List<List<String>>` 的方式读写excel , 这里我读写操作使用**基于java模型**的方式  , 通过java类的属性与excel每一列的数据进行对应 

关键注解 : `@ExcelProperty`

具体如何使用注解建立java模型与Excel表数据的映射可以参考 `com.wukun.module.easyexcel.pojo`下的两个java模型类`Order` 类与`User`类

```java
/**
 * @author WuKun
 * @since 2019/10/09
 */
@Data
@AllArgsConstructor
@NoArgsConstructor //必须要保证无参构造方法存在,否则会报初始化对象失败
public class User {

    /**
     * {@code @ExcelIgnore} 用于标识该字段不用做excel读写过程中的数据转换
     */
    @ExcelIgnore
    private Integer userId;

    /**
     * <pre>
     * {@code @ExcelIgnore} 中的属性 不建议 index 和 name 同时用
     *
     * 要么一个对象统一只用index表示列号,
     * 例如 : {@code @ExcelProperty(index = 0)}
     *
     * 要么一个对象统一只用value去匹配列名
     * 例如 : {@code @ExcelProperty("姓名")}
     *
     * 用名字去匹配,这里需要注意,如果名字重复,会导致只有一个字段读取到数据
     * </pre>
     */
    @ExcelProperty("姓名")
    private String name;

    @ExcelProperty("年龄")
    private Integer age;

    @ExcelProperty("地址")
    private String address;

    /**
     * <pre>
     * {@code @EnumFormat} 注解 :
     *  作用 : 与 {@code @ExcelProperty(converter = EnumExcelConverter.class)} 搭配使用
     *         转换java枚举与excel中指定的内容
     *  属性 :
     *   - value : 要转换的枚举类class对象
     *   - fromExcel : 指定excel中用户输入的枚举值的名字的字符串形式,与toJavaEnum中指定的枚举值一一对应
     *                 以下面的示例来说,fromExcel指定的 "男" 对应 toJavaEnum中的 "MAN" ,
     *                 当excel中该列读取到"男" 这个字符串时,会自动转化为枚举{@code GenderEnum.MAN},
     *                 同理在写excel时,如果该字段为{@code GenderEnum.MAN} 时, 写到excel时则转化为 "男"
     *   - toJavaEnum : 如上所述
     *
     *  注意 : fromExcel 与 toJavaEnum 这两个属性必须同时使用, 而且两个属性的字符串的数组长度必须相同,
     *        若两个属性都不指定 , 则默认 枚举值名字符串转化为对应的枚举 例如: "MAN" <--> {@code GenderEnum.MAN}
     * </pre>
     */
    @EnumFormat(value = GenderEnum.class,
            fromExcel = {"男", "女"},
            toJavaEnum = {"MAN", "WOMAN"}) // "男" <--> GenderEnum.MAN ; "女" <--> GenderEnum.WOMAN
    @ExcelProperty(value = "性别", converter = EnumExcelConverter.class)
    private GenderEnum gender;

    /**
     * <pre>
     * {@code @LocalDateTimeFormat} 注解
     *  作用: 与 {@code  @ExcelProperty(converter = LocalDateTimeExcelConverter.class)} 搭配使用,
     *        指定导入导出的时间格式.
     *  属性 :
     *   - value : 日期格式字符串 
     * </pre>
     */
    @ExcelProperty(value = "生日", converter = LocalDateTimeExcelConverter.class)
    @LocalDateTimeFormat("yyyy-MM-dd HH:mm:ss")
    private LocalDateTime birthday;
}

```



## 2.2 导出excel

### 2.2.1 相关API介绍

web导出excel 根据03 / 07版本分为两个不同的方法 ,分别为`EasyExcelUtil`类中以下两个方法  : 

-   导出03版本 : `exportExcel2003Format(EasyExcelParams excelParams)`

-   导出07版本 : `exportExcel2007Format(EasyExcelParams excelParams)`

`EasyExcelParams`是使用EasyExcel导出excel需要设置的相关参数 , 包括需要导出的`List<T>`数据以及对应的Java模型 , 使用时根据实际情况设置相应的参数即可.

```java
/**
 * @author WuKun
 * @since 2019/10/14
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class EasyExcelParams implements Serializable {


    /**
     * excel文件名（不带拓展名)
     */
    private String excelNameWithoutExt;
    /**
     * sheet名称
     */
    private String sheetName;

    /**
     * 数据
     */
    private List data;

    /**
     * 数据模型类型
     */
    private Class dataModelClazz;

    /**
     * 响应
     */
    private HttpServletResponse response;


    public EasyExcelParams() {
    }

    /**
     * 检查不允许为空的属性
     *
     * @return this
     */
    public EasyExcelParams checkValid() {
        Assert.isTrue(ObjectUtils.allNotNull(excelNameWithoutExt, data, dataModelClazz, response), "导出excel参数不合法!");
        return this;
    }
}
```

### 2.2.2 导出excel示例

```java
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
```



## 2.3 读取excel

### 2.3.1 相关API介绍

-   读取excel时用到的是`EasyExcelUtils`的`readExcel`方法  ;

```java
/**
 * 读取 Excel(支持单个model的多个sheet)
 *
 * @param excel    文件
 * @param rowModel 实体类映射
 * @param listener 用于读取excel的listener
 */
public static void readExcel(MultipartFile excel, Class rowModel, BaseExcelListener listener) {
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
}
```

-   easyexcel的读取操作需要自建一个类继承`AnalysisEventListener`抽象类 , 这里我创建`BaseExcelListener`类继承并重写读取excel的相关方法  , 每个方法的具体作用可直接查看方法头部注释 , 使用时直接创建一个listener类继承`BaseExcelListener`即可 , 如果默认的`BaseExcelListener`不满足需求 , 也可以直接自定义一个Listener 类继承 `BaseExcelListener`并重写相应方法. 

```java
/**
 * @author WuKun
 * @since 2019-10-10
 * <p>
 * 由于在实际中可能会根据不同的业务场景需要的读取到的不同的excel表的数据进行不同操作,
 * 所以这里将{@link BaseExcelListener}作为所有listener的父类,根据读取不同的java模型自定义一个listener类继承{@link BaseExcelListener},
 * 根据不同的业务场景选择性对以下方法进行重写,具体如{@link OrderListener}所示
 * </p>
 *
 * <p>如果默认实现的方法不满足业务,则直接自定义一个listener继承{@link BaseExcelListener},重写一遍方法即可.</p>
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
     * 入库,继承该类后实现该方法即可
     */
    abstract void saveData();

    /**
     * 解析监听器
     * 每个sheet解析结束会执行该方法
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
     * @param context
     * @return
     */
    private Integer getCurrentSheetNo(AnalysisContext context) {
        return context.readSheetHolder().getSheetNo();
    }

}
```
-   读取时不区分03或07版本 , 底层会自动判断 ;



### 2.3.2 读取excel示例

1. 自定义一个listener类继承`BaseExcelListener`

```java
package com.luwei.module.easyexcel.listener;

import com.wukun.module.easyexcel.pojo.User;
import lombok.extern.slf4j.Slf4j;

/**
 * @author WuKun
 * @since 2019/10/10
 */
@Slf4j
public class UserListener extends BaseExcelListener<User> {
    /**
     * 这里需要注意入库使用到的Service或者DAO层需要使用到的相关方法时,
     * 不要通过Spring 使用{@code @Autowired}注入,同时该Listener也不要交由Spring IOC进行管理
     * 直接通过构造方法传入相关`xxxService` 或者 `xxxMapper`
     */
    private UserService userService;

    public UserListener(UserService userService) {
        this.userService = userService;
    }
    
    @Override
    void saveData() {
        // 批量插入数据
        userService.saveBatchUsers(this.getData())
        log.info("/*------- 写入数据 -------*/");
    }
}
```

2. 调用工具方法 

```java
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
```



# 3. 注意事项

- java模型必须要保证无参构造方法存在 , 否则会在读写excel时报无法初始化java模型对象的异常

- ~~使用java模型读取excel时不能对Java模型使用`@Accessors(chain = true)`注解, 会导致数据无法转换~~ (easyexcel 2.x的API该问题已解决)

- sheetNo 从 0开始 , 行号不包括表头 , 例如log中打印的是第9行, 实际在excel中对应的是第10行

  ```powershell
  2019-10-20 15:34:57.236  INFO 38012 --- [nio-8081-exec-8] c.l.e.listener.BaseExcelListener         : /*------- 当前sheet读取完毕,sheetNo : 1 , 读取错误的行号列表 : {"1":[9]} -------*/
  ```

  ![image.png](https://i.loli.net/2019/10/20/UeL7G39JNXmC1lK.png)

