import lombok.extern.slf4j.Slf4j;

/**
 * @author WuKun
 * @since 2019/10/10
 * <p>
 * 读excel demo
 */
@Slf4j
public class Test01 {

    private static final String pathName1 = "C:/Users/studymachine/Desktop/user.xlsx";

//    @Test
//    public void test01() {
//        // 写法1：
//        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
//        EasyExcel.read(pathName1, User.class, new UserListener()).sheet().doRead();
//
//    }
//
//    @Test
//    public void test02() {
//        // 写法2：
//        ExcelReader excelReader = null;
//        try {
//            excelReader = EasyExcel.read(pathName1, User.class, new UserListener()).build();
//
//            ReadSheet readSheet = EasyExcel.readSheet(0)
//                    .head(User.class)
//                    .registerReadListener(new UserListener())
//                    .build();
//
//            String sheetName = excelReader.excelExecutor().sheetList().get(0).getSheetName();
//            log.info("/*-------  sheetName : {}-------*/", sheetName);
//
//            excelReader.read(readSheet);
//
//        } finally {
//            // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
//            Optional.ofNullable(excelReader).ifPresent(ExcelReader::finish);
//        }
//    }
//
//    /**
//     * sheet1 sheet2 数据不一致 写法
//     */
//    @Test
//    public void test03() {
//        ExcelReader excelReader = null;
//        try {
//            excelReader = EasyExcel.read(pathName1).build();
//
//            ReadSheet readSheet1 = EasyExcel.readSheet(0)
//                    .head(User.class)
//                    .registerReadListener(new UserListener())
//                    .build();
//            excelReader.read(readSheet1);
//
////            String sheetName = excelReader.excelExecutor().sheetList().get(0).getSheetName();
////            log.info("/*-------  sheetName : {} -------*/", sheetName);
//
//            ReadSheet readSheet2 = EasyExcel.readSheet(1)
//                    .head(Order.class)
//                    .registerReadListener(new OrderListener())
//                    .build();
//            excelReader.read(readSheet2);
//
//        } finally {
//            // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
//            Optional.ofNullable(excelReader).ifPresent(ExcelReader::finish);
//        }
//
//    }
//
//
//    @Test
//    public void test04() {
//        String path = "C:\\Users\\studymachine\\Desktop\\Order.xlsx";
//
//        ExcelReader excelReader = null;
//        try {
//            excelReader = EasyExcel.read(path).build();
//
//            ReadSheet readSheet1 = EasyExcel.readSheet(0)
//                    .head(Order.class)
//                    .registerReadListener(new OrderListener())
//                    .build();
//            excelReader.read(readSheet1);
//
//        } finally {
//            // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
//            Optional.ofNullable(excelReader).ifPresent(ExcelReader::finish);
//        }
//
//    }
}
