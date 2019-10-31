package com.bing.studyexcel.util;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Description: Excel表格导入导出工具
 * @Author: 杨亚兵
 * @Date: 2019/10/30 10:06
 */
public class ExcelUtil {

    /**
     * 文件类型，xsl格式
     */
    public static final String EXCEL_XLS = "xls";
    /**
     * 文件类型，xslx格式
     */
    public static final String EXCEL_XLSX = "xlsx";
    /**
     * 表头所在行数
     */
    private static final Integer EXCEL_HEAD_ROW_NUM = 1;
    /**
     * 日志打印
     */
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    public static <T> List<T> importExcel(MultipartFile file, Class<T> entityClass)
            throws Exception {
        //检查文件
        checkFile(file);
        //获取工作簿
        Workbook workbook = getWorkbook(file);

        List<T> dataList = new LinkedList<>();

        int sheetCount = workbook.getNumberOfSheets();
        if (sheetCount == 0) {
            throw new IOException("文件中没有任何数据");
        }
        for (int i = 0; i < sheetCount; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            Object[] heads = null;
            for (Row row : sheet) {
                int num = row.getLastCellNum();
                //第一次进入获取表头，以后跳过表头
                if (row.getRowNum() < EXCEL_HEAD_ROW_NUM) {
                    heads = new Object[num];
                    for (int k = 0; k < num; k++) {
                        if (row.getCell(k) == null || "".equals(row.getCell(k).toString())){
                            heads[k] = "";
                        }
                        else
                        {
                            heads[k] = row.getCell(k);
                        }
                    }
                    //校验表头
                    checkExcelHead(heads,entityClass);
                    continue;
                }
                T entity = entityClass.newInstance();
                //获得表格某行的值
                Object[] objects = new Object[num];
                for (int k = 0; k < num; k++) {
                    if (row.getCell(k) == null || "".equals(row.getCell(k).toString())){
                        objects[k] = "";
                    }
                    else{
                        objects[k] = row.getCell(k);
                    }
                }
                if (heads == null){
                    throw new IOException("表头为空");
                }
                //给对象赋值
                setValue(entity, objects, heads);
                //将对象添加至数据列表
                dataList.add(entity);
            }
        }
        return dataList;
    }

    /**
     * 校验表头
     * @param heads 表头字段
     */
    private static void checkExcelHead(Object[] heads,Class<?>entityClass) throws Exception {
        //校验表头字段是否存在重复表头
        if (heads == null || heads.length == 0){
            throw new Exception("表头为空");
        }
        Set<String>headSet = new HashSet<>();
        int length = heads.length;
        for (int i = 0; i < length; i++){
            String content = heads[i].toString();
            for (int j = i+1;j<length; j++){
                if (!"".equals(content) && heads[i].equals(heads[j])){
                    throw new Exception("表头中存在重复字段");
                }else {
                    headSet.add(content);
                }
            }
        }
        //获取@Excel注解的required为true的字段对应的value值（即Excel表必须具备的表头字段）
        Set<String> headsRequired = getHeadsRequired(entityClass);
        //校验表头中是否包含所有必须的字段
        for (String name : headsRequired){
            if (!headSet.contains(name)){
                throw new Exception("Excel表头中缺少“"+name+"”字段");
            }
        }
    }

    /**
     * 获取必须字段名
     * @param clazz 实体类
     * @return 必须字段名集合
     */
    private static Set<String> getHeadsRequired(Class<?> clazz) throws Exception {
        if (clazz == null){
            throw new Exception("类不能为空");
        }
        Field[] fields = clazz.getDeclaredFields();
        if (fields.length == 0){
            throw new Exception(clazz.getName() + " 类属性为空");
        }
        Set<String>nameSet = new HashSet<>();
        for (Field field : fields){
            if (field.getAnnotation(Excel.class)!= null && field.getAnnotation(Excel.class).required()){
                String value = field.getAnnotation(Excel.class).value();
                if (StringUtils.isEmpty(value)){
                    throw new Exception(clazz.getName() + " 类的"+field.getName()+"字段的@Excel注解的value值为空或空串");
                }else {
                    nameSet.add(value);
                }
            }
        }
        return nameSet;
    }

    /**
     * 获取工作簿
     *
     * @param file 文件
     * @return workbook
     */
    private static Workbook getWorkbook(MultipartFile file) throws IOException {
        String fileType = getFileType(file);
        if (StringUtils.isEmpty(fileType)) {
            throw new IOException("文件类型不明确");
        }
        if (EXCEL_XLS.equals(fileType)) {
            return new HSSFWorkbook(file.getInputStream());
        } else if (EXCEL_XLSX.equals(fileType)) {
            return new XSSFWorkbook(file.getInputStream());
        } else {
            throw new IOException("不支持的文件类型");
        }
    }

    /**
     * 获取文件类型
     *
     * @param file 文件
     * @return 文件类型
     */
    private static String getFileType(MultipartFile file) throws IOException {
        String filename = file.getOriginalFilename();
        if (StringUtils.isEmpty(filename)) {
            throw new IOException("文件名不能为空");
        } else {
            return FilenameUtils.getExtension(filename);
        }
    }

    /**
     * 检查文件
     *
     * @param file 文件
     * @throws IOException IO异常
     */
    private static void checkFile(MultipartFile file) throws IOException {
        if (file == null) {
            throw new FileNotFoundException("文件不存在，请检查后重试");
        }
        //获取文件名
        String originalFilename = file.getOriginalFilename();
        //获取后缀名（即文件类型）
        String extension = FilenameUtils.getExtension(originalFilename);
        if (StringUtils.isEmpty(extension)) {
            throw new IOException(originalFilename + "文件类型不明");
        } else if (!extension.equals(EXCEL_XLS) && !extension.equals(EXCEL_XLSX)) {
            throw new IOException(originalFilename + "不是Excel文件");
        }
    }

    /**
     * 给对象赋值
     *
     * @param entity 实体
     * @param objs   数据
     * @param heads  表头
     */
    private static void setValue(Object entity, Object[] objs, Object[] heads) throws Exception {
        //数据校验
        if (objs == null || objs.length == 0) {
            throw new Exception("数据不存在");
        }
        if (heads == null || heads.length == 0) {
            throw new Exception("表头不存在");
        }
        //获取字段@Excel注解的value与字段名的映射
        Map<String, Field> fields = getFiledByName(entity.getClass());

        if (fields.size() == 0) {
            throw new Exception("实体类不包含任何属性");
        }
        //创建属性与value值的映射
        Map<Field,String>fieldName = new LinkedHashMap<>(16);
        for (Map.Entry<String,Field> entry:fields.entrySet()){
            fieldName.put(entry.getValue(),entry.getKey());
        }

        //获取Excel表头与列编号的映射
        Map<String, Integer> headRowNumMap = new LinkedHashMap<>(16);
        for (int i = 0; i < heads.length; i++) {
            headRowNumMap.put(heads[i].toString(), i);
        }

        //获取字段名与列编号的映射
        Map<Integer, Field> fieldMap = new LinkedHashMap<>(16);
        for (Map.Entry<String, Field> entry : fields.entrySet()) {
            String headName = entry.getKey();
            if (headRowNumMap.get(headName) == null){
                //判断此字段是否为必须字段
                if (!checkRequired(entity.getClass(),entry.getValue())){
                    continue;
                }
            }
            fieldMap.put(headRowNumMap.get(headName), entry.getValue());
        }

        for (Map.Entry<Integer, Field> entry : fieldMap.entrySet()) {
            Field field = entry.getValue();
            if (field != null) {
                field.setAccessible(true);
                Class<?> fieldType = field.getType();
                Object obj;
                //明确要获取哪一列的数据
                Integer cellNum = entry.getKey();
                if (objs.length <= cellNum){
                    //检验该字段是否为必须字段
                    Boolean required = checkRequired(entity.getClass(),field);
                    //是必须字段却没数据
                    if (required){
                        throw new Exception("“"+fieldName.get(field) + "”字段的数据为必须数据，不能为空");
                    }
                    //不是必须字段也没数据
                    else {
                        continue;
                    }
                }else if (StringUtils.isEmpty(objs[entry.getKey()].toString())){
                    //检验该字段是否为必须字段
                    Boolean required = checkRequired(entity.getClass(),field);
                    //是必须字段却没数据
                    if (required) {
                        throw new Exception("“"+fieldName.get(field) + "”字段的数据为必须数据，不能为空");
                    }
                    //不是必须字段也没数据
                    else {
                        continue;
                    }
                }
                else {
                    obj = objs[entry.getKey()];
                }
                if (String.class == fieldType) {
                    field.set(entity, String.valueOf(obj));
                } else if (Integer.TYPE == fieldType || Integer.class == fieldType) {
                    if (obj.toString().contains(".")) {
                        obj = new Double(obj.toString()).intValue();
                    }
                    field.set(entity, Integer.parseInt(obj.toString()));
                } else if (Long.TYPE == fieldType || Long.class == fieldType) {
                    field.set(entity, Long.valueOf(obj.toString()));
                } else if (Float.TYPE == fieldType || Float.class == fieldType) {
                    field.set(entity, Float.valueOf(obj.toString()));
                } else if (Short.TYPE == fieldType || Short.class == fieldType) {
                    field.set(entity, Short.valueOf(obj.toString()));
                } else if (Double.TYPE == fieldType || Double.class == fieldType) {
                    field.set(entity, Double.valueOf(obj.toString()));
                } else if (Date.class == fieldType) {
                    if (StringUtils.isNotEmpty(obj.toString())) {
                        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        Date date = format.parse(obj.toString());
                        field.set(entity, date);
                    }
                } else if (Character.TYPE == fieldType) {
                    if ((obj != null) && (obj.toString().length() > 0)) {
                        field.set(entity, Character.valueOf(obj.toString().charAt(0)));
                    }
                } else {
                    field.set(entity, obj);
                }
            } else {
                throw new Exception("存在不允许导入的字段");
            }
        }
    }

    /**
     * 检查字段
     * @param clazz 实体类
     * @param field 字段
     * @return true:是必须字段 false:非必须字段
     */
    private static Boolean checkRequired(Class<?> clazz,Field field) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field fie : fields){
            if (fie.getAnnotation(Excel.class) != null && fie.getAnnotation(Excel.class).required() && fie.equals(field)){
                return true;
            }
        }
        return false;
    }

    /**
     * 获取类的全部字段
     *
     * @param clazz 类
     * @return
     */
    private static Map<String, Field> getFiledByName(Class<?> clazz) throws Exception {
        Map<String, Field> fieldMap = new LinkedHashMap<>(16);
        Field[] fields = clazz.getDeclaredFields();
        if (fields.length == 0) {
            throw new Exception("此实体类不含任何属性");
        }
        for (Field field : fields) {
            if (field.getAnnotation(Excel.class) != null) {
                fieldMap.put(field.getAnnotation(Excel.class).value(), field);
            }
        }
        return fieldMap;
    }

    /**
     * 获取字段名
     *
     * @param clazz 实体类
     * @return 字段名
     */
    private static Map<Integer, String> getFiledNameByClassName(Class clazz) throws Exception {
        Map<Integer, String> fieldNameMap = new LinkedHashMap<>(16);
        Field[] fields = clazz.getDeclaredFields();
        if (fields.length == 0) {
            throw new Exception("此实体类不含任何属性");
        }
        for (int i = 0; i < fields.length; i++) {
            if (fields[i].getAnnotation(Excel.class) != null) {
                String name = fields[i].getAnnotation(Excel.class).value();
                if (StringUtils.isEmpty(name)){
                    throw new Exception(clazz.getName() + " 类中，字段"+fields[i].getName()+"的@Excel注解的value值不能为空或空串");
                }else {
                    fieldNameMap.put(i, name);
                }
            }
        }
        //检查有无重复的注解字段名（即@Excel注解的value值）
        Set<String> name = new HashSet<>();
        for (Map.Entry<Integer,String> entry:fieldNameMap.entrySet()){
            name.add(entry.getValue());
        }
        if (name.size() != fieldNameMap.size()){
            throw new Exception(clazz.getName() + " 类中@Excel注解的value值存在重复现象");
        }
        return fieldNameMap;
    }

    /**
     * 导出Excel
     *
     * @param fileType 文件类型（xls/xlsx）
     * @param title    sheet名
     * @param data     数据
     * @param clazz    传入数据使用的类
     * @param out      输出流
     * @param <T>
     * @throws Exception
     */
    public static <T> void exportExcel(String fileType, String title, Integer sheetSize, List<T> data, Class clazz, OutputStream out)
            throws Exception {
        //校验数据
        if (!EXCEL_XLS.equals(fileType) && !EXCEL_XLSX.equals(fileType)) {
            throw new Exception("不支持的文件格式");
        }
        if (out == null) {
            throw new Exception("未确定输出目标流");
        }
        if (StringUtils.isEmpty(title)){
            title = "工作表";
        }
        //设置单页行数
        if (sheetSize == null || sheetSize <= 0) {
            sheetSize = 10000;
        }
        //计算页数,不足一页算一页
        int pages = data.size() / sheetSize;
        if (data.size() % sheetSize > 0) {
            pages += 1;
        }
        Workbook workbook = getWorkbook(fileType);
        Map<Integer, String> fieldNameMap = getFiledNameByClassName(clazz);
        Map<String, Field> fieldMap = getFiledByName(clazz);
        Map<Integer,Field> tempMap = new LinkedHashMap<>(16);
        for (Map.Entry<Integer,String> entry:fieldNameMap.entrySet()){
            tempMap.put(entry.getKey(),fieldMap.get(fieldNameMap.get(entry.getKey())));
        }
        for (int i = 0; i < pages; i++) {
            int startData = i * sheetSize;
            int endData = (i + 1) * sheetSize - 1 > data.size() ? data.size() : (i + 1) * sheetSize - 1;
            int rowNum = 0;
            Sheet sheet;
            //设置sheet名
            if (pages > 1) {
                sheet = workbook.createSheet(title+Integer.valueOf(i+1).toString());
            } else {
                sheet = workbook.createSheet(title);
            }
            Row row = sheet.createRow(rowNum++);
            //设置表头
            for (int j = 0; j < fieldNameMap.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(fieldNameMap.get(j));
            }
            //填充数据
            for (int j = startData; j <= endData; j++) {
                row = sheet.createRow(rowNum++);
                T item = data.get(j);
                for (int k = 0; k < fieldNameMap.size(); k++) {
                    Field field = tempMap.get(k);
                    field.setAccessible(true);
                    Object obj = field.get(item);
                    String value;
                    if (obj == null) {
                        value = "";
                    } else {
                        if (field.getType() == Date.class) {
                            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                            value = format.format(obj);
                        } else {
                            value = obj.toString();
                        }
                    }
                    Cell cell = row.createCell(k);
                    cell.setCellValue(value);
                }
            }
        }
        workbook.write(out);
    }

    /**
     * 根据需要文件类型，获取工作簿
     *
     * @param fileType 所需文件类型
     * @return workbook
     */
    private static Workbook getWorkbook(String fileType) throws IOException {
        if (StringUtils.isEmpty(fileType)) {
            throw new IOException("未规定文件类型");
        }
        if (EXCEL_XLS.equals(fileType)) {
            return new HSSFWorkbook();
        } else if (EXCEL_XLSX.equals(fileType)) {
            return new XSSFWorkbook();
        } else {
            throw new IOException("不支持此文件类型");
        }
    }
}
