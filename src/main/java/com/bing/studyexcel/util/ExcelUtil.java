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
                //跳过首行（表头）
                if (row.getRowNum() < EXCEL_HEAD_ROW_NUM) {
                    heads = new Object[num];
                    for (int k = 0; k < num; k++) {
                        if (row.getCell(k) == null){
                            heads[k] = "";
                        }
                        else
                        {
                            heads[k] = row.getCell(k);
                        }
                    }
                    continue;
                }
                T entity = entityClass.newInstance();
                //获得表格某行的值
                Object[] objects = new Object[num];
                for (int k = 0; k < num; k++) {
                    if (row.getCell(k) == null){
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
        if (objs == null || objs.length == 0) {
            throw new Exception("数据不存在");
        }
        if (heads == null || heads.length == 0) {
            throw new Exception("表头不存在");
        }
        //获取字段@Excel注解的value与字段名的映射
        Map<String, Field> fields = getFiledByName(entity.getClass());

        if (fields == null || fields.size() == 0) {
            throw new Exception("实体类不包含任何属性");
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
            fieldMap.put(headRowNumMap.get(headName), fields.get(headName));
        }

        for (Map.Entry<Integer, Field> entry : fieldMap.entrySet()) {
            Field field = entry.getValue();
            if (field != null) {
                field.setAccessible(true);
                Class<?> fieldType = field.getType();
                if (String.class == fieldType) {
                    field.set(entity, String.valueOf(objs[entry.getKey()]));
                } else if (Integer.TYPE == fieldType || Integer.class == fieldType) {
                    if (objs[entry.getKey()].toString().contains(".")) {
                        objs[entry.getKey()] = new Double(objs[entry.getKey()].toString()).intValue();
                    }
                    field.set(entity, Integer.parseInt(objs[entry.getKey()].toString()));
                } else if (Long.TYPE == fieldType || Long.class == fieldType) {
                    field.set(entity, Long.valueOf(objs[entry.getKey()].toString()));
                } else if (Float.TYPE == fieldType || Float.class == fieldType) {
                    field.set(entity, Float.valueOf(objs[entry.getKey()].toString()));
                } else if (Short.TYPE == fieldType || Short.class == fieldType) {
                    field.set(entity, Short.valueOf(objs[entry.getKey()].toString()));
                } else if (Double.TYPE == fieldType || Double.class == fieldType) {
                    field.set(entity, Double.valueOf(objs[entry.getKey()].toString()));
                } else if (Date.class == fieldType) {
                    if (StringUtils.isNotEmpty(objs[entry.getKey()].toString())) {
                        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        Date date = format.parse(objs[entry.getKey()].toString());
                        field.set(entity, date);
                    }
                } else if (Character.TYPE == fieldType) {
                    if ((objs[entry.getKey()] != null) && (objs[entry.getKey()].toString().length() > 0)) {
                        field.set(entity, Character.valueOf(objs[entry.getKey()].toString().charAt(0)));
                    }
                } else {
                    field.set(entity, objs[entry.getKey()]);
                }
            } else {
                throw new Exception("存在不允许导入的字段");
            }
        }
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
        if (StringUtils.isEmpty(fileType) || (!EXCEL_XLS.equals(fileType) && !EXCEL_XLSX.equals(fileType))) {
            throw new Exception("请确认要导出的文件类型为Excel文件格式");
        }
        if (out == null) {
            throw new Exception("未确定输出目标流");
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
                sheet = workbook.createSheet(title + i);
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
            for (int j = startData; j < endData; j++) {
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
