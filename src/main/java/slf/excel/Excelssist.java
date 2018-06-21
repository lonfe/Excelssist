package slf.excel;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

public class Excelssist {
    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    private Workbook wb;
    private Sheet sheet;
    private Map<Integer, String> propertyMap = new TreeMap<>();

    public Excelssist() {

    }

    public Excelssist(String filePath) {
        try {
            InputStream in = new FileInputStream(filePath);
            if (filePath.endsWith(".xls") || filePath.endsWith(".et")) {
                wb = new HSSFWorkbook(in);
            } else {
                wb = new XSSFWorkbook(in);
            }
            sheet = wb.getSheetAt(0);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public <T> void propertiesParse(T t) {
        Class tClass = t.getClass();
        Field[] fields = tClass.getDeclaredFields();
        for (Field field : fields) {
            Sign sign = field.getAnnotation(Sign.class);
            if (sign != null) {
                propertyMap.put(sign.num(), field.getName());
            }
        }
    }

    public int getRowLen() {
        int lastRowNum = sheet.getLastRowNum();
        if (lastRowNum == 0) {
            Row xRow = sheet.getRow(lastRowNum);
            if (xRow != null)
                return ++lastRowNum;
            else {
                return lastRowNum;
            }
        }
        return ++lastRowNum;
    }

    public int getColLenByAppointRow(int rowNum) {
        if (rowNum < 0) {
            throw new IllegalArgumentException("illegal parameter, rowNum: " + rowNum);
        }
        if (rowNum > getRowLen()) {
            throw new IllegalArgumentException("illegal parameter, rowNum: " + rowNum + " > rowMaxNum: " + getRowLen());
        }
        int colLen = 0;
        Row row = sheet.getRow(rowNum);
        if (row != null) {
            colLen = row.getLastCellNum();
        }
        return colLen;
    }

    public <T> List<T> cellValueToTargetObject(int rowNum, T targetObject) {
        propertiesParse(targetObject);
        List<String> propertyList = new ArrayList<>();
        propertyMap.forEach((k, v) -> propertyList.add(v));
        return cellValueToTargetObject(rowNum, targetObject, propertyList);
    }

    public <T> List<T> cellValueToTargetObject(int rowNum, T targetObject, List<String> propertyList) {
        if (rowNum <= 0) {
            throw new IllegalArgumentException("illegal parameter, rowNum: " + rowNum);
        }
        if (rowNum > getRowLen()) {
            throw new IllegalArgumentException("illegal parameter, rowNum: " + rowNum + ", rowMaxNum: " + getRowLen());
        }
        List<T> targetList = new ArrayList<>();
        for (int row = --rowNum; row < getRowLen(); row++) {
            try {
                Row xrow = sheet.getRow(row);
                T tClone = (T) BeanUtils.cloneBean(targetObject);
                Class tClass = tClone.getClass();
                for (int col = 0; col < propertyList.size(); col++) {
                    String property = propertyList.get(col);
                    Field field = tClass.getDeclaredField(property);
                    field.setAccessible(true);
                    String type = field.getType().getTypeName();
                    Cell xcell = xrow.getCell(col);
                    xcell.setCellType(CellType.STRING);
                    String cellValue = xcell.toString();
                    try {
                        switch (type) {
                            case "java.lang.Integer":
                                field.set(tClone, Integer.valueOf(cellValue));
                                break;
                            case "java.lang.String":
                                field.set(tClone, cellValue);
                                break;
                            case "java.lang.BigDecimal":
                                field.set(tClone, new BigDecimal(cellValue));
                                break;
                            case "java.lang.Double":
                                field.set(tClone, Double.valueOf(cellValue));
                                break;
                            case "java.lang.Long":
                                field.set(tClone, Long.valueOf(cellValue));
                                break;
                            case "java.lang.Date":
                                field.set(tClone, sdf.parse(cellValue));
                                break;
                        }
                    } catch (IllegalArgumentException e) {
                        throw new IllegalAccessException("property set error, row: " + (++row) + ", column: "+ (++col) + ", property type: " + type + ", value: " + cellValue);
                    }
                }
                targetList.add(tClone);
            } catch (Exception e) {
                e.printStackTrace();
                return Collections.EMPTY_LIST;
            }
        }
        return targetList;
    }
}
