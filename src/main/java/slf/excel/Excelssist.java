package slf.excel;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class Excelssist {
    private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    private Workbook wb;
    private Sheet sheet;
    private Map<Integer, String> propertyMap = new TreeMap<>();
    private Map<Integer, String> cellMap = new TreeMap<>();
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

    public <T> void cellParse(T t) {
        Class tClass = t.getClass();
        Field[] fields = tClass.getDeclaredFields();
        for (Field field : fields) {
            slf.excel.Cell cell = field.getAnnotation(slf.excel.Cell.class);
            if (cell != null) {
                cellMap.put(cell.num(), field.getName());
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

    public <T> List<T> excelToObject(int rowNum, T targetObject) {
        propertiesParse(targetObject);
        List<String> propertyList = new ArrayList<>();
        propertyMap.forEach((k, v) -> propertyList.add(v));
        return excelToObject(rowNum, targetObject, propertyList);
    }

    public <T> HSSFWorkbook objectToExcel(String sheetName, T object, List<String> cellList) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);
        Class clazz = object.getClass();
        return workbook;
    }

    public <T> List<T> excelToObject(int rowNum, T object, List<String> propertyList) {
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
                T tClone = (T) BeanUtils.cloneBean(object);
                Class tClass = tClone.getClass();
                for (int col = 0; col < propertyList.size(); col++) {
                    String property = propertyList.get(col);
                    Field field = tClass.getDeclaredField(property);
                    field.setAccessible(true);
                    String type = field.getType().getTypeName();
                    Cell xcell = xrow.getCell(col);
                    String cellValue = "";
                    int cellType = xcell.getCellType();
                    switch (cellType) {
                        case Cell.CELL_TYPE_STRING:
                            cellValue = xcell.getStringCellValue().trim();
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            cellValue = String.valueOf(xcell.getBooleanCellValue());
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            cellValue = String.valueOf(xcell.getCellFormula().trim());
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            if (HSSFDateUtil.isCellDateFormatted(xcell)) {
                                cellValue = sdf.format(xcell.getDateCellValue());
                            } else {
                                cellValue = new DecimalFormat("#.##").format(xcell.getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_BLANK:
                            cellValue = "NULL";
                            break;
                        case Cell.CELL_TYPE_ERROR:
                            cellValue = "ERROR";
                            break;
                        default:
                            cellValue = xcell.toString().trim();
                            break;
                    }
                    try {
                        switch (type) {
                            case "java.lang.Integer":
                                xcell.setCellType(CellType.STRING);
                                field.set(tClone, Integer.valueOf(cellValue));
                                break;
                            case "java.lang.String":
                                field.set(tClone, cellValue);
                                break;
                            case "java.math.BigDecimal":
                                field.set(tClone, new BigDecimal(cellValue).setScale(2,   BigDecimal.ROUND_HALF_UP));
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
