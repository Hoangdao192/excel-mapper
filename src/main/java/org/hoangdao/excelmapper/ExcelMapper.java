package org.hoangdao.excelmapper;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.format.annotation.DateTimeFormat;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.Instant;
import java.util.*;

public class ExcelMapper {

    public <T> void write(List<T> items, Class<T> c, OutputStream outputStream) throws IOException {
        write(items, c, outputStream, new ArrayList<>());
    }

    public <T> void write(List<T> items, Class<T> c, OutputStream outputStream, List<String> ignoredProperties) throws IOException {
        Annotation[] annotations = c.getDeclaredAnnotations();
        String sheetName = c.getName();
        for (Annotation annotation : annotations) {
            if (annotation instanceof ExcelSheet) {
                sheetName = ((ExcelSheet) annotation).name();
            }
        }
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            Row headRow = createHeadRow(sheet, c);
            int rowIndex = 1;
            for (T item : items) {
                createRow(
                        sheet, rowIndex, item, item.getClass(), ignoredProperties
                );
                ++rowIndex;
            }
            workbook.write(outputStream);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    private <T> void createRow(Sheet sheet, int rowIndex, T data, Class<?> c, List<String> ignoredProperties)
            throws IllegalAccessException {
        Row headRow = sheet.getRow(0);
        List<Field> fields = getAllFields(c);
        Row row = sheet.createRow(rowIndex);

        for (Field field : fields) {
            if (ignoredProperties.contains(field.getName())) {
                continue;
            }
            Annotation[] annotations = field.getDeclaredAnnotations();
            for (Annotation annotation : annotations) {
                if (annotation instanceof CellProperty) {
                    String header = ((CellProperty) annotation).value();
                    int columnIndex = findCellIndexByContent(
                            headRow, header
                    );
                    field.setAccessible(true);
                    if (columnIndex >= 0 && field.get(data) != null) {
                        Cell cell = row.createCell(columnIndex);
                        field.setAccessible(true);
                        if (field.getType() == Date.class) {
                            String datePattern = "yyyy-MM-dd";
                            for (Annotation dateAnnotation : annotations) {
                                if (dateAnnotation instanceof DateTimeFormat) {
                                    DateTimeFormat dateTimeFormat = (DateTimeFormat) dateAnnotation;
                                    if (!dateTimeFormat.pattern().isEmpty()) {
                                        datePattern = dateTimeFormat.pattern();
                                    }
                                }
                            }
                            cell.setCellValue(DateUtil.formatDate((Date) field.get(data), datePattern));
                        } else {
                            cell.setCellValue(
                                    field.get(data).toString()
                            );
                        }
                    }
                }
            }
        }
    }

    private int findCellIndexByContent(Row row, String content) {
        Iterator<Cell> iterator = row.cellIterator();
        while (iterator.hasNext()) {
            Cell cell = iterator.next();
            if (cell.getStringCellValue().equals(content)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    private <T> Row createHeadRow(Sheet sheet, Class<T> c) {
        try {
            List< HeaderInfo> headerInfoList = new ArrayList<>();
            List<String> headers = new ArrayList<>();
            T t = c.getDeclaredConstructor().newInstance();
            List<Field> fields = getAllFields(c);
            for (Field field : fields) {
                String headerName = null;
                int headerIndex = -1;
                Annotation[] annotations = field.getDeclaredAnnotations();
                for (Annotation annotation : annotations) {
                    if (annotation instanceof CellProperty) {
                        headerName = ((CellProperty) annotation).value();
                        headerIndex = ((CellProperty) annotation).index();
                        headers.add(
                                ((CellProperty) annotation).value()
                        );
                    }
                }

                headerInfoList.add(new HeaderInfo(headerName, headerIndex));
            }

            Row row = sheet.createRow(0);
            CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            int columnIndex = 0;
            for (HeaderInfo headerInfo : headerInfoList) {
                int currentIndex = columnIndex;
                if (headerInfo.getIndex() != -1) {
                    currentIndex = headerInfo.getIndex();
                }
                Cell cell = row.createCell(currentIndex);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(headerInfo.getName());
                ++columnIndex;
            }
            return row;
        } catch (NoSuchMethodException e) {
            throw new RuntimeException("Provided class does not have default constructor");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private Object convertStringToType(String text, Field field) {
        Class<?> c = field.getType();
        Annotation[] annotations = field.getDeclaredAnnotations();
        if (text != null) {
            text = text.trim();

            if (c == Long.class) {
                return Long.parseLong(text);
            }
            else if (c == Integer.class) {
                return Integer.parseInt(text);
            }
            else if (c == Boolean.class) {
                return Boolean.parseBoolean(text);
            }
            else if (c == String.class) {
                return text;
            }
            else if (c == Double.class) {
                return Double.parseDouble(text);
            }
            else if (c == Float.class) {
                return Float.parseFloat(text);
            }
            else if (c == Short.class) {
                return Short.parseShort(text);
            }
            else if (c == Instant.class) {
                return Instant.parse(text);
            }
            else if (c == Date.class) {
                String datePattern = "yyyy-MM-dd";
                for (Annotation annotation : annotations) {
                    if (annotation instanceof DateTimeFormat) {
                        DateTimeFormat dateTimeFormat = (DateTimeFormat) annotation;
                        if (!dateTimeFormat.pattern().isEmpty()) {
                            datePattern = dateTimeFormat.pattern();
                        }
                    }
                }
                return DateUtil.parseDate(text, datePattern);
            }
            else {
                throw new RuntimeException("Unsupported data types");
            }
        }
        return null;
    }

    public <T> List<T> read(InputStream inputStream, Class<T> c) throws IOException {
        List<T> list = new ArrayList<>();

        Annotation[] annotations = c.getDeclaredAnnotations();
        String sheetName = null;
        for (Annotation annotation : annotations) {
            if (annotation instanceof ExcelSheet) {
                sheetName = ((ExcelSheet) annotation).name();
            }
        }
        List<Field> fields = getAllFields(c);

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = null;
        if (sheetName != null) {
            sheet = workbook.getSheet(sheetName);
        } else {
            sheet = workbook.getSheetAt(0);
        }

        if (sheet != null) {
            Iterator<Row> rowIterator = sheet.rowIterator();
            Row headRow = null;
            if (rowIterator.hasNext()) {
                headRow = rowIterator.next();
            }

            if (headRow != null) {
                while (rowIterator.hasNext()) {
                    Row currentRow = rowIterator.next();
                    try {
                        T instance = c.getConstructor().newInstance();
                        Iterator<Cell> cellIterator = currentRow.cellIterator();
                        List<Cell> cells = new ArrayList<>();
                        while (cellIterator.hasNext()) {
                            cells.add(cellIterator.next());
                        }
                        for (Field field : fields) {
                            String propertyName = null;
                            int propertyIndex = -1;
                            Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                            for (Annotation annotation : fieldAnnotations) {
                                if (annotation instanceof CellProperty) {
                                    propertyName = ((CellProperty) annotation).value();
                                    propertyIndex = ((CellProperty) annotation).index();
                                }
                            }
                            Cell cell = null;
                            if (propertyIndex != -1) {
                                cell = currentRow.getCell(propertyIndex);
                            } else {
                                for (Cell item : cells) {
                                    if (getCellValue(item).equals(propertyName)) {
                                        cell = item;
                                        break;
                                    }
                                }
                            }
                            if (cell != null) {
                                String cellValueAsString = getCellValue(cell);
                                if (cellValueAsString != null) {
                                    cellValueAsString = cellValueAsString.trim();
                                    boolean fieldAccessible = field.canAccess(instance);
                                    field.setAccessible(true);
                                    field.set(instance, convertStringToType(cellValueAsString, field));
                                    field.setAccessible(fieldAccessible);
                                }
                            }
                        }
                        list.add(instance);
                    } catch (NoSuchMethodException | InstantiationException | IllegalAccessException |
                             InvocationTargetException e) {
                        throw new RuntimeException("Class does not have no arguments constructor or constructor not public");
                    }
                }
            }
        }
        return list;
    }

    private List<Field> getAllFields(Class<?> type) {
        List<Field> fields = new ArrayList<>(Arrays.asList(type.getDeclaredFields()));
        if (type.getSuperclass() != null) {
            fields.addAll(getAllFields(type.getSuperclass()));
        }
        return fields;
    }

    private String getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }

        return String.valueOf(cellValue);
    }


}
