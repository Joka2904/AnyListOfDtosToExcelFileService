
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.text.SimpleDateFormat;
import java.util.*;

@Service
@RequiredArgsConstructor
public class AnyListOfDtoToExcelServiceImpl {

    public StreamingResponseBody anyDtoToExcel(List<?> anyDto) {

        XSSFWorkbook workbook = new XSSFWorkbook();

        String mainSheetName = anyDto.get(0).getClass().getSimpleName();
        log.warn(mainSheetName);
        Map<String, Integer> indexOfRow = new HashMap<>();
        getSheet(workbook, mainSheetName, indexOfRow);

        for (Object object : anyDto) {
            LinkedHashMap<String, Object> objectMap = objectToMap(object);
            writeObjectToExcel(objectMap, incrementIndexOfRow(indexOfRow, mainSheetName), workbook, mainSheetName);
        }
        return workbook::write;
    }

    public static void writeObjectToExcel(Map<String, Object> objectMap, Map<String, Integer> indexOfRow, XSSFWorkbook workbook, String sheetName) {

        XSSFSheet sheet = getSheet(workbook, sheetName, indexOfRow);
        createHeaderOfSheet(sheet, objectMap.keySet());
        XSSFRow row = sheet.createRow(indexOfRow.get(sheetName));

        int cellIndex = 0;
        for (var object : objectMap.entrySet()) {

            String key = object.getKey();
            Object value = object.getValue();
            if (value instanceof LinkedHashMap) {
                LinkedHashMap<String, Object> valueMap = (LinkedHashMap<String, Object>) value;

                XSSFSheet sheetMap = getSheet(workbook, key, indexOfRow);

                var idForList = String.format("%sId", sheetName.toLowerCase(Locale.ROOT));
                valueMap.put(idForList, objectMap.get("id"));

                row.createCell(cellIndex).setCellValue(String.valueOf(((LinkedHashMap<?, ?>) value).get("id")));

                writeObjectToExcel(valueMap, incrementIndexOfRow(indexOfRow, key), workbook, key);
                createHeaderOfSheet(sheetMap, valueMap.keySet());

            } else if (value instanceof ArrayList && ((ArrayList<?>) value).get(0) instanceof LinkedHashMap) {
                ArrayList<LinkedHashMap<String, Object>> valueList = (ArrayList<LinkedHashMap<String, Object>>) value;

                XSSFSheet sheetListMap = getSheet(workbook, key, indexOfRow);
                Set<String> valuesOfHeader = valueList.get(0).keySet();

                List<String> idOfValueList = new ArrayList<>();
                for (var valueListMap: valueList){
                    var id = String.valueOf(valueListMap.get("id"));
                    idOfValueList.add(id);
                }
                row.createCell(cellIndex).setCellValue(String.valueOf(idOfValueList));

                for (var valueOfList : valueList) {
                    var idForList = String.format("%sId", sheetName.toLowerCase(Locale.ROOT));
                    valueOfList.put(idForList, objectMap.get("id"));
                    writeObjectToExcel(valueOfList, incrementIndexOfRow(indexOfRow, key), workbook, key);
                    createHeaderOfSheet(sheetListMap, valuesOfHeader);
                }
            } else {
                row.createCell(cellIndex).setCellValue(String.valueOf(value));
            }
            cellIndex++;
        }
    }

    private static Map<String, Integer> incrementIndexOfRow(Map<String, Integer> indexOfRow, String sheetName) {
        indexOfRow.merge(sheetName, 1, Integer::sum);
        return indexOfRow;
    }

    public static XSSFSheet getSheet(XSSFWorkbook workbook, String sheetName, Map<String, Integer> indexOfRow) {
        XSSFSheet sheet;
        if (workbook.getSheet(sheetName) == null) {
            sheet = workbook.createSheet(sheetName);
            indexOfRow.put(sheetName, 0);
        } else {
            sheet = workbook.getSheet(sheetName);
        }
        sheet.setDefaultColumnWidth(15);
        return sheet;
    }

    public static void createHeaderOfSheet(XSSFSheet sheet, Set<String> headerCellNameFromKey) {
        if (sheet.getRow(0) == null) {

            XSSFRow rowHeader = sheet.createRow(0);

            CellStyle style;

            XSSFWorkbook workbook = sheet.getWorkbook();

            XSSFFont font = workbook.createFont();
            font.setFontHeightInPoints((short) 10);
            font.setFontName("Arial");
            font.setColor(IndexedColors.DARK_BLUE.getIndex());
            font.setBold(false);
            font.setItalic(true);

            style = workbook.createCellStyle();

            style.setFont(font);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setHidden(true);

            int cellIndex = 0;
            for (var cellValue : headerCellNameFromKey) {
                XSSFCell cellHeader = rowHeader.createCell(cellIndex++);
                cellHeader.setCellStyle(style);
                cellHeader.setCellValue(cellValue);
            }
        }
    }

    public static LinkedHashMap<String, Object> objectToMap(Object object) {
        return mapper().convertValue(object, LinkedHashMap.class);
    }

    public static ObjectMapper mapper() {
        ObjectMapper mapper = new ObjectMapper();
        mapper.registerModules(new JavaTimeModule());
        mapper.setDateFormat(new SimpleDateFormat());
        return mapper;
    }


}
