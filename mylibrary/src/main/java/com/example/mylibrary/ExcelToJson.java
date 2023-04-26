package com.example.mylibrary;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.serializer.SerializerFeature;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;


// String a = new String("Tiếng Việt".getBytes("utf-8")); // 越南文

public class ExcelToJson {

    private static String inputPath = "G:\\codeTest\\excelToJson\\mylibrary\\resource.xlsx";
    private static String outputJsonPath = "G:\\codeTest\\excelToJson\\mylibrary\\out\\langs.json";
    public static void main(String[] args) {
        pp("main log start");

//        List<String> languageNeed = Arrays.asList("zh", "en", "in", "hi", "tr", "es", "pt", "ru", "th", "ar", "ur", "vi");
        List<String> languageNeed = Arrays.asList("zh", "en", "in", "vi");
        System.out.println("从表中需要读取的语言: " + languageNeed);
        ArrayList<String> listAllL10n = new ArrayList<>();

        HashMap<String, HashMap<String, String>> mapAll = new HashMap<>();
        LinkedList<String> keyAll = new LinkedList<>();

        HashMap<String, JsonObject> mapJsonLong = new HashMap<>();
        HashMap<String, JsonObject> mapJsonShort = new HashMap<>();

        String filePath = inputPath;
        Workbook wb = readExcel(filePath);
        if (wb != null) {
            //用来存放表中数据
            //            list = new ArrayList<Map<String, String>>();
            Sheet sheet = wb.getSheetAt(0);
            int rownum = sheet.getPhysicalNumberOfRows();
            Row row0 = sheet.getRow(0);
            Row row = sheet.getRow(0);

            int colnum = row.getPhysicalNumberOfCells();
            System.out.println(rownum + "~" + colnum);

            System.out.println(sheet.getFirstRowNum() + "~" + sheet.getLastRowNum() + "~" + sheet);


            for (int j = 0; j < colnum; j++) {
                String cellData = (String) getCellFormatValue(row0.getCell(j));
                if (languageNeed.contains(cellData)) {
                    mapJsonLong.put(cellData, new JsonObject());
                    mapJsonShort.put(cellData, new JsonObject());
                    listAllL10n.add(cellData);
                } else {

                    listAllL10n.add("");
                }
            }


            for (int i = 0; i < rownum; i++) {
                Map<String, String> map = new LinkedHashMap<String, String>();

                row = sheet.getRow(i);
                if (row != null) {

                    for (int z = 0; z < colnum; z++) {
                        String value_r0_key = getCellFormatValue(row0.getCell(z));
                        String value_r = getCellFormatValue(row.getCell(z));
                        String key_r = getCellFormatValue(row.getCell(0));
                        mapAll.putIfAbsent(key_r, new HashMap<>());
//                        HashMap<String,String> map_r = new HashMap<>();
//                        map_r.put(value_r0_key, value_r);
                        mapAll.get(key_r).put(value_r0_key, value_r);
//                        mapAll.put(key_r, map_r);
                        keyAll.add(key_r);


                        String languageKey = listAllL10n.get(z);
                        if (languageNeed.contains(languageKey)) {
                            String cellKey = (String) getCellFormatValue(row.getCell(0));
//                            String cellId = (String) getCellFormatValue(row.getCell(1));
                            String cellAlias = (String) getCellFormatValue(row.getCell(1));
                            int zh = listAllL10n.indexOf("zh");
                            if (zh < 0) return;
                            String cellZH = (String) getCellFormatValue(row.getCell(zh));
                            String cellEN = (String) getCellFormatValue(row.getCell(listAllL10n.indexOf("en")));
                            String cells = (String) getCellFormatValue(row.getCell(z));
                            if (cells.isEmpty()) {
                                cells = cellEN;
                            }
//                            if (cells.isEmpty()) {
//                                cells = cellZH;
//                            }

                            // LinkedHashMap<String, Object> linkedHashMapLong = mapAllL10n.get(languageKey);
                            // linkedHashMapLong.put(cellKey, cells);
                            // LinkedHashMap<String, Object> cellsMap = new LinkedHashMap<>();
                            // cellsMap.put("k_0", cellKey);
                            // cellsMap.put("k_1", cellId);
                            // cellsMap.put("k_2", cellAlias);
                            // cellsMap.put("k_3", cellZH);
                            // linkedHashMapLong.put("@" + cellKey, cellsMap);

                            JsonObject itemInfo = new JsonObject();
                            itemInfo.addProperty("语种", languageKey);
                            itemInfo.addProperty("键名", cellKey);
//                            itemInfo.addProperty("旧id", cellId);
                            itemInfo.addProperty("别名", cellAlias);
                            itemInfo.addProperty("中文", cellZH);


                            JsonObject jsonObjectLong = mapJsonLong.get(languageKey);
                            jsonObjectLong.addProperty(cellKey, cells.trim());
                            jsonObjectLong.add("@" + cellKey, itemInfo);

                            JsonObject jsonObjectShort = mapJsonShort.get(languageKey);
                            jsonObjectShort.addProperty(cellKey, cells.trim());

                            if (z == listAllL10n.indexOf("zh") &&
                                    !cellKey.isEmpty()
                            ) {

                            }
                        }
                    }
                } else {
                    continue;
                }
//                list.add(map);
            }

            for (String lang : listAllL10n) {
                if (!languageNeed.contains(lang)) {
                    continue;
                }
                for (String key : keyAll) {
                    if (key.isEmpty()) {
//                        System.out.println("key empty~" + key + "~");
                        continue;
                    }
                    HashMap<String, String> map_key = mapAll.get(key);
                    HashMap<String, String> map_key_bak = new HashMap<>();

                    String alias = map_key.get("alias");
                    String alias_value = "";
//                    System.out.println("alias~~~~" + alias + "~~~~" + key + "~");
//                    System.out.println("alias~~~~" + alias.isEmpty() + "~~~~");

                    // StringUtils.equals(lang, "zh")
                    if (!alias.isEmpty() && !"zh".equals(lang) && !alias.equals("alias")) {
                        map_key_bak = mapAll.get(alias);
                        System.out.println("alias  alias  ~~~~" + mapAll.get(alias) + "~" + alias);
                        alias_value = map_key_bak.get(lang);
                    }

                    String value = map_key.get(lang);
                    if (!alias_value.isEmpty()) {
                        value = alias_value;
                    }
                    if (value.isEmpty()) {
                        value = map_key.get("en");
                    }

                }
            }

            JsonObject jsonInfo = new JsonObject();
            JsonObject jsonShort = new JsonObject();
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            Date date = new Date();
            String timeBuildStr = simpleDateFormat.format(date);
            jsonInfo.addProperty("_timeBuild", date.getTime());
            jsonInfo.addProperty("_timeBuildStr", timeBuildStr);
            jsonShort.addProperty("_timeBuild", date.getTime());
            jsonShort.addProperty("_timeBuildStr", timeBuildStr);

            for (int z = 0; z < colnum; z++) {
                String languageKey = listAllL10n.get(z);
                if (languageNeed.contains(languageKey)) {
                    JsonObject jsonObjectLong = mapJsonLong.get(languageKey);
                    JsonObject jsonObjectShort = mapJsonShort.get(languageKey);
                    jsonInfo.add(languageKey, jsonObjectLong);
                    jsonShort.add(languageKey, jsonObjectShort);
                }
            }

            jsonInfo.add("ms", mapJsonLong.get("id"));
            jsonShort.add("ms", mapJsonShort.get("id"));

            jsonShort.remove("zh");
            // jsonShort.remove("ms");
            writeToJson(toJsonStr(jsonShort), outputJsonPath);


        }
    }


    //读取excel
    public static Workbook readExcel(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(extString)) {
                return wb = new HSSFWorkbook(is);
            } else if (".xlsx".equals(extString)) {
                return wb = new XSSFWorkbook(is);
            } else {
                return wb = null;
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }


    public static String getCellFormatValue(Cell cell) {

        String cellValue = "";
        if (cell != null) {
            CellType copyCellType = cell.getCellType();

            switch (copyCellType) {
                //case _NONE:{ // -1
                //    break;
                //}
                //case NUMERIC:{ // 0
                //    break;
                //}
                case STRING: { // 1
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                //case FORMULA: { // 2
                //    break;
                //}
                case BLANK: { // 3
                    cellValue = "";
                    break;
                }
                //case BOOLEAN: { // 4
                //    break;
                //}
                //case ERROR:{ // 5
                //    break;
                //}
                default:
                    cellValue = "";
            }
            if (cell.getColumnIndex() > 3 && checkValue(cellValue)) {
            }

        } else {
            cellValue = "";
        }
        return cellValue.trim();
    }

    public static void writeExcel() {
        XSSFWorkbook xssWorkbook = new XSSFWorkbook();
        XSSFSheet sheet = xssWorkbook.createSheet("sheet1");

        for (int row = 0; row < 10; row++) {
            XSSFRow rows = sheet.createRow(row);
            for (int col = 0; col < 10; col++) {

                rows.createCell(col).setCellValue("data" + row + col);
            }
        }
        //
        File xlsFile = new File("poi.xlsx");
        try {
            FileOutputStream xlsStream = new FileOutputStream(xlsFile);
            xssWorkbook.write(xlsStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void formatJson(String json) {
        System.out.println();
        Map<String, Object> map = new HashMap<String, Object>();
        Map<String, String> mapNew = new HashMap<String, String>();
        map.put("1", 11);
        map.put("@1", 11);
        map.put("2", 22);
        map.put("@2", 22);
        map.put("3", 33);
        map.put("@3", 33);
        map.put("4", 44);
        map.put("@4", 44);
        for (String string : map.keySet()) {
            mapNew.put(string, map.get(string).toString());
        }

        com.alibaba.fastjson.JSONObject jsonObject1 = new com.alibaba.fastjson.JSONObject(map);
        String string = jsonObject1.toString();
        System.out.println(string);

        com.alibaba.fastjson.JSONObject jsonObject3 = JSON.parseObject(json);
        String rs = JSON.toJSONString(jsonObject3, SerializerFeature.MapSortField, SerializerFeature.PrettyFormat);
        System.out.println(rs);
    }

    public static void writeToJson(String json, String filePath) {
        try {
            File fileInfo = new File(filePath);
            FileOutputStream fileOutputStreamInfo = new FileOutputStream(fileInfo);
            fileOutputStreamInfo.write(json.getBytes());
            fileOutputStreamInfo.flush();
            fileOutputStreamInfo.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static boolean checkValue(String value) {


        if (Pattern.matches(".*[\\n]+.*", value)) {
            System.out.println("~~~~ \\n 1 : " + value);
            System.out.println("~~~~ \\n 1 修复: ~\n" + value.replaceAll("[\\n]+", ""));
            System.out.println();
            return true;
        }
        if (Pattern.matches(".*[\\\\]+[ ]+[n]+.*", value)) {
            System.out.println("~~~~ \\n 2 : " + value);
            System.out.println("~~~~ \\n 2 修复: ~\n" + value.replaceAll("[\\\\]+[ ]+[n]+", "\\\\n"));
            System.out.println();
            return true;
        }
        if (value.contains("%") || value.contains("$")) {
            String value_1 = value.replaceAll("%[\\d]+\\$s", "").replaceAll("[\\s]", "");
            if (Pattern.matches(".*[\\%]+[\\d]+[\\$]+[s]+.*", value_1)) {
                System.out.println("~~~~ %n$s : " + value);
                return true;
            }
        }

        if (value.contains(" ")) {
            System.out.println("~~~~ \\n NBSP : " + value);
            System.out.println("~~~~ \\n NBSP 修复: ~\n" + value.replaceAll(" ", " "));
            System.out.println();
            return true;
        }
        return false;
    }

    private static String toJsonStr(JsonObject jsonObject) {
        JsonElement cSetting = jsonObject.get("vi").getAsJsonObject().get("cSetting");
        String s = cSetting.toString();
        System.out.println("toJsonStr s = " + s);
        String jsonStr = new GsonBuilder().setPrettyPrinting().create().toJson(jsonObject)
                .replace("\\\\n", "\\n")
                .replace("\\u0027", "\'");
        return jsonStr;
    }

    private static void pp(String s) {
        System.out.println(s);
    }
}
