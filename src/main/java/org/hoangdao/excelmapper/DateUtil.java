package org.hoangdao.excelmapper;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;

public class DateUtil {

    public static final String VIETNAMESE_DATE_PATTERN = "dd/MM/yyyy";

    public static String getJsonSearchLogDaily() {

        Calendar cal = Calendar.getInstance();
        cal.add(Calendar.MONTH, 1);
        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH);
        int date = cal.get(Calendar.DATE);
        String fromDate = null, toDate = null;
        if (date != 1) {
            fromDate = year + "-" + month + "-" + (date - 2) + "T17:00:00.000Z";
            toDate = year + "-" + month + "-" + (date - 1) + "T16:59:59.000Z";
        } else {
            cal.add(Calendar.MONTH, -1);
            month = cal.get(Calendar.MONTH);
            int maxDateOfMonth = cal.getMaximum(Calendar.DAY_OF_MONTH);
            fromDate = year + "-" + month + "-" + (maxDateOfMonth - 1) + "T17:00:00.000Z";
            toDate = year + "-" + month + "-" + maxDateOfMonth + "T16:59:59.000Z";
        }
        String json = "{\n" +
                " \"fromDate\": \"" + fromDate + "\",\n" +
                " \"toDate\": \"" + toDate + "\"\n" +
                " }";
        return json;
    }

    public static Date toDate(LocalDate localDate) {
        return Date.from(localDate.atStartOfDay(ZoneId.systemDefault()).toInstant());
    }

    public String getDate() {
        Calendar cal = Calendar.getInstance();
        cal.add(Calendar.MONTH, 1);
        int year = cal.get(Calendar.YEAR);
        int month = cal.get(Calendar.MONTH);
        int date = cal.get(Calendar.DATE);
        return date + "-" + month + "-" + year;
    }

    public static String formatDate(Date date, String pattern) {
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
            return simpleDateFormat.format(date);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * If not specify pattern, dd/MM/yyyy would be use
     * @param date
     * @param pattern
     * @return
     */
    public static String format(Date date, String pattern) {
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
            return simpleDateFormat.format(date);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Format date as default pattern
     * @param date
     * @return
     */
    public static String format(Date date) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(VIETNAMESE_DATE_PATTERN);
        return simpleDateFormat.format(date);
    }

    public static String formatDate(Date date) {
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");
            return simpleDateFormat.format(date);
        } catch (Exception e) {
            return null;
        }
    }

    public static Date parseDate(String input){
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");
            return simpleDateFormat.parse(input);
        } catch (Exception e) {
            return null;
        }
    }

    public static Date parseDate(String input, String pattern){
        try {
            DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern(pattern);
            LocalDate localDate = LocalDate.parse(input.trim(), dateTimeFormatter);
            return Date.from(localDate.atStartOfDay(ZoneId.of(ZoneOffset.UTC.getId())).toInstant());
        } catch (Exception e) {
            return null;
        }
    }
}
