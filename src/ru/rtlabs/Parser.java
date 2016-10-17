package ru.rtlabs;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.rtlabs.DB.DBWorker;

import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

public class Parser {

    private Integer snilsCount;
    private Integer patientCount;
    private Integer problemCount;

    public void parse(DBWorker connection, String file){

        snilsCount = 0;
        patientCount = 0;
        problemCount = 0;
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = wb.getSheetAt(0);
            for (int i = 8; i < sheet.getLastRowNum() + 1; i++) {
                XSSFRow row = sheet.getRow(i);
                Patient patient = new Patient();
                PatientSearch search = new  PatientSearch();
                //PatientAdd patientAdd = new PatientAdd();
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    XSSFCell cell = row.getCell(j);
                    switch (j) {
                        case 7:
                            if (cell == null) {
                                continue;
                            } else {
                                switch (cell.getCellType()) {
                                    case XSSFCell.CELL_TYPE_STRING:
                                        if (Formatter.format(cell.getStringCellValue()).length == 3){
                                            patient.setSurname(Formatter.format(cell.getStringCellValue())[0]);
                                            patient.setName(Formatter.format(cell.getStringCellValue())[1]);
                                            patient.setpName(Formatter.format(cell.getStringCellValue())[2]);
                                        }else if(Formatter.format(cell.getStringCellValue()).length == 2) {
                                           patient.setSurname(Formatter.format(cell.getStringCellValue())[0]);
                                            patient.setName(Formatter.format(cell.getStringCellValue())[1]);
                                        }else {
                                            patient.setSurname(Formatter.format(cell.getStringCellValue())[0]);
                                            patient.setName(Formatter.format(cell.getStringCellValue())[1]);
                                            patient.setpName(Formatter.format(cell.getStringCellValue())[2] + " " + Formatter.format(cell.getStringCellValue())[3]);
                                        }
                                        break;
                                }
                            }
                            break;
                        case  8:
                            if (cell == null) {
                                continue;
                            } else {
                                switch (cell.getCellType()) {
                                    case XSSFCell.CELL_TYPE_STRING:
                                        DateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                                        Date parsedB = format.parse(cell.getStringCellValue());
                                        java.sql.Date sqlD = new java.sql.Date(parsedB.getTime());
                                        patient.setBdate(sqlD);
                                        break;
                                    case XSSFCell.CELL_TYPE_NUMERIC:
                                        DateFormat formatqq = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzzz yyyy", Locale.ENGLISH);
                                        Date parseB = formatqq.parse(String.valueOf(cell.getDateCellValue()));
                                        java.sql.Date sqlD2 = new java.sql.Date(parseB.getTime());
                                        patient.setBdate(sqlD2);
                                        break;
                                }
                            }
                            break;
                        case 12:
                            if (cell == null) {
                                continue;
                            } else {
                                switch (cell.getCellType()) {
                                    case XSSFCell.CELL_TYPE_STRING:
                                        patient.setSnils(cell.getStringCellValue());
                                        break;
                                    case XSSFCell.CELL_TYPE_NUMERIC:
                                        patient.setSnils(String.valueOf(cell.getNumericCellValue()));
                                        break;
                                }
                            }
                            break;
                    }
                }
                FileWriter fileWriter2 = new FileWriter("log.txt", true);
                BufferedWriter writer2 = new BufferedWriter(fileWriter2);
                writer2.write(java.util.Calendar.getInstance().getTime() + " Фамилия: " + patient.getSurname() + " Имя: " + patient.getName() + " Отчество: " + patient.getpName() + " День Рождения: " + patient.getBdate());
                writer2.newLine();
                writer2.flush();
                writer2.close();
                System.out.println(java.util.Calendar.getInstance().getTime() + " Фамилия: " + patient.getSurname() + " Имя: " + patient.getName() + " Отчество: " + patient.getpName() + " День Рождения: " + patient.getBdate());
                System.out.println("Поиск пациента по базе данных...");
                search.search(patient.getSurname(), patient.getName(), patient.getpName(), patient.getBdate(), patient.getPolNumber(), connection);
                patientCount++;
                if (search.getCount() > 1){
                    FileWriter fileWriterd = new FileWriter("dubli.txt", true);
                    BufferedWriter writerd = new BufferedWriter(fileWriterd);
                    writerd.write(java.util.Calendar.getInstance().getTime() + " Фамилия: " + patient.getSurname() + " Имя: " + patient.getName() + " Отчество: " + patient.getpName() + " День Рождения: " + patient.getBdate() + " " + patient.getSnils());
                    writerd.newLine();
                    writerd.flush();
                    writerd.close();
                    System.out.println(java.util.Calendar.getInstance().getTime() + " Фамилия: " + patient.getSurname() + " Имя: " + patient.getName() + " Отчество: " + patient.getpName() + " День Рождения: " + patient.getBdate() + " " + patient.getSnils());
                }

                if (search.isHasId() && search.getCount() == 1){
                    if (patient.getSnils() != null){
                        search.hasDoc1(connection);
                        if (search.isHasDoc1()){
                            search.docUpdate1(patient.getSnils(), connection);
                            snilsCount++;
                        }else {
                            search.docInsert1(patient.getSnils(), connection);
                        }
                        System.out.println("--------------------------------------------------------------------------------------------------------------------------------------");
                        FileWriter fileWriter = new FileWriter("log.txt", true);
                        BufferedWriter writer = new BufferedWriter(fileWriter);
                        writer.write("--------------------------------------------------------------------------------------------------------------------------------------");
                        writer.newLine();
                        writer.flush();
                        writer.close();
                    }else {
                        problemCount ++;
                        System.out.println("--------------------------------------------------------------------------------------------------------------------------------------");
                        FileWriter fileWriter = new FileWriter("log.txt", true);
                        BufferedWriter writer = new BufferedWriter(fileWriter);
                        writer.write(java.util.Calendar.getInstance().getTime() + " у  пациента недостаточно данных по документам для вставки в РМИС (отсутствует СНИЛС) " + patient.getSurname() + " " + patient.getName() + " " + patient.getpName() + " День Рождения " + patient.getBdate());
                        writer.newLine();
                        writer.write("--------------------------------------------------------------------------------------------------------------------------------------");
                        writer.newLine();
                        writer.flush();
                        writer.close();
                    }
                }else if(search.getCount() == 0){
                    System.out.println(java.util.Calendar.getInstance().getTime() + " не найден в РМИС " + patient.getSurname() + " " + patient.getName() + " " + patient.getpName() + " День Рождения " + patient.getBdate());
                    System.out.println("--------------------------------------------------------------------------------------------------------------------------------------");
                    FileWriter fileWriter = new FileWriter("log.txt", true);
                    BufferedWriter writer = new BufferedWriter(fileWriter);
                    writer.write(java.util.Calendar.getInstance().getTime() + " не найден в РМИС " + patient.getSurname() + " " + patient.getName() + " " + patient.getpName() + " День Рождения " + patient.getBdate());
                    writer.newLine();
                    writer.write("--------------------------------------------------------------------------------------------------------------------------------------");
                    writer.newLine();
                    writer.flush();
                    writer.close();
                }

            }
            }catch (IOException | ParseException e) {
            e.printStackTrace();
        }
        try {
            System.out.println("Кол-во обработанных пациентов: " + this.patientCount);
            System.out.println("Кол-во пациентов, у которых нет в файле СНИЛСа: " + this.problemCount);
            System.out.println("Кол-во паценитов, которым был обновлен СНИЛС: " + this.snilsCount);
            System.out.println("Кол-во пациентов, которым был добавлен СНИЛС: " + (this.patientCount - this.snilsCount));
            FileWriter fileWriter = new FileWriter("stats.txt", true);
            BufferedWriter writer = new BufferedWriter(fileWriter);
            writer.write("Кол-во обработанных пациентов: " + this.patientCount);
            writer.newLine();
            writer.write("Кол-во пациентов, у которых нет в файле СНИЛСа: " + this.problemCount);
            writer.newLine();
            writer.write("Кол-во пациентов, которым был обновлен СНИЛС: " + this.snilsCount);
            writer.newLine();
            writer.write("Кол-во пациентов, которым был добавлен СНИЛС: " + (this.patientCount - this.snilsCount));
            writer.newLine();
            writer.flush();
            writer.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    }