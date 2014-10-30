package com.gmail.zigfridus;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class MainClass {

    public static void main(String[] args) throws Exception {
        //Get name of file with users
        String fileName = "";
        if (args.length > 0) {
            fileName = args[0];
            System.out.println("[INFO]Parse file with users " + fileName);
        } else {
            fileName = "import.xls";
            System.out.println("[INFO]There aren't any arguments. Parse file import.xls by default");
        }
        try {
            FileInputStream file = new FileInputStream(fileName);
            HSSFWorkbook wb = new HSSFWorkbook(file);
            HSSFSheet sheet = wb.getSheetAt(0);
            Boolean fl = true;
            int i = 4;
            int sum = 0;

            //get root home dir from xls
            HSSFRow row = sheet.getRow(0);
            HSSFCell cell = row.getCell(1);
            String rootDir = "-";
            //check for right path
            //if value in cell is "-" then don't create user's home dir
            try {
                if (cell.getRichStringCellValue().getString() != "-") {
                    if (cell.getRichStringCellValue().getString().endsWith("\\")) {
                        rootDir = cell.getRichStringCellValue().getString();
                    } else {
                        rootDir = cell.getRichStringCellValue().getString() + "\\";
                    }
                }
            } catch (NullPointerException e) {
                fl = false;
                System.out.println("[ERROR]Wrong value in the B1 cell");
            }

            //get admins group name
            row = sheet.getRow(1);
            cell = row.getCell(1);
            String adminGrp = "-";
            try {
                if (cell.getRichStringCellValue().getString() != "-") {
                    adminGrp = cell.getRichStringCellValue().getString();
                }
            } catch (NullPointerException e) {
                fl = false;
                System.out.println("[ERROR]Wrong value in the B2 cell");
            }

            // parse table until meet null value (empty cell)
            while (fl) {
                row = sheet.getRow(i);
                try {
                    Boolean fl2 = true;
                    cell = row.getCell(0);
                    String login = cell.getRichStringCellValue().getString();
                    cell = row.getCell(1);
                    String password = cell.getRichStringCellValue().getString();
                    cell = row.getCell(2);
                    String fullName = cell.getRichStringCellValue().getString();

                    cell = row.getCell(3);
                    String active = cell.getRichStringCellValue().getString();
                    //check active value
                    if (!(active.equalsIgnoreCase("yes")) && !(active.equalsIgnoreCase("no"))) {
                        fl2 = false;
                        System.out.println("[ERROR]Row " + (i + 1) + " skipped because of wrong value in the D" + (i + 1) + " cell");
                    }

                    cell = row.getCell(4);
                    String accountExpires = cell.getRichStringCellValue().getString();

                    cell = row.getCell(5);
                    String passChange = cell.getRichStringCellValue().getString();
                    //check passChange value
                    if (!(passChange.equalsIgnoreCase("yes")) && !(passChange.equalsIgnoreCase("no"))) {
                        fl2 = false;
                        System.out.println("[ERROR]Row " + (i + 1) + " skipped because of wrong value in the F" + (i + 1) + " cell");
                    }

                    cell = row.getCell(6);
                    String passReq = cell.getRichStringCellValue().getString();
                    //check passReq value
                    if (!(passReq.equalsIgnoreCase("yes")) && !(passReq.equalsIgnoreCase("no"))) {
                        fl2 = false;
                        System.out.println("[ERROR]Row " + (i + 1) + " skipped because of wrong value in the G" + (i + 1) + " cell");
                    }

                    cell = row.getCell(7);
                    String passExpire = cell.getRichStringCellValue().getString();
                    //check passExpire value
                    if (!(passExpire.equalsIgnoreCase("false")) && !(passExpire.equalsIgnoreCase("true"))) {
                        fl2 = false;
                        System.out.println("[ERROR]Row " + (i + 1) + " skipped because of wrong value in the H" + (i + 1) + " cell");
                    }

                    //if there are no errors in current row
                    if (fl2) {
                        //add user
                        String command1 = "net user " + login + " " + password + " /add /fullname:\"" + fullName + "\" /active:" + active + " /expires:" + accountExpires + " /passwordchg:" + passChange + " /passwordreq:" + passReq;
                        Process p = Runtime.getRuntime().exec(command1);
                        p.waitFor();
                        //Set user's password expires value
                        String command2 = "wmic path Win32_Useraccount where Name='" + login + "' set passwordexpires=" + passExpire + " /nointeractive";
                        p = Runtime.getRuntime().exec(command2);
                        p.waitFor();

                        //create home dir
                        //root home folder should be created
                        if (rootDir != "-") {
                            //create users dir with parents
                            new File(rootDir + login).mkdirs();
                            //remove inheritance from root dir
                            String command3 = "icacls " + rootDir + login + " /inheritance:r";
                            p = Runtime.getRuntime().exec(command3);
                            p.waitFor();
                            //add access for admins group
                            String command4 = "icacls " + rootDir + login + " /grant " + adminGrp + ":(F)";
                            p = Runtime.getRuntime().exec(command4);
                            p.waitFor();
                            //add access for user
                            String command5 = "icacls " + rootDir + login + " /grant " + login + ":(M)";
                            p = Runtime.getRuntime().exec(command5);
                            p.waitFor();
                        }

                        System.out.println("[INFO]User " + login + " added");
                        sum++;
                    }
                    i++;
                } catch (NullPointerException ex) {
                    fl = false;
                }
            }
            System.out.println("[INFO]" + Integer.toString(sum) + " user(s) added");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
