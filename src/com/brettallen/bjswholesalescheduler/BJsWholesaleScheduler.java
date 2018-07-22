package com.brettallen.bjswholesalescheduler;

import java.awt.*;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;

import javax.swing.text.BadLocationException;

import com.brettallen.bjswholesalescheduler.excel.ExcelWriter;
import com.brettallen.bjswholesalescheduler.gui.Menu;
import com.brettallen.bjswholesalescheduler.utils.CsvConverter;
import com.brettallen.bjswholesalescheduler.utils.Shift;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;

/*
 * Created By: Brett Allen on Oct 22, 2015 at 1:13:31 PM
 *
 *      The main objective of this program is to save as much time as possible by delegating the schedule-writing
 *      process to an information system. What used to take about 30 minutes per day to do by hand now takes approximately 3-5 minutes
 *      to do for an entire weeks worth of schedule-writing in one sitting.
 *
 *      This program takes a schedule of employees in as input in the form of an excel document, parses the contents
 *      of the excel document to a comma separated version (.csv), and creates a new chronologically sorted
 *      schedule.
 *
 *      Through a graphical user interface (GUI), the chronologically sorted schedule awaits manipulation from the
 *      end user to either process schedules for each day, search for
 *      employees to see their shift for each consecutive day, or email employees their schedules for the week.
 */

public class BJsWholesaleScheduler
{
    //A backup method of obtaining the list of employees that wish to receive their schedules via email
    public static final InputStream BACKUP_EMAIL_FILE = BJsWholesaleScheduler.class.getResourceAsStream("res/emails.txt");

    //Front-line schedule template
    private static final String FRONT_LINE_TEMPLATE = "res/frontLineTemplate03.xls";
    private static final String CASHIER_TEMPLATE = "res/cashierTemplate.xls";
    private static final String OVERNIGHT_TEMPLATE = "res/overnightTemplate.xls";

    private static final String DATES = "Sun,Mon,Tue,Wed,Thu,Fri,Sat";
    private static final String[] SPECIAL_SHIFTS =
            {
                    "Open",
                    "Early",
                    "Close",
                    "Mid",
                    "Holiday",
                    "Unpaid Day Off",
                    "Vac",
                    "Personal",
                    "Birthday",
                    "Star"
            };

    public enum ScheduleTypes {CASHIER, FRONT_LINE, OVERNIGHT}

    public static final int DAYS = 7;

    public static void main(String[] args) throws IOException, BiffException,
            WriteException, InterruptedException, BadLocationException
    {
        Menu menu = new Menu();
        menu.setVisible(true);
        menu.setResizable(false);

        //Set the value of the public menu variable in the Menu class to the value of the menu variable here
        menu.menu = menu;
    }//End main

    //This method embodies all the central processing of this program. It acts as an overloaded main method
    //and will be called only by the Menu class
    @SuppressWarnings("ConstantConditions")
    public static void mainDelegate(Menu menu) throws BadLocationException, BiffException
    {
        //Initialize local variables
        //Creating a CsvConverter variable that will be instantiated once the initial
        //xls file is edited programmatically
        CsvConverter converter = null;

        Shift[][] schedule;
        Shift shift1;
        Shift shift2;
        //Shift shift3;

        ArrayList<Shift> multiShifts = new ArrayList<Shift>();
        ArrayList<String> truncDates = new ArrayList<String>();

        boolean multiplier;

        //Used to determine if the employee contains SPECIAL_SHIFTS
        boolean isSpecial;

        String employee,
                shiftPosition,
                shiftPosition2 = "",
                //shiftPosition3 = "",
                shiftTime,
                shiftTime2 = "",
                //shiftTime3  = "",
                specialShiftTime = "",
                lineA,
                lineB,
                lines;

        String fileName = menu.getXlsPath();

        String[] dates;

        int startCol = 0,
                numEmployees = 0,
                dayNum,
                e = -1,
                iOffset,
                count = -1;

        ArrayList<String> chosenFile = new ArrayList<String>();

        menu.print("Initialization complete.\n\n");
        //End initialization

        //Try opening chosen file
        try
        {
            //Create a new ExcelWriter that will receive an excel file as input and eventually parse to .csv
            ExcelWriter writer = new ExcelWriter(fileName);

            //Deleting column A
            writer.deleteFirstColumn();
            writer.writeAndClose();

            try
            {
                converter = new CsvConverter();

                converter.convertExcelToCSV(writer.getXlsFilePath(), new File(".").getAbsolutePath());

                //Delete input.xls
                //noinspection ResultOfMethodCallIgnored
                writer.getXlsFile().delete();
            }
            catch(Exception ex)
            {
                System.err.println("Caught an: " + ex.getClass().getName());
                System.err.println("Message: " + ex.getMessage());
                System.err.println("Stacktrace follows:.....");
                ex.printStackTrace(System.out);
            }

            //Reading text files in the default encoding.
            FileReader fileReader = new FileReader(converter.getCsvFilePath());

            // Always wrap FileReader in BufferedReader.
            BufferedReader bufferedReader = new BufferedReader(fileReader);

            while((lines = bufferedReader.readLine()) != null)
            {
                //System.out.println(lines);
                chosenFile.add(lines);
            }

            // Always close files.
            bufferedReader.close();
        }
        catch(FileNotFoundException ex)
        {
            //errorProcessing = true;
            System.err.println("Unable to open file '" + fileName + "'");
            menu.print("Unable to open file '" + fileName + "'\n");
        }
        catch(IOException ex)
        {
            //errorProcessing = true;
            System.err.println("Error reading file '" + fileName + "'");
            menu.print("Error reading file '" + fileName + "'\n");
        }

        //Delete input.csv
        //noinspection ResultOfMethodCallIgnored
        converter.getCsvFile().delete();

        menu.print("File Opened.\n\n");
        //End open file

        //Parse the contents from the csv file to a string array representing the data in the csv file
        String[] csv = chosenFile.toArray(new String[chosenFile.size()]);

        // Search where the dates column starts for data parsing
        for (String line : csv)
        {
            if (line.contains(DATES))
            {
                String[] tokens = line.split(",");

                startCol = indexOfSunday(tokens);

                if (startCol == -1)
                {
                    System.err.println("Sunday not found!");
                    menu.print("Sunday not found!\n");
                    System.exit(-1);
                }
            }

            if (line.startsWith("\""))
                numEmployees++;
        }
        // End search

        dates = csv[startCol + 3].split(",");

        schedule = new Shift[numEmployees][DAYS];

        // Parsing data
        for (int i = 0; i < csv.length; i++)
        {
            // Line that contains position for each day
            lineA = csv[i];

            // Line that starts with an employee's name (names are surrounded in "")
            if (lineA.startsWith("\""))
            {
                // Current row in schedule
                e++;

                // Line that contains start and end times
                lineB = csv[i + 1];

                // Gets employee name
                employee = lineA.substring(1, lineA.lastIndexOf("\""));

                for (int j = startCol; j < startCol + 7; j++)
                {
                    multiplier = false;
                    isSpecial = false;

                    dayNum = j - 2;

                    count++;

                    if(count < dates.length)
                        if (startsWithNumber(dates[count]))
                            truncDates.add(dates[count]);

                    shiftPosition = lineA.split(",")[j + 1];

                    if (shiftPosition.equals("."))
                    {
                        //If the cell directly below the initially set shiftPosition is also blank
                        //skip that employee
                        if(csv[i + 1].split(",")[j].equals("."))
                            continue;

                        shiftPosition = csv[i + 1].split(",")[j];
                    }
                    else
                    {
                        //Determine if special position (specifically if manager)
                        if(shiftPosition.contains("Manager") || shiftPosition.contains("Mgr"))
                            isSpecial = true;

                        specialShiftTime = csv[i + 1].split(",")[j];

                        //Determine if the employee's shift time contains a special shift
                        for (String specialShift : SPECIAL_SHIFTS)
                        {
                            if (specialShiftTime.contains(specialShift))
                                isSpecial = true;
                        }
                    }

                    shiftTime = lineB.split(",")[j];

                    //Skips start and end times (vacations, unpaid days off) if both
                    //the initial field and field below are invalid
                    if (!startsWithNumber(shiftTime) && !isSpecial)
                    {
                        //Determine if the shiftTime is offset
                        iOffset = i + 1;
                        lineB = csv[iOffset + 1];

                        shiftTime = lineB.split(",")[j];

                        //If the shiftTime is still not a number skip that employee
                        if(!startsWithNumber(shiftTime))
                            continue;
                    }

                    //Create new shift
                    try
                    {
                        //If the employee has a special shift (taking vacation, unpaid day off, a manager, etc.)
                        //Make that employee's shiftTime equal to their special shift time
                        if(isSpecial)
                        {
                            shift1 = new Shift(employee, shiftPosition, specialShiftTime, "", dayNum, isSpecial);
                        }
                        else
                        {
                            shift1 = new Shift(employee, shiftPosition, shiftTime.split("-")[0], shiftTime.split("-")[1],
                                    dayNum, isSpecial);
                        }
                    }
                    catch (ArrayIndexOutOfBoundsException e1)
                    {
                        break;
                    }

                    //Assign new shift1
                    schedule[e][j - startCol] = shift1;

                    //Determine if the there is a second portion of the current employee's shift
                    try
                    {
                        String[] split1 = csv[i + 3].split(",");
                        String[] split2 = csv[i + 4].split(",");
                        //Determine plausible second shift
                        if(j < split1.length && j < split2.length){
                            if(!split1[j].equals(".") &&
                                    !startsWithNumber(split1[j]) &&
                                    !csv[i + 3].startsWith("\""))
                            {
                                multiplier = true;

                                shiftPosition2 = split1[j];
                                shiftTime2 = split2[j];

                            /*if(!csv[i + 6].split(",")[j].equals(".") &&
                                    !startsWithNumber(csv[i + 6].split(",")[j]) &&
                                    !csv[i + 6].startsWith("\""))
                            {
                                //Get the third shift position if it exists
                                shiftPosition3 = csv[i + 6].split(",")[j];

                                //Get the third shift time is it exists
                                shiftTime3 = csv[i + 7].split(",")[j];
                            }*/
                            }
                            else if(!split2[j].equals(".") &&
                                    !startsWithNumber(split2[j]) &&
                                    !csv[i + 4].startsWith("\""))
                            {
                                multiplier = true;

                                shiftPosition2 = split2[j];
                                shiftTime2 = csv[i + 5].split(",")[j];
                            }
                        }

                        //Determine if there was a second shift
                        if(multiplier && !shiftTime2.equals("."))
                        {
                            try
                            {
                                shift2 = new Shift(employee, shiftPosition2, shiftTime2.split("-")[0],
                                        shiftTime2.split("-")[1],
                                        dayNum, isSpecial);

                                shift2.setSeparator(" | ");

                                multiShifts.add(shift2);

                                //Add a third shift if one exists
                                /*if(!shiftTime3.equals("."))
                                {
                                    shift3 = new Shift(employee, shiftPosition3, shiftTime3.split("-")[0],
                                            shiftTime3.split("-")[1], dayNum, isSpecial);

                                    shift3.setSeparator(" | ");

                                    multiShifts.add(shift3);
                                }*/
                            }
                            catch (ArrayIndexOutOfBoundsException e1)
                            {
                                e1.printStackTrace();
                            }
                        }
                    }
                    catch(ArrayIndexOutOfBoundsException e1){
                        e1.printStackTrace();
                    }
                }
            }
        }
        //End parsing

        //Combine original schedule with multiShifts. Overwrite null elements of original schedule
        combine(schedule, numEmployees, multiShifts);

        //Remove all duplicate shifts from schedule array
        removeDuplicates(schedule, numEmployees);

        menu.print("Processing complete.\n");

        //Transfer specified variables to the Menu class for rendering purposes
        menu.transferVariables(schedule, numEmployees, truncDates);

        //Change the text of the buttons that indicate the days to the actual dates
        menu.setDayButtonTexts(truncDates);

    }//End mainDelegate

    private static void combine(Shift[][] myArray1, int rows, ArrayList<Shift> myArray2)
    {
        int i = 0;

        System.out.println("Combining multi-shifts...");

        //Traverse myArray1 and replace all null elements with an element of myArray2
        for(int x = 0; x < rows; x++)
        {
            for(int y = 0; y < DAYS; y++)
            {
                if(myArray1[x][y] == null)
                {
                    if(i < myArray2.size())
                    {
                        myArray1[x][y] = myArray2.get(i);
                        i++;
                    }
                }
            }
        }

        System.out.println("Done.");
    }

    //Traverse schedule and eraseEmail all duplicates
    private static void removeDuplicates(Shift[][] myArray, int rows)
    {
        System.out.println("Removing duplicates...");

        //Using nested loops to compare the schedule to itself (2D array compared with 2D array)
        for(int a = 0; a < rows; a++)
        {
            for(int b = 0; b < DAYS; b++)
            {
                for(int c = 0; c < rows; c++)
                {
                    for(int d = 0; d < DAYS; d++)
                    {
                        //Determine if the shift that we are comparing to is the same element at the same location
                        if ((a != c) && (b != d))
                        {
                            if (myArray[c][d] != null && myArray[a][b] != null)
                            {
                                if (myArray[c][d].day.equals(myArray[a][b].day)
                                        && myArray[c][d].position.contains(myArray[a][b].position)
                                        && myArray[c][d].employee.contains(myArray[a][b].employee))
                                {
                                    myArray[c][d] = null;
                                }
                            }
                        }
                    }
                }
            }
        }

        System.out.println("Done.");
    }

    //Chronologically sorting schedule
    //Sorting algorithm: Ascending (least to greatest)(earliest to latest)
    public static void sortAscending(Shift[][] myArray, int rows, Menu menu) throws BadLocationException
    {
        Shift temp;

        System.out.println("Sorting in ascending order...");

        //numEmployees represents rows, days represents columns
        for(int a = 0; a < rows; a++)
        {
            for(int b = 0; b < DAYS; b++)
            {
                for(int c = 0; c < rows; c++)
                {
                    for(int d = 0; d < DAYS; d++)
                    {
						/*If schedule is not null and military time of the initial element is
						greater than military time of current element then swap the current element
						and	initial element*/
                        if(myArray[c][d] != null && myArray[a][b] != null)
                        {
                            //Don't attempt to sort the employees with special shifts
                            if(!myArray[c][d].isSpecial() && !myArray[a][b].isSpecial())
                            {
                                if(getMilitaryTime(myArray[c][d].startTime) > getMilitaryTime(myArray[a][b].startTime))
                                {
                                    temp = myArray[a][b];
                                    myArray[a][b] = myArray[c][d];
                                    myArray[c][d] = temp;
                                }
                            }
                        }
                    }
                }
            }
        }

        System.out.println("Done.");

        if(menu != null){
            menu.print("\nSorted in ascending order.\n");
        }

    }

    //Chronologically sorting schedule
    //Sorting algorithm: Ascending (least to greatest)(earliest to latest)
    public static void sortDescending(Shift[][] myArray, int rows, Menu menu) throws BadLocationException
    {
        Shift temp;

        System.out.println("Sorting in descending order...");

        //numEmployees represents rows, days represents columns
        for(int a = 0; a < rows; a++)
        {
            for(int b = 0; b < DAYS; b++)
            {
                for(int c = 0; c < rows; c++)
                {
                    for(int d = 0; d < DAYS; d++)
                    {
						/*If schedule is not null and military time of the initial element is
						greater than military time of current element then swap the current element
						and	initial element*/
                        if(myArray[c][d] != null && myArray[a][b] != null)
                        {
                            //Don't attempt to sort the employees with special shifts
                            if(!myArray[c][d].isSpecial() && !myArray[a][b].isSpecial())
                            {
                                if(getMilitaryTime(myArray[c][d].startTime) < getMilitaryTime(myArray[a][b].startTime))
                                {
                                    temp = myArray[a][b];
                                    myArray[a][b] = myArray[c][d];
                                    myArray[c][d] = temp;
                                }
                            }
                        }
                    }
                }
            }
        }

        System.out.println("Done.");
        menu.print("\nSorted in descending order.\n");
    }

    public static void renderChoice(String day, Shift[][] myArray, int rows, int cols,
                                    int choice,  ArrayList<String> truncDates, Menu menu) throws IOException, BiffException,
            WriteException, BadLocationException
    {
        //Set cursor to loading symbol
        menu.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        String[] tokens;
        String date;
        String year = Integer.toString(Calendar.getInstance().get(Calendar.YEAR));

        tokens = truncDates.get(choice - 1).split("-");
        date = transformDate(tokens, year);

        //Determine which schedule(s) to print based on checkbox state(s) of menu
        //Print frontLineSchedule if frontLineCbx is checked
        if(menu.getCbxFrontLineState())
        {
            menu.print("\nRendering choice...\n");

            //Load in the Excel template
            InputStream frontLineTemplate = BJsWholesaleScheduler.class.getResourceAsStream(FRONT_LINE_TEMPLATE);

            //Create Excel writer for Front Line shifts
            ExcelWriter frontLineSchedule = new ExcelWriter(frontLineTemplate, ScheduleTypes.FRONT_LINE, day, date);

            //Write the day to cell B1
            frontLineSchedule.overwriteEmptyCell(1, 0, day, 0);

            //Write the date to cell N1
            frontLineSchedule.overwriteEmptyCell(13, 0, date, 0);

            displayFrontLineSchedule(day, myArray, rows, cols, frontLineSchedule);

            menu.print("\nFront Line schedule for " + day + ", " + date + " has been created.\n");

            frontLineSchedule.writeAndClose();
        }

        //Print cashierSchedule if cashierCbx is checked
        if(menu.getCbxCashierState())
        {
            menu.print("\nRendering choice...\n");

            //Load in the Excel template
            InputStream cashierTemplate = BJsWholesaleScheduler.class.getResourceAsStream(CASHIER_TEMPLATE);

            //Create Excel writer for Cashier shifts
            ExcelWriter cashierSchedule = new ExcelWriter(cashierTemplate, ScheduleTypes.CASHIER, day, date);

            //Merge col[0-2] and row[0] then write the day and date to cell A1
            cashierSchedule.sheet1.mergeCells(0, 0, 2, 0);
            cashierSchedule.overwriteEmptyCell(0, 0, day + ", " + date, 0);

            displayCashierSchedule(day, myArray, rows, cols, cashierSchedule);

            menu.print("\nCashier schedule for " + day + ", " + date + " has been created.\n");

            cashierSchedule.writeAndClose();
        }

        //Print overnightSchedule if overnightCbx is checked
        if(menu.getCbxOvernightState())
        {
            menu.print("\nRendering choice...\n");

            //Load in the Excel template
            InputStream overnightTemplate = BJsWholesaleScheduler.class.getResourceAsStream(OVERNIGHT_TEMPLATE);

            //Create Excel writer for overnight shifts
            ExcelWriter overnightSchedule = new ExcelWriter(overnightTemplate, ScheduleTypes.OVERNIGHT, day, date);

            //Write the day to cell B1
            overnightSchedule.overwriteEmptyCell(1, 0, day, 0);

            //Write the date to cell N1
            overnightSchedule.overwriteEmptyCell(13, 0, date, 0);

            displayOvernightSchedule(day, myArray, rows, cols, overnightSchedule);

            menu.print("\nOvernight schedule for " + day + ", " + date + " has been created.\n");

            overnightSchedule.writeAndClose();
        }

        menu.setCursor(null);
    }

    public static String transformDate(String[] tokens, String year)
    {
        String month = "";
        String dayNum;

        try
        {
            //Switch implementation for compliance level 1.6
            if(tokens[1].contains("Jan"))
            {
                month = "January";
            }
            else if(tokens[1].contains("Feb"))
            {
                month = "February";
            }
            else if(tokens[1].contains("Mar"))
            {
                month = "March";
            }
            else if(tokens[1].contains("Apr"))
            {
                month = "April";
            }
            else if(tokens[1].contains("May"))
            {
                month = "May";
            }
            else if(tokens[1].contains("Jun"))
            {
                month = "June";
            }
            else if(tokens[1].contains("Jul"))
            {
                month = "July";
            }
            else if(tokens[1].contains("Aug"))
            {
                month = "August";
            }
            else if(tokens[1].contains("Sep"))
            {
                month = "September";
            }
            else if(tokens[1].contains("Oct"))
            {
                month = "October";
            }
            else if(tokens[1].contains("Nov"))
            {
                month = "November";
            }
            else if(tokens[1].contains("Dec"))
            {
                month = "December";
            }

            dayNum = tokens[0];
        }
        catch(IndexOutOfBoundsException e)
        {
            System.err.println("Arg passed into transformDate(String[] tokens) was too small");
            return "";
        }

        return month + " " + dayNum + ", " + year;
    }

    //Method that calls the displayShifts method using the positions included in wall schedule
    private static void displayFrontLineSchedule(String day, Shift[][] myArray, int rows,
                                                 int cols, ExcelWriter frontLineSchedule) throws WriteException
    {
        displayFrontLineShifts("Supv", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Selfcheck Attendant, Scan & Pan, Self Service Attend", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Member Services", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Detective", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Front Door", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Stock/Cart Retriever", day, myArray, rows, cols, frontLineSchedule);

        //Special case with recovery
        //Need to create a temporary partition of the schedule that includes a conglomeration of the shifts
        //That qualify as recovery shifts and then sort them so that they can be displayed in proper order
        //In a consolidated Recovery section of the Frontline schedule
        Shift[][] recoveryPartition = new Shift[rows][cols];

        for(int x = 0; x < rows; x++) {
            for (int y = 0; y < cols; y++) {
               if(myArray[x][y] != null){
                   boolean validRecoveryPos = myArray[x][y].position.contains("Stock Clerk") ||
                           myArray[x][y].position.contains("Ticketer") ||
                           myArray[x][y].position.contains("Recovery");

                   if(myArray[x][y].day.equals(day) && validRecoveryPos && getMilitaryTime(myArray[x][y].endTime) > 900) {
                       Shift moddedShift = myArray[x][y];

                       //Mod a copy of the shift to be recovery so that we only have to search for recovery in the next step
                       moddedShift.position = "Recovery";
                       recoveryPartition[x][y] = moddedShift;
                   }
               }
            }
        }

        //Sort qualifying recovery shifts
        try {
            sortAscending(recoveryPartition, rows, null);
        } catch (BadLocationException e) {
            e.printStackTrace();
        }

        displayFrontLineShifts("Recovery", day, recoveryPartition, rows, cols, frontLineSchedule);

        //End special case

        displayFrontLineShifts("Tire", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Maintenance", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Deli", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Baker", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Office", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Meat", day, myArray, rows, cols, frontLineSchedule);

        displayFrontLineShifts("Produce", day, myArray, rows, cols, frontLineSchedule);
    }

    private static void displayCashierSchedule(String day, Shift[][] myArray, int rows, int cols,
                                               ExcelWriter cashierSchedule) throws WriteException
    {
        displayCashierShifts(day, myArray, rows, cols, cashierSchedule);
    }

    private static void displayOvernightSchedule(String day, Shift[][] myArray, int rows,
                                                 int cols, ExcelWriter overnightSchedule) throws WriteException
    {
        displayOvernightShifts("Specialist", day, myArray, rows, cols, overnightSchedule);
        displayOvernightShifts("Ticketer", day, myArray, rows, cols, overnightSchedule);
        displayOvernightShifts("Baker", day, myArray, rows, cols, overnightSchedule);
        displayOvernightShifts("Receiving", day, myArray, rows, cols, overnightSchedule);
        displayOvernightShifts("Stock Clerk", day, myArray, rows, cols, overnightSchedule);
        displayOvernightShifts("Forklift", day, myArray, rows, cols, overnightSchedule);
    }

    private static void displayOvernightShifts(String position, String day, Shift[][] myArray, int rows,
                                               int cols, ExcelWriter overnightSchedule) throws WriteException
    {
        int excelCol = 0,
                excelRow = 0,
                initRow;

        for(int x = 0; x < rows; x++)
        {
            for(int y = 0; y < cols; y++)
            {
                //Display only existing shifts that correspond to the correct day
                if(myArray[x][y] != null && myArray[x][y].day.equals(day)
                        && myArray[x][y].position.contains(position))
                {
                    //Display shifts that have starting times from 10:00p up to 6:45a (really 6:59a)
                    if(getMilitaryTime(myArray[x][y].startTime) >= 2200 || getMilitaryTime(myArray[x][y].startTime) < 700)
                    {
                        if(myArray[x][y].position.contains("Specialist"))
                        {
                            //Start at cell A5
                            excelCol = 0;
                            excelRow = 4;
                        }
                        else if(myArray[x][y].position.contains("Ticketer"))
                        {
                            //Start at cell D5
                            excelCol = 3;
                            excelRow = 4;
                        }
                        else if(myArray[x][y].position.contains("Baker"))
                        {
                            //Start at cell G5
                            excelCol = 6;
                            excelRow = 4;
                        }
                        else if(myArray[x][y].position.contains("Receiving"))
                        {
                            //Start at cell J5
                            excelCol = 9;
                            excelRow = 4;
                        }
                        else if(myArray[x][y].position.contains("Stock Clerk"))
                        {
                            //Start at cell M5
                            excelCol = 12;
                            excelRow = 4;
                        }
                        else if(myArray[x][y].position.contains("Forklift"))
                        {
                            //Start at cell A13
                            excelCol = 0;
                            excelRow = 12;
                        }

                        initRow = excelRow;

                        while(excelRow < initRow + 6) //Originally initRow + 6
                        {
                            if(overnightSchedule.isEmptyCell(excelCol, excelRow))
                            {
                                //Write the employees name to the ExcelWriter
                                overnightSchedule.overwriteEmptyCell(excelCol, excelRow, myArray[x][y].getName(), 1);

                                //Write that employees shift time next to the employee's name
                                overnightSchedule.overwriteEmptyCell(excelCol + 1, excelRow, myArray[x][y].getShiftTime(), 2);

                                break;
                            }

                            excelRow++;
                        }
                    }
                }
            }
        }
    }

    @SuppressWarnings("ConstantConditions")
    private static void displayFrontLineShifts(String position, String day, Shift[][] myArray, int rows,
                                               int cols, ExcelWriter frontLineSchedule) throws WriteException
    {
        final int offset = 6; //6

        int excelCol = 0,
                excelRow = 0,
                initRow;

        for(int x = 0; x < rows; x++)
        {
            for(int y = 0; y < cols; y++)
            {
                //Only display the shifts that share the correct day, position, and are not null and
                //the end time of the shift is past 9:00am
                if(myArray[x][y] != null && myArray[x][y].day.equals(day)
                        && (position.contains(myArray[x][y].position) || myArray[x][y].position.contains(position))
                        && getMilitaryTime(myArray[x][y].endTime) > 900)
                {
                    boolean recovery = myArray[x][y].position.contains("Recovery") || myArray[x][y].position.contains("Ticketer") ||
                            myArray[x][y].position.contains("Stock Clerk");

                    //Implementing switch logic for jre compliance level 1.6
                    if(myArray[x][y].position.contains("Supv"))
                    {
                        //Do not display inv supv
                        if(myArray[x][y].position.contains("Inv") || myArray[x][y].position.contains("Receiving"))
                            continue;

                        //Start at cell A5
                        excelCol = 0;
                        excelRow = 4;
                    }
                    else if(myArray[x][y].position.contains("Selfcheck Attendant") || myArray[x][y].position.contains("Scan & Pan")
                    		|| myArray[x][y].position.contains("Self Service Attend"))
                    {	
                    	//Start at cell A13
                        excelCol = 0;
                        excelRow = 12; //12
                    }
                    else if(myArray[x][y].position.contains("Member Services"))
                    {
                        //Start at cell A21
                        excelCol = 0;
                        excelRow = 20; //20
                    }
                    else if(myArray[x][y].position.contains("Front Door") || myArray[x][y].position.contains("Detective"))
                    {
                        //Start at cell A29
                        excelCol = 0;
                        excelRow = 28; //28
                    }
                    else if(myArray[x][y].position.contains("Stock/Cart Retriever"))
                    {
                        //Start at cell A37
                        excelCol = 0;
                        excelRow = 36; //36
                    }
                    else if(recovery)
                    {
                        //Start at cell G29
                        excelCol = 6;
                        excelRow = 28; //28
                    }
                    else if(myArray[x][y].position.contains("Tire"))
                    {
                        //Start at cell G13
                        excelCol = 6;
                        excelRow = 12; //20
                    }
                    else if(myArray[x][y].position.contains("Maintenance"))
                    {
                        //Start at cell G5
                        excelCol = 6;
                        excelRow = 4;
                    }
                    else if(myArray[x][y].position.contains("Deli"))
                    {
                        //Start at cell M29
                        excelCol = 12; //6
                        excelRow = 28; //36
                    }
                    else if(myArray[x][y].position.contains("Baker"))
                    {
                        //Start at cell M5
                        excelCol = 12;
                        excelRow = 4;
                    }
                    else if(myArray[x][y].position.contains("Office"))
                    {
                        //Start at cell M13
                        excelCol = 12;
                        excelRow = 12; //12
                    }
                    else if(myArray[x][y].position.contains("Meat"))
                    {
                        //Start at cell M21
                        excelCol = 12;
                        excelRow = 20; //20
                    }
                    else if(myArray[x][y].position.contains("Produce"))
                    {
                        //Start at cell G21
                        excelCol = 6; //12
                        excelRow = 20; //28
                    }
                    //End switch

                    initRow = excelRow;
                    
                    String name = myArray[x][y].getName();
                	
                	// Make employee special if the employee falls into a category with various shifts, i.e., Scan & Pan
                	if(myArray[x][y].position.contains("Scan & Pan")) {
                		name = "* " + name;
                	}

                    //Traverse excel template and fill name and time only if there is not already a name and time there
                    while(excelRow < initRow + offset || recovery) //Originally initRow + 6
                    {	
                        if(frontLineSchedule.isEmptyCell(excelCol, excelRow))
                        {	
                            //Write the employees name to the ExcelWriter
                            frontLineSchedule.overwriteEmptyCell(excelCol, excelRow, name, 1);

                            //Write that employees shift time next to the employee's name
                            frontLineSchedule.overwriteEmptyCell(excelCol + 1, excelRow, myArray[x][y].getShiftTime(), 2);

                            //Make boarders for breaks and lunch cells
                            frontLineSchedule.overwriteEmptyCell(excelCol + 2, excelRow, "", 1);
                            frontLineSchedule.overwriteEmptyCell(excelCol + 3, excelRow, "", 1);
                            frontLineSchedule.overwriteEmptyCell(excelCol + 4, excelRow, "", 1);
                            break;
                        }

                        excelRow++;
                    }

                    //If the current excelRow is greater than the initial value of excelRow plus six, insert a row and place
                    //the employee in the newly inserted row
                    if(excelRow >= initRow + offset && !recovery) //Originally initRow + 6
                    {
                        frontLineSchedule.insertRow(excelRow);

                        //Write the employees name to the ExcelWriter
                        frontLineSchedule.overwriteEmptyCell(excelCol, excelRow, name, 1);

                        //Write that employees shift time next to the employee's name
                        frontLineSchedule.overwriteEmptyCell(excelCol + 1, excelRow, myArray[x][y].getShiftTime(), 2);

                        //Make boarders for breaks and lunch cells
                        frontLineSchedule.overwriteEmptyCell(excelCol + 2, excelRow, "", 1);
                        frontLineSchedule.overwriteEmptyCell(excelCol + 3, excelRow, "", 1);
                        frontLineSchedule.overwriteEmptyCell(excelCol + 4, excelRow, "", 1);
                    }
                }
            }
        }
    }

    private static void displayCashierShifts(String day, Shift[][] myArray, int rows,
                                             int cols, ExcelWriter cashierSchedule) throws WriteException
    {
        int excelCol = 0;

        //Start at cell A5 for managers
        int openingExcelRow = 4;
        int midExcelRow = 4;
        int closingExcelRow = 4;

        //Start at cell A7 for cashiers
        int cashierExcelRow = 7;
        int cashierRowStart = cashierExcelRow;

        int x, y;

        for(x = 0; x < rows; x++)
        {
            for(y = 0; y < cols; y++)
            {
                //Make sure to filter out shifts that end at or before 9:00a
                if(myArray[x][y] != null && myArray[x][y].day.equals(day))
                {
                    //Add managers starting at cell A5
                    if(!myArray[x][y].getEmployeeType().equals(Shift.EmployeeType.NON_MANAGER))
                    {
                        if(myArray[x][y].getEmployeeType().equals(Shift.EmployeeType.OPENING_MANAGER))
                        {
                            //Write the manager name to the ExcelWriter
                            cashierSchedule.overwriteEmptyCell(excelCol, openingExcelRow, myArray[x][y].getName(), 1);

                            //Make boarders for the empty adjacent cells
                            if(cashierSchedule.isEmptyCell(excelCol + 1, openingExcelRow))
                                cashierSchedule.overwriteEmptyCell(excelCol + 1, openingExcelRow, "", 1);
                            if(cashierSchedule.isEmptyCell(excelCol + 2, openingExcelRow))
                                cashierSchedule.overwriteEmptyCell(excelCol + 2, openingExcelRow, "", 1);

                            openingExcelRow++;

                            //Determine if a row should be inserted based on how close the current row is to the cashier
                            //row start position
                            if((openingExcelRow + 2) >= cashierRowStart)
                            {
                                //Insert a new row at the corresponding location for opening managers
                                cashierSchedule.insertRow(openingExcelRow);
                                cashierRowStart++;
                                cashierExcelRow++;
                            }
                        }
                        else if(myArray[x][y].getEmployeeType().equals(Shift.EmployeeType.MID_SHIFT_MANAGER))
                        {
                            //Write the manager name to the ExcelWriter
                            cashierSchedule.overwriteEmptyCell(excelCol + 1, midExcelRow, myArray[x][y].getName(), 1);

                            //Make boarders for the empty adjacent cells
                            if(cashierSchedule.isEmptyCell(excelCol, midExcelRow))
                                cashierSchedule.overwriteEmptyCell(excelCol, midExcelRow, "", 1);
                            if(cashierSchedule.isEmptyCell(excelCol + 2, midExcelRow))
                                cashierSchedule.overwriteEmptyCell(excelCol + 2, midExcelRow, "", 1);

                            midExcelRow++;

                            if((midExcelRow + 2) >= cashierRowStart)
                            {
                                //Insert a new row at the corresponding location for mid-shift managers
                                cashierSchedule.insertRow(midExcelRow);
                                cashierRowStart++;
                                cashierExcelRow++;
                            }
                        }
                        else if(myArray[x][y].getEmployeeType().equals(Shift.EmployeeType.CLOSING_MANAGER))
                        {
                            //Write the manager name to the ExcelWriter
                            cashierSchedule.overwriteEmptyCell(excelCol + 2, closingExcelRow, myArray[x][y].getName(), 1);

                            //Make boarders for the empty adjacent cells
                            if(cashierSchedule.isEmptyCell(excelCol, closingExcelRow))
                                cashierSchedule.overwriteEmptyCell(excelCol, closingExcelRow, "", 1);
                            if(cashierSchedule.isEmptyCell(excelCol + 1, closingExcelRow))
                                cashierSchedule.overwriteEmptyCell(excelCol + 1, closingExcelRow, "", 1);

                            closingExcelRow++;

                            if((closingExcelRow + 2) >= cashierRowStart)
                            {
                                //Insert a new row at the corresponding location for closing managers
                                cashierSchedule.insertRow(closingExcelRow);
                                cashierRowStart++;
                                cashierExcelRow++;
                            }
                        }
                    }

                    //Write the cashier shifts to the cashierSchedule
                    if(myArray[x][y].position.contains("Cashier") && getMilitaryTime(myArray[x][y].endTime) > 900)
                    {
                        //Write the employees name to the ExcelWriter
                        cashierSchedule.overwriteEmptyCell(excelCol, cashierExcelRow, myArray[x][y].getName(), 1);

                        //Write that employees shift time next to the employee's name
                        cashierSchedule.overwriteEmptyCell(excelCol + 1, cashierExcelRow, myArray[x][y].getShiftTime(), 2);

                        //Make boarders for the breaks, lunch, and comments cells
                        cashierSchedule.overwriteEmptyCell(excelCol + 2, cashierExcelRow, "", 1);
                        cashierSchedule.overwriteEmptyCell(excelCol + 3, cashierExcelRow, "", 1);
                        cashierSchedule.overwriteEmptyCell(excelCol + 4, cashierExcelRow, "", 1);
                        cashierSchedule.overwriteEmptyCell(excelCol + 5, cashierExcelRow, "", 1);

                        cashierExcelRow++;
                    }
                }
            }
        }

        //Add extra rows for additional shifts to be written in by hand
        for(x = 0; x < 5; x++)
        {
            for(y = 0; y < 6; y++)
                cashierSchedule.overwriteEmptyCell(excelCol + y, cashierExcelRow, "", 1);

            cashierExcelRow++;
        }

        //Create a section for notes
        cashierSchedule.overwriteEmptyCell(excelCol, cashierExcelRow + 1, "Notes:", 3);
    }

    //Get time in standard format and convert it to military time for comparison purposes
    public static int getMilitaryTime(String standardTime)
    {
        int militaryTime = 0;
        String[] tmp;
        boolean pastNoon = false;

        if(standardTime.equals("2400"))
        {
            try
            {
                militaryTime = Integer.parseInt(standardTime);
            }
            catch(NumberFormatException e)
            {
                e.printStackTrace();
            }
            return militaryTime;
        }
        else
        {
            tmp = standardTime.split(":");

            if(standardTime.contains("p"))
            {
                pastNoon = true;
            }

            //Prevent IndexOutOfBoundsException
            if(tmp.length < 2)
                return -1;

            //Convert to numeric version by eliminating the 'a' or 'p'
            tmp[1] = tmp[1].replaceAll("[^\\d.]", "");

            try
            {
                militaryTime = Integer.parseInt(tmp[0] + tmp[1]);
            }
            catch (NumberFormatException e)
            {
                e.printStackTrace();
            }

            if(pastNoon)
            {
                if(militaryTime < 1200)
                    return militaryTime + 1200;

                return militaryTime;
            }
            else
            {
                if(militaryTime < 1200)
                    return militaryTime;

                return militaryTime - 1200;
            }
        }
    }

    //Searching for shift times
    public static boolean startsWithNumber(String shiftTime)
    {
        String[] nums = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"};

        for (String num : nums)
            if (shiftTime.startsWith(num))
                return true;
        return false;
    }

    private static int indexOfSunday(String[] tokens)
    {
        for (int i = 0; i < tokens.length; i++)
            if (tokens[i].contains("Sun"))
                return i;
        return -1;
    }
}