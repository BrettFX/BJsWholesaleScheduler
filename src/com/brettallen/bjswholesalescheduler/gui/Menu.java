package com.brettallen.bjswholesalescheduler.gui;

import java.awt.*;
import java.awt.event.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Calendar;

import javax.imageio.ImageIO;
import javax.mail.MessagingException;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.DefaultCaret;
import javax.swing.text.Document;

import com.brettallen.bjswholesalescheduler.BJsWholesaleScheduler;
import com.brettallen.bjswholesalescheduler.utils.*;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

/*
 * Created by Brett Allen on 7/17/2016 at 6:13 PM.
 *
 *      Main GUI for the BJsWholesaleScheduler program
 */
public class Menu extends JFrame implements ActionListener
{
    private static final long serialVersionUID = 1L;
    private static final String VERSION = "4.5.7 (Beta v2)";

    //Toggle this on for a better debugging experience (Don't have to open a file to test button events)
    private static final boolean DEBUG = false;

    private static final String AUTHENTICATION_MESSAGE = "Authentication Required";

    private Shift[][] employeeSchedule;
    private int rows,
            cols;
    private ArrayList<String> truncDates;

    //Enum to determine the sorted state of the schedule
    private enum SortedOrder {UNSORTED, ASCENDING, DESCENDING}

    //Incoming schedule will not be sorted in version 4.3.0
    private SortedOrder currentOrderState;

    //Enum to determine the file-load state
    private enum FileLoadState {INIT, LOADED, RESET}

    private FileLoadState currentFileLoadState;

    //Objects for mailing
    private final EmailFile emailFile;
    private String selectedEmail;
    private String password;
    private DefaultListModel emailingListModel;
    private ArrayList<String> emails;
    private String[] emailSpliter;
    private JTextField txtEmailName;
    private JTextField txtEmailAddress;
    private JTextField txtManualEmailName;
    private JTextField txtManualEmailAddress;
    private JList emailList;
    private JButton btnAddEmail;
    private JButton btnResetAddEmail;
    private JButton btnDeleteEmail;
    private JButton btnResetManualEmail;
    private JButton btnSendManualEmail;
    private JButton btnSendAll;
    private JButton btnClearManualName;
    private JButton btnClearManualEmail;
    private JRadioButton rdbtnEmailAsTxt;
    private  JRadioButton rdbtnEmailAsPic;
    private JRadioButton rdbtnEmailAsPDF;

    //emailList default separator
    public static final String DEFAULT_SEPARATOR = ": ";

    //EmailMode for deleting or adding emails to the emails.txt file
    private enum EmailMode {DELETING_SELECTED, DELETING_ALL, ADDING}
    //End objects for mailing

    private String currentEmployee;

    //Create tab objects
    private JTextField txtXLSPath;
    private JTextField txtPDFPath;
    private JCheckBox cbxFrontLine;
    private JCheckBox cbxOvernight;
    private JCheckBox cbxCashier;
    private JButton btnSunday;
    private JButton btnMonday;
    private JButton btnTuesday;
    private JButton btnWednesday;
    private JButton btnThursday;
    private JButton btnFriday;
    private JButton btnSaturday;
    private JButton btnAll;
    private JButton btnOpenXlsFile;
    private JButton btnOpenPdfFile;
    private JButton btnReset;
    private JTextArea logArea;
    private String currentFilePath;
    //End Create tab objects

    //Search tab objects
    private JButton btnSearch;
    private JButton btnScreenshot;
    private JButton btnScreenshotAll;
    private JLabel lblSortedOrder;
    private ArrayList<String> screenshots;
    private JTextField txtSearch;
    private JLabel lblSunday;
    private JLabel lblMonday;
    private JLabel lblTuesday;
    private JLabel lblWednesday;
    private JLabel lblThursday;
    private JLabel lblFriday;
    private JLabel lblSaturday;
    private JLabel lblSearch;
    private JTextField txtSunday;
    private JTextField txtMonday;
    private JTextField txtTuesday;
    private JTextField txtWednesday;
    private JTextField txtFriday;
    private JTextField txtSaturday;
    private JTextField txtThursday;
    private JTextField txtEmployee;
    //End Search tab objects

    //Delegate menu object
    public Menu menu;

    public Menu()
    {
        //Set the title of the JFrame
        super("BJ's Wholesale Scheduler " + VERSION);

        //Initialize password
        password = "";

        try
        {
            // Set System L&F
            UIManager.setLookAndFeel("com.jtattoo.plaf.texture.TextureLookAndFeel");
        }
        catch (UnsupportedLookAndFeelException e){e.printStackTrace();}
        catch (ClassNotFoundException e){e.printStackTrace();}
        catch (InstantiationException e){e.printStackTrace();}
        catch (IllegalAccessException e){e.printStackTrace();}

        //Create/get email file
        emailFile = new EmailFile();

        //Initialize currentOrderState to UNSORTED
        currentOrderState = SortedOrder.UNSORTED;

        //Initialize currentFileLoadState to INIT
        currentFileLoadState = FileLoadState.INIT;

        getContentPane().setFont(new Font("Consolas", Font.PLAIN, 13));

        //Set icon image to BJ's logo
        this.setIconImage(new ImageIcon(getClass().getResource("logo.png")).getImage());

        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        setBounds(100, 100, 676, 747); //100, 100, 676, 717
        getContentPane().setLayout(null);

        JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
        tabbedPane.setFont(new Font("Consolas", Font.PLAIN, 13));
        tabbedPane.setBounds(0, 0, 669, 710); //0, 0, 669, 680

        //Create tab
        JComponent createPanel = makeTextPanel("create");
        tabbedPane.add("Create", createPanel);

        //Search tab
        JComponent searchPanel = makeTextPanel("search");
        tabbedPane.add("Search", searchPanel);

        //Email tab
        JComponent emailPanel = makeTextPanel("email");
        tabbedPane.add("Email", emailPanel);

        //Add the tabs to the content pane
        getContentPane().add(tabbedPane);

        //Make it so the JFrame appears center screen
        setLocationRelativeTo(null);

        this.displayInstructions();

        //Make the debugging experience a bit easier by enabling everything without having to open files
        //or select checkboxes to enable buttons
        if(DEBUG)
        {
            toggleSearch(true);
            toggleEmail(true);
            toggleCheckboxes(true);
            toggleButtons(true);
            btnReset.setEnabled(true);
        }

        /*NB: ONCE A LIST OF ALL EMPLOYEES AS BEEN CREATED AND STORED IN REGISTRY getScreenshotFiles CAN BE
         * PLACED HERE SO THAT A SCHEDULE DOES NOT NEED TO BE LOADED IN ORDER TO SEND CURRENTLY EXISTING
          * SCREENSHOTS*/

        //Try loading all previously existing screenshot files into the screenshots ArrayList
        getScreenshotFiles();
    }

    //Inherited classes can access this
    private JComponent makeTextPanel(String panelIndex)
    {
        JPanel panel = new JPanel(false);
        panel.setLayout(null);

        //Place contents of panel here for editing

        //Panel1: Create tab
        if(panelIndex.equals("create"))
        {
            btnOpenPdfFile = new JButton("Open PDF");
            btnOpenPdfFile.setBounds(12, 43, 97, 25); //12, 13, 97, 25
            btnOpenPdfFile.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try
                    {
                        openFileDialog(new FileNameExtensionFilter("PDF (*.pdf)", "pdf"));
                    }
                    catch (BiffException e1) {e1.printStackTrace();}
                    catch (BadLocationException e1) {e1.printStackTrace();}
                }
            });
            panel.add(btnOpenPdfFile);

            btnOpenXlsFile = new JButton("Open XLS");
            btnOpenXlsFile.setBounds(12, 13, 97, 25); //12, 43, 97, 25
            btnOpenXlsFile.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try
                    {
                        openFileDialog(new FileNameExtensionFilter("Excel 97-2003 Workbook (*.xls)", "xls"));
                    }
                    catch (BiffException e1) {e1.printStackTrace();}
                    catch (BadLocationException e1) {e1.printStackTrace();}
                }
            });
            panel.add(btnOpenXlsFile);

            txtXLSPath = new JTextField();
            txtXLSPath.setEditable(false);
            txtXLSPath.setBounds(121, 14, 520, 22); //121, 44, 520, 22
            panel.add(txtXLSPath);
            txtXLSPath.setColumns(10);

            txtPDFPath = new JTextField();
            txtPDFPath.setEditable(false);
            txtPDFPath.setBounds(121, 44, 520, 22); //121, 14, 520, 22
            panel.add(txtPDFPath);
            txtPDFPath.setColumns(10);

            JLabel lblCbxChooser = new JLabel("Select schedule(s) to create:");
            lblCbxChooser.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblCbxChooser.setBounds(12, 81, 266, 16); //12, 51, 266, 16
            panel.add(lblCbxChooser);

            cbxFrontLine = new JCheckBox("Front Line Schedule");
            cbxFrontLine.setEnabled(false);
            cbxFrontLine.setFont(new Font("Consolas", Font.PLAIN, 13));
            cbxFrontLine.setBounds(22, 106, 168, 25); //22, 76, 168, 25
            cbxFrontLine.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    if(cbxFrontLine.isSelected())
                    {
                        cbxOvernight.setSelected(false);

                        //Sort the schedule in ascending order if it has not been done so already
                        if(currentOrderState != SortedOrder.ASCENDING)
                        {
                            try
                            {
                                BJsWholesaleScheduler.sortAscending(employeeSchedule, rows, menu);
                                setOrderState(SortedOrder.ASCENDING);
                            }
                            catch(BadLocationException ble)
                            {
                                ble.printStackTrace();
                            }
                        }

                        toggleButtons(true);
                    }
                    else if (!cbxFrontLine.isSelected() && !cbxCashier.isSelected() && !cbxOvernight.isSelected())
                        toggleButtons(false);
                }
            });
            panel.add(cbxFrontLine);

            cbxCashier = new JCheckBox("Cashier Schedule");
            cbxCashier.setEnabled(false);
            cbxCashier.setFont(new Font("Consolas", Font.PLAIN, 13));
            cbxCashier.setBounds(22, 136, 164, 25); //22, 106, 164, 25
            cbxCashier.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    if(cbxCashier.isSelected())
                    {
                        cbxOvernight.setSelected(false);

                        //Sort the schedule in ascending order if it has not been done so already
                        if(currentOrderState != SortedOrder.ASCENDING)
                        {
                            try
                            {
                                BJsWholesaleScheduler.sortAscending(employeeSchedule, rows, menu);
                                setOrderState(SortedOrder.ASCENDING);
                            }
                            catch(BadLocationException ble)
                            {
                                ble.printStackTrace();
                            }
                        }

                        toggleButtons(true);
                    }
                    else if (!cbxFrontLine.isSelected() && !cbxCashier.isSelected() && !cbxOvernight.isSelected())
                        toggleButtons(false);
                }
            });
            panel.add(cbxCashier);

            cbxOvernight = new JCheckBox("Overnight Schedule");
            cbxOvernight.setEnabled(false);
            cbxOvernight.setFont(new Font("Consolas", Font.PLAIN, 13));
            cbxOvernight.setBounds(22, 166, 168, 25); //22, 136, 168, 25
            cbxOvernight.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    //If cbxOvernight is in its checked state then uncheck cbxCashier and cbxFrontLine components
                    //because the overnight function will have to temporarily sort the schedule in descending order
                    if(cbxOvernight.isSelected())
                    {
                        cbxCashier.setSelected(false);
                        cbxFrontLine.setSelected(false);

                        //Sort the schedule in descending order so that the individuals arriving before midnight
                        //are displayed first
                        if(currentOrderState != SortedOrder.DESCENDING)
                        {
                            try
                            {
                                BJsWholesaleScheduler.sortDescending(employeeSchedule, rows, menu);
                                setOrderState(SortedOrder.DESCENDING);
                            }
                            catch(BadLocationException ble)
                            {
                                ble.printStackTrace();
                            }
                        }

                        toggleButtons(true);
                    }
                    else if (!cbxFrontLine.isSelected() && !cbxCashier.isSelected() && !cbxOvernight.isSelected())
                        toggleButtons(false);
                }
            });
            panel.add(cbxOvernight);

            JLabel lblChoice = new JLabel("Choose day(s) to process:");
            lblChoice.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblChoice.setBounds(12, 200, 266, 16); //12, 170, 266, 16
            panel.add(lblChoice);

            btnSunday = new JButton("Sunday");
            btnSunday.setEnabled(false);
            btnSunday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnSunday.setBounds(22, 229, 256, 25); //22, 199, 256, 25
            btnSunday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Sunday", employeeSchedule, rows, cols, 1, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnSunday);

            btnMonday = new JButton("Monday");
            btnMonday.setEnabled(false);
            btnMonday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnMonday.setBounds(22, 267, 256, 25); //22, 237, 256, 25
            btnMonday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Monday", employeeSchedule, rows, cols, 2, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnMonday);

            btnTuesday = new JButton("Tuesday");
            btnTuesday.setEnabled(false);
            btnTuesday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnTuesday.setBounds(22, 305, 256, 25); //22, 275, 256, 25
            btnTuesday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Tuesday", employeeSchedule, rows, cols, 3, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnTuesday);

            btnWednesday = new JButton("Wednesday");
            btnWednesday.setEnabled(false);
            btnWednesday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnWednesday.setBounds(22, 343, 256, 25); //22, 313, 256, 25
            btnWednesday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Wednesday", employeeSchedule, rows, cols, 4, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnWednesday);

            btnThursday = new JButton("Thursday");
            btnThursday.setEnabled(false);
            btnThursday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnThursday.setBounds(22, 381, 256, 25); //22, 351, 256, 25
            btnThursday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Thursday", employeeSchedule, rows, cols, 5, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnThursday);

            btnFriday = new JButton("Friday");
            btnFriday.setEnabled(false);
            btnFriday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnFriday.setBounds(22, 419, 256, 25); //22, 389, 256, 25
            btnFriday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Friday", employeeSchedule, rows, cols, 6, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnFriday);

            btnSaturday = new JButton("Saturday");
            btnSaturday.setEnabled(false);
            btnSaturday.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnSaturday.setBounds(22, 457, 256, 25); //22, 427, 256, 25
            btnSaturday.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Saturday", employeeSchedule, rows, cols, 7, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnSaturday);

            btnAll = new JButton("All");
            btnAll.setEnabled(false);
            btnAll.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnAll.setBounds(22, 495, 256, 25); //22, 465, 256, 25
            btnAll.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    try{render("Sunday", employeeSchedule, rows, cols, 1, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}

                    try{render("Monday", employeeSchedule, rows, cols, 2, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}

                    try{render("Tuesday", employeeSchedule, rows, cols, 3, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}

                    try{render("Wednesday", employeeSchedule, rows, cols, 4, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}

                    try{render("Thursday", employeeSchedule, rows, cols, 5, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}

                    try{render("Friday", employeeSchedule, rows, cols, 6, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}

                    try{render("Saturday", employeeSchedule, rows, cols, 7, truncDates);}
                    catch (RowsExceededException e2){e2.printStackTrace();}
                    catch (BiffException e2) {e2.printStackTrace();}
                    catch (WriteException e2){e2.printStackTrace();}
                    catch (IOException e2){e2.printStackTrace();}
                    catch (BadLocationException e2){e2.printStackTrace();}
                }
            });
            panel.add(btnAll);

            JScrollPane scrollPane = new JScrollPane();
            scrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
            scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
            scrollPane.setBounds(304, 81, 337, 439); //304, 51, 337, 409
            panel.add(scrollPane);

            logArea = new JTextArea();
            logArea.setEditable(false);
            logArea.setFont(new Font("Consolas", Font.PLAIN, 16));
            logArea.setLineWrap(true);
            DefaultCaret logAreaCaret = (DefaultCaret)logArea.getCaret();
            logAreaCaret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);
            scrollPane.setViewportView(logArea);

            btnReset = new JButton("Reset");
            btnReset.setEnabled(false);
            btnReset.setFont(new Font("Consolas", Font.PLAIN, 16));
            btnReset.setBounds(22, 548, 256, 29); //22, 518, 256, 29
            btnReset.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    logArea.setText("");
                    txtXLSPath.setText("");
                    txtPDFPath.setText("");
                    lblSearch.setText("");

                    resetDays();
                    resetDayButtonTexts();
                    resetSearchFields();
                    resetEmailFields();

                    uncheckCheckboxState();

                    toggleCheckboxes(false);
                    toggleButtons(false);
                    toggleSearch(false);

                    setOrderState(SortedOrder.UNSORTED);

                    rdbtnEmailAsTxt.setSelected(false);
                    rdbtnEmailAsPic.setSelected(false);
                    btnReset.setEnabled(false);
                    btnScreenshot.setEnabled(false);
                }
            });
            panel.add(btnReset);

            JSeparator separator = new JSeparator();
            separator.setBounds(12, 533, 629, 2); //12, 503, 629, 2
            panel.add(separator);

            JButton btnInstructions = new JButton("Instructions");
            btnInstructions.setFont(new Font("Consolas", Font.PLAIN, 16));
            btnInstructions.setBounds(22, 586, 256, 27); //22, 556, 256, 27
            btnInstructions.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    displayInstructions();
                }
            });
            panel.add(btnInstructions);

            JButton btnCopyright = new JButton("Created By: Brett Allen");
            btnCopyright.setFont(new Font("Consolas", Font.PLAIN, 16));
            btnCopyright.setEnabled(false);
            btnCopyright.setBounds(304, 548, 337, 103); //304, 518, 337, 103
            panel.add(btnCopyright);

            JButton btnTutorial = new JButton("Video Tutorial");
            btnTutorial.setFont(new Font("Consolas", Font.PLAIN, 16));
            btnTutorial.setBounds(22, 624, 256, 27); //22, 594, 256, 27
            btnTutorial.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    Media tutorialVideo;

                    //Get the tutorial video from the default file txtXLSPath
                    String tutorialPath = new File("tutorial.avi").getAbsolutePath();

                    tutorialVideo = new Media(tutorialPath);
                    tutorialVideo.playVideo();

                    //If the tutorial video is not there, display appropriate message
                    if(!tutorialVideo.fileExists)
                    {
                        try
                        {
                            menu.print("Could not find " + tutorialPath + ". Please place the file outside of\n" +
                                    "the program's executable or play the video externally.\n");
                        }
                        catch(BadLocationException ble)
                        {
                            ble.printStackTrace();
                        }
                    }
                }
            });
            panel.add(btnTutorial);
        }

        //Panel2: Search tab
        if(panelIndex.equals("search"))
        {
            lblSearch = new JLabel("Enter employee's name:");
            lblSearch.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblSearch.setBounds(12, 13, 505, 16);
            panel.add(lblSearch);

            btnSearch = new JButton("Search");
            btnSearch.setEnabled(false);
            btnSearch.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnSearch.setBounds(22, 42, 97, 25);
            btnSearch.setMnemonic(KeyEvent.VK_ENTER);
            btnSearch.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    //Find the employee schedule and show messages
                    findEmployeeSchedule(true);
                }
            });
            panel.add(btnSearch);

            txtSearch = new JTextField();
            txtSearch.setEnabled(false);
            txtSearch.setFont(new Font("Consolas", Font.PLAIN, 16));
            txtSearch.setBounds(130, 45, 504, 22);
            panel.add(txtSearch);
            txtSearch.setColumns(10);

            JLabel lblFound = new JLabel("Employee's schedule:");
            lblFound.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblFound.setBounds(12, 80, 198, 16);
            panel.add(lblFound);

            lblSunday = new JLabel("Sunday");
            lblSunday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblSunday.setBounds(36, 116, 264, 16);
            panel.add(lblSunday);

            lblMonday = new JLabel("Monday");
            lblMonday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblMonday.setBounds(36, 179, 264, 16);
            panel.add(lblMonday);

            lblTuesday = new JLabel("Tuesday");
            lblTuesday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblTuesday.setBounds(36, 250, 264, 16);
            panel.add(lblTuesday);

            lblWednesday = new JLabel("Wednesday");
            lblWednesday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblWednesday.setBounds(36, 320, 264, 16);
            panel.add(lblWednesday);

            lblThursday = new JLabel("Thursday");
            lblThursday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblThursday.setBounds(36, 387, 264, 16);
            panel.add(lblThursday);

            lblFriday = new JLabel("Friday");
            lblFriday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblFriday.setBounds(36, 452, 264, 16);
            panel.add(lblFriday);

            lblSaturday = new JLabel("Saturday");
            lblSaturday.setFont(new Font("Consolas", Font.PLAIN, 16));
            lblSaturday.setBounds(36, 520, 264, 16);
            panel.add(lblSaturday);

            txtSunday = new JTextField();
            txtSunday.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtSunday.setEditable(false);
            txtSunday.setBounds(36, 134, 598, 22); //36, 134, 598, 22
            panel.add(txtSunday);
            txtSunday.setColumns(10);

            txtMonday = new JTextField();
            txtMonday.setFont(new Font("Consolas", Font.PLAIN, 11)); //"Consolas", Font.PLAIN, 13
            txtMonday.setEditable(false);
            txtMonday.setColumns(10);
            txtMonday.setBounds(36, 198, 598, 22); //36, 198, 598, 22

            panel.add(txtMonday);

            txtTuesday = new JTextField();
            txtTuesday.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtTuesday.setEditable(false);
            txtTuesday.setColumns(10);
            txtTuesday.setBounds(36, 271, 598, 22);
            panel.add(txtTuesday);

            txtWednesday = new JTextField();
            txtWednesday.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtWednesday.setEditable(false);
            txtWednesday.setColumns(10);
            txtWednesday.setBounds(36, 337, 598, 22);
            panel.add(txtWednesday);

            txtThursday = new JTextField();
            txtThursday.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtThursday.setEditable(false);
            txtThursday.setColumns(10);
            txtThursday.setBounds(36, 404, 598, 22);
            panel.add(txtThursday);

            txtFriday = new JTextField();
            txtFriday.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtFriday.setEditable(false);
            txtFriday.setColumns(10);
            txtFriday.setBounds(36, 470, 598, 22);
            panel.add(txtFriday);

            txtSaturday = new JTextField();
            txtSaturday.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtSaturday.setEditable(false);
            txtSaturday.setColumns(10);
            txtSaturday.setBounds(36, 537, 598, 24);
            panel.add(txtSaturday);

            txtEmployee = new JTextField();
            txtEmployee.setFont(new Font("Consolas", Font.PLAIN, 13));
            txtEmployee.setEditable(false);
            txtEmployee.setColumns(10);
            txtEmployee.setBounds(220, 78, 414, 22);
            panel.add(txtEmployee);

            JRootPane rootPane = getRootPane();
            rootPane.setDefaultButton(btnSearch);

            screenshots = new ArrayList<String>();

            /*Take a screenshot of the employee that was found via search*/
            btnScreenshot = new JButton("Screenshot Current");
            btnScreenshot.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnScreenshot.setBounds(500, 584, 150, 25); //522, 584, 112, 25
            btnScreenshot.setEnabled(false);
            btnScreenshot.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent a)
                {
                    //Take a screenshot of the search tab and show messages
                    takeScreenshot(true);
                }
            });
            panel.add(btnScreenshot);

            /*Take a screen shot of all the employees that are in the email list*/
            btnScreenshotAll = new JButton("Screenshot All");
            btnScreenshotAll.setFont(new Font("Consolas", Font.PLAIN, 13));
            btnScreenshotAll.setBounds(522, 614, 112, 25);
            btnScreenshotAll.setEnabled(false);
            btnScreenshotAll.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent a)
                {
                    //Traverse the email ArrayList and get each employees' name and store it in the search text field
                    //Search and take a screenshot of each employee
                    for(String recipient : emails)
                    {
                        txtSearch.setText(recipient.split(DEFAULT_SEPARATOR)[0].trim());

                        //Only take a screenshot if the employee was found and don't show any messages (faster)
                        if(findEmployeeSchedule(false))
                            takeScreenshot(false);
                    }

                    //Reset the text fields
                    resetSearchFields();
                    resetDays();

                    JOptionPane.showMessageDialog(null, "Successfully took a screenshot of all employees in the email list\n" +
                                "who are working this week.",
                        "OPERATION COMPLETE", JOptionPane.INFORMATION_MESSAGE);
                }
            });
            panel.add(btnScreenshotAll);

            //Create a label to display the sorted order of the schedule
            lblSortedOrder = new JLabel("Sorted Order: " + currentOrderState);
            lblSortedOrder.setFont(new Font("Consolas", Font.PLAIN, 18));
            lblSortedOrder.setBounds(36, 614, 300, 25); //522, 614, 112, 25
            panel.add(lblSortedOrder);
        }

        //Panel3: Email tab
        if(panelIndex.equals("email"))
        {
            JLabel lblEmailTo = new JLabel("Employee Email List:");
            lblEmailTo.setFont(new Font("Consolas", Font.BOLD, 12));
            lblEmailTo.setBounds(10, 11, 148, 14);
            panel.add(lblEmailTo);

            JButton btnEmailListReset = new JButton("Refresh Email List");
            btnEmailListReset.setFont(new Font("Consolas", Font.BOLD, 12));
            btnEmailListReset.setBounds(10, 31, 148, 104);
            btnEmailListReset.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    //Delete all emails from the email list
                    writeToEmailFile(EmailMode.DELETING_ALL);
                }
            });
            panel.add(btnEmailListReset);

            //Emailing
            JScrollPane spEmailList = new JScrollPane();
            spEmailList.setBounds(168, 11, 475, 124);
            panel.add(spEmailList);

            emailingListModel = new DefaultListModel();
            emails = new ArrayList<String>();

            //noinspection unchecked
            emailList = new JList(emailingListModel);
            emailList.addMouseListener(new MouseAdapter()
            {
                @Override
                public void mouseClicked(MouseEvent e)
                {
                    selectedEmail = emailList.getSelectedValue().toString();

                    emailSpliter = selectedEmail.split(DEFAULT_SEPARATOR);
                    System.out.println("SELECTED: " + selectedEmail);

                    btnDeleteEmail.setEnabled(true);

                    //Set the text of the manual email fields to selected entry within emailList
                    //Only if the text fields are enabled
                    if(txtManualEmailName.isEnabled() && txtManualEmailAddress.isEnabled())
                    {
                        txtManualEmailName.setText(emailSpliter[0].trim());
                        txtManualEmailAddress.setText(emailSpliter[1].trim());
                    }
                }
            });

            spEmailList.setViewportView(emailList);

            txtEmailName = new JTextField();
            txtEmailName.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtEmailName.setBounds(123, 174, 148, 20);
            txtEmailName.addKeyListener(new KeyListener()
            {
                @Override
                public void keyTyped(KeyEvent e)
                {
                    //Determine if both txtEmailName and txtEmailAddress have content
                    determineIfReady();
                }

                @Override
                public void keyPressed(KeyEvent e)
                {
                    //Determine if both txtEmailName and txtEmailAddress have content
                    determineIfReady();
                }

                @Override
                public void keyReleased(KeyEvent e)
                {
                    //Determine if both txtEmailName and txtEmailAddress have content
                    determineIfReady();
                }
            });
            panel.add(txtEmailName);
            txtEmailName.setColumns(10);

            btnAddEmail = new JButton("Add Email");
            btnAddEmail.setFont(new Font("Consolas", Font.PLAIN, 12));
            btnAddEmail.setEnabled(false);
            btnAddEmail.setBounds(10, 219, 170, 23);
            btnAddEmail.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent a)
                {
                    writeToEmailFile(EmailMode.ADDING);

                    btnAddEmail.setEnabled(false);
                    btnResetAddEmail.setEnabled(false);
                }
            });
            panel.add(btnAddEmail);

            JLabel lblEmailListForm = new JLabel("Add an Employee to the Email List (This will save to a file):");
            lblEmailListForm.setFont(new Font("Consolas", Font.BOLD, 12));
            lblEmailListForm.setBounds(10, 152, 633, 14);
            panel.add(lblEmailListForm);

            JLabel lblEmailToName = new JLabel("*Employee Name:");
            lblEmailToName.setFont(new Font("Consolas", Font.PLAIN, 12));
            lblEmailToName.setBounds(10, 177, 114, 14);
            panel.add(lblEmailToName);

            JLabel lblEmailAddress = new JLabel("*Email Address:");
            lblEmailAddress.setFont(new Font("Consolas", Font.PLAIN, 12));
            lblEmailAddress.setBounds(326, 177, 114, 14);
            panel.add(lblEmailAddress);

            txtEmailAddress = new JTextField();
            txtEmailAddress.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtEmailAddress.setColumns(10);
            txtEmailAddress.setBounds(435, 174, 148, 20);
            txtEmailAddress.addKeyListener(new KeyListener()
            {
                @Override
                public void keyTyped(KeyEvent e)
                {
                    //Determine if both txtEmailName and txtEmailAddress have content
                    determineIfReady();
                }

                @Override
                public void keyPressed(KeyEvent e)
                {
                    //Determine if both txtEmailName and txtEmailAddress have content
                    determineIfReady();
                }

                @Override
                public void keyReleased(KeyEvent e)
                {
                    //Determine if both txtEmailName and txtEmailAddress have content
                    determineIfReady();
                }
            });
            panel.add(txtEmailAddress);

            JSeparator separator_1 = new JSeparator();
            separator_1.setBounds(10, 367, 622, 2);
            panel.add(separator_1);

            JLabel lblManualEmail = new JLabel("Manually Send Schedule to Employee as Email (This will not save):");
            lblManualEmail.setFont(new Font("Consolas", Font.BOLD, 12));
            lblManualEmail.setBounds(10, 380, 633, 14);
            panel.add(lblManualEmail);

            JLabel lblManualEmailName = new JLabel("*Employee Name:");
            lblManualEmailName.setFont(new Font("Consolas", Font.PLAIN, 12));
            lblManualEmailName.setBounds(10, 408, 114, 14);
            panel.add(lblManualEmailName);

            txtManualEmailName = new JTextField();
            txtManualEmailName.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtManualEmailName.setColumns(10);
            txtManualEmailName.setBounds(123, 405, 148, 20);
            panel.add(txtManualEmailName);

            JLabel lblManualEmailAddress = new JLabel("*Email Address:");
            lblManualEmailAddress.setFont(new Font("Consolas", Font.PLAIN, 12));
            lblManualEmailAddress.setBounds(326, 408, 114, 14);
            panel.add(lblManualEmailAddress);

            txtManualEmailAddress = new JTextField();
            txtManualEmailAddress.setFont(new Font("Consolas", Font.PLAIN, 11));
            txtManualEmailAddress.setColumns(10);
            txtManualEmailAddress.setBounds(435, 405, 148, 20);
            panel.add(txtManualEmailAddress);

            btnSendManualEmail = new JButton("Send to Individual");
            btnSendManualEmail.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    if(!rdbtnEmailAsTxt.isSelected() && !rdbtnEmailAsPic.isSelected())
                    {
                        JOptionPane.showMessageDialog(null, "Please select how you would like to send the email.",
                                "ERROR: EMAILING METHOD NOT SELECTED", JOptionPane.ERROR_MESSAGE);
                    }
                    else
                    {
                        //Only send email if required information is available
                        if(!txtManualEmailName.getText().isEmpty() && !txtManualEmailAddress.getText().isEmpty())
                        {
                            //Loading...
                            setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

                            //Authenticate once and create an email object for sending as text
                            if(password.equals(""))
                                password = getPassword();

                            //If the user cancels do not process
                            if(!password.equals(""))
                            {
                                Email email = new Email(password);

                                if(rdbtnEmailAsPic.isSelected())
                                {
                                    //If the corresponding screenshot of employee's schedule was not found, display appropriate error message
                                    //Otherwise send the employee's schedule as a picture
                                    if(!sendScheduleAsPic())
                                    {
                                        String msg = "Screenshot must be taken for " + txtManualEmailName.getText() + " first, ";
                                        msg += "or an invalid password was entered.";

                                        JOptionPane.showMessageDialog(null, msg, "ERROR", JOptionPane.ERROR_MESSAGE);

                                        email.errorInEmailProcess  = true;
                                    }
                                }
                                else if(rdbtnEmailAsTxt.isSelected())
                                {
                                    if(isEmployee(employeeSchedule, rows, cols, txtManualEmailName.getText()))
                                    {
                                        System.out.println("Sending email for " + txtManualEmailName.getText() + " to "
                                                + txtManualEmailAddress.getText() + "...");

                                        try
                                        {
                                            email.sendGenericMessage(txtManualEmailAddress.getText(), getEmailSubject(),
                                                    getEmailMessage(txtManualEmailName.getText()));
                                        }
                                        catch(MessagingException me)
                                        {
                                            email.errorInEmailProcess = true;

                                            System.err.println("Messaging Exception was Thrown!");
                                            JOptionPane.showMessageDialog(null, Email.DEV_INFO + e.toString(),
                                                    "ERROR", JOptionPane.ERROR_MESSAGE);
                                            me.printStackTrace();
                                        }
                                    }
                                    else
                                    {
                                        System.err.println(txtManualEmailName.getText() + " either is not working this week or is not an employee.");
                                    }
                                }

                                if(!email.errorInEmailProcess)
                                {
                                    JOptionPane.showMessageDialog(null, "Schedule for " + txtManualEmailName.getText() + " was sent to " +
                                                    txtManualEmailAddress.getText() + " successfully!",
                                            "OPERATION COMPLETE", JOptionPane.INFORMATION_MESSAGE);
                                }
                                else
                                {
                                    email.errorInEmailProcess = false;
                                    password = "";

                                    JOptionPane.showMessageDialog(null, "Authentication Initialized...",
                                            "OPERATION FAILED", JOptionPane.ERROR_MESSAGE);
                                }
                            }
                        }
                        else
                        {
                            JOptionPane.showMessageDialog(null, "One or more of the required fields is blank!", "ERROR", JOptionPane.ERROR_MESSAGE);
                        }
                    }

                    //Done loading.
                    setCursor(null);
                }
            });
            btnSendManualEmail.setFont(new Font("Consolas", Font.PLAIN, 12));
            btnSendManualEmail.setBounds(10, 446, 170, 23);
            panel.add(btnSendManualEmail);

            JSeparator separator_2 = new JSeparator();
            separator_2.setBounds(10, 480, 622, 2);
            panel.add(separator_2);

            final JCheckBox cbxToggleBtnSendAll = new JCheckBox("Unlock to Send All");
            cbxToggleBtnSendAll.setFont(new Font("Consolas", Font.PLAIN, 12));
            cbxToggleBtnSendAll.setBounds(10, 493, 200, 25);
            cbxToggleBtnSendAll.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    if(cbxToggleBtnSendAll.isSelected())
                        btnSendAll.setEnabled(true);
                    else
                        btnSendAll.setEnabled(false);
                }
            });
            panel.add(cbxToggleBtnSendAll);

            btnSendAll = new JButton("Send to All");
            btnSendAll.setFont(new Font("Consolas", Font.PLAIN, 12));
            btnSendAll.setBounds(10, 523, 633, 146); //10, 493, 633, 116
            btnSendAll.setEnabled(false);
            btnSendAll.addActionListener(new ActionListener()
            {
                @Override
                public void actionPerformed(ActionEvent e)
                {
                    //Create a boolean that will determine when the emailing process is complete
                    boolean done = false;

                    if(!rdbtnEmailAsTxt.isSelected() && !rdbtnEmailAsPic.isSelected())
                    {
                        JOptionPane.showMessageDialog(null, "Please select how you would like to send the email.",
                                "ERROR: EMAILING METHOD NOT SELECTED", JOptionPane.ERROR_MESSAGE);
                    }
                    else if(rdbtnEmailAsTxt.isSelected())
                    {
                        System.out.println("Sending schedules to all as text format...");

                        //Set cursor to loading symbol
                        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

                        //Traverse through the email list and send each schedule to the corresponding employee in
                        //text format
                        for(String recipient : emails)
                        {
                            String[] tokens = recipient.split(DEFAULT_SEPARATOR);

                            //Send emails to the employees that exist/are working this week
                            if(isEmployee(employeeSchedule, rows, cols, tokens[0].trim()))
                            {
                                if(password.equals(""))
                                    password = getPassword();

                                //Only process if the password is not blank
                                if(!password.equals(""))
                                {
                                    Email email = new Email(JOptionPane.showInputDialog(password));

                                    try
                                    {
                                        email.sendGenericMessage(tokens[1].trim(), getEmailSubject(), getEmailMessage(tokens[0].trim()));
                                    }
                                    catch(MessagingException me)
                                    {
                                        email.errorInEmailProcess = true;

                                        System.err.println("Messaging Exception was Thrown!");
                                        JOptionPane.showMessageDialog(null, Email.DEV_INFO + e.toString(),
                                                "ERROR", JOptionPane.ERROR_MESSAGE);
                                        me.printStackTrace();
                                    }
                                }
                            }
                            else
                            {
                                System.err.println(tokens[0].trim() + " either is not working this week or is not an employee.");
                            }
                        }

                        done = true;

                        //Done loading
                        setCursor(null);
                    }
                    else if(rdbtnEmailAsPic.isSelected())
                    {
                        System.out.println("Sending schedules to all as picture format...");

                        //Set cursor to loading symbol
                        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

                        //Traverse through the email list and send each schedule to the corresponding employee in
                        //picture format
                        if(!screenshots.isEmpty())
                        {
                            for(String recipient : emails)
                            {
                                //First set the individual email name and email address text fields to each recipient
                                //in the email list
                                txtManualEmailName.setText(recipient.split(DEFAULT_SEPARATOR)[0].trim());
                                txtManualEmailAddress.setText(recipient.split(DEFAULT_SEPARATOR)[1].trim());

                                //Send the email as a picture to the current email recipient
                                sendScheduleAsPic();
                            }

                            done = true;

                            //Done loading
                            setCursor(null);
                        }
                        else
                        {
                            JOptionPane.showMessageDialog(null, "No screenshots have been taken to send!",
                                    "ERROR", JOptionPane.ERROR_MESSAGE);
                        }
                    }

                    if(done)
                    {
                        JOptionPane.showMessageDialog(null, "Emailing process complete!",
                                "PROCESS COMPLETED", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            });
            panel.add(btnSendAll);

            btnResetAddEmail = new JButton("Clear Text Fields");
            btnResetAddEmail.setEnabled(false);
            btnResetAddEmail.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    txtEmailName.setText("");
                    txtEmailAddress.setText("");

                    btnAddEmail.setEnabled(false);
                    btnResetAddEmail.setEnabled(false);
                }
            });
            btnResetAddEmail.setFont(new Font("Consolas", Font.PLAIN, 12));
            btnResetAddEmail.setBounds(190, 219, 170, 23);
            panel.add(btnResetAddEmail);

            btnResetManualEmail = new JButton("Clear Text Fields");
            btnResetManualEmail.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    txtManualEmailName.setText("");
                    txtManualEmailAddress.setText("");
                }
            });
            btnResetManualEmail.setFont(new Font("Consolas", Font.PLAIN, 12));
            btnResetManualEmail.setBounds(190, 445, 170, 23);
            panel.add(btnResetManualEmail);

            btnDeleteEmail = new JButton("Delete Entry");
            btnDeleteEmail.setEnabled(false);
            btnDeleteEmail.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    writeToEmailFile(EmailMode.DELETING_SELECTED);
                    btnDeleteEmail.setEnabled(false);
                }
            });
            btnDeleteEmail.setFont(new Font("Consolas", Font.PLAIN, 12));
            btnDeleteEmail.setBounds(370, 220, 170, 23);
            panel.add(btnDeleteEmail);

            btnClearManualName = new JButton("X");
            btnClearManualName.setFont(new Font("Tahoma", Font.PLAIN, 8));
            btnClearManualName.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    txtManualEmailName.setText("");
                }
            });
            btnClearManualName.setBounds(281, 408, 39, 17);
            panel.add(btnClearManualName);

            btnClearManualEmail = new JButton("X");
            btnClearManualEmail.setFont(new Font("Tahoma", Font.PLAIN, 8));
            btnClearManualEmail.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    txtManualEmailAddress.setText("");
                }
            });
            btnClearManualEmail.setBounds(593, 408, 39, 17);
            panel.add(btnClearManualEmail);

            JButton btnClearListName = new JButton("X");
            btnClearListName.setFont(new Font("Tahoma", Font.PLAIN, 8));
            btnClearListName.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    txtEmailName.setText("");
                }
            });
            btnClearListName.setBounds(281, 177, 39, 17);
            panel.add(btnClearListName);

            JButton btnClearListEmail = new JButton("X");
            btnClearListEmail.setFont(new Font("Tahoma", Font.PLAIN, 8));
            btnClearListEmail.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    txtEmailAddress.setText("");
                }
            });
            btnClearListEmail.setBounds(593, 177, 39, 17);
            panel.add(btnClearListEmail);

            JSeparator separator_3 = new JSeparator();
            separator_3.setBounds(10, 253, 622, 2);
            panel.add(separator_3);

            JLabel lblEmailType = new JLabel("Select How to Send Email(s) (Sending as text is more efficient):");
            lblEmailType.setFont(new Font("Consolas", Font.BOLD, 12));
            lblEmailType.setBounds(10, 262, 622, 14);
            panel.add(lblEmailType);

            rdbtnEmailAsTxt = new JRadioButton("Send Schedule as Text Format");
            rdbtnEmailAsTxt.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent arg0)
                {
                    rdbtnEmailAsTxt.setSelected(true);
                    rdbtnEmailAsPic.setSelected(false);
                    rdbtnEmailAsPDF.setSelected(false);
                }
            });
            rdbtnEmailAsTxt.setFont(new Font("Consolas", Font.PLAIN, 11));
            rdbtnEmailAsTxt.setBounds(10, 294, 209, 23); //10, 294, 209, 23
            panel.add(rdbtnEmailAsTxt);

            rdbtnEmailAsPDF = new JRadioButton("Send Schedule As PDF Format Only");
            rdbtnEmailAsPDF.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent arg0)
                {
                    rdbtnEmailAsPDF.setSelected(true);
                    rdbtnEmailAsPic.setSelected(false);
                    rdbtnEmailAsTxt.setSelected(false);

                    if(txtPDFPath.getText().equals(""))
                    {
                        JOptionPane.showMessageDialog(null, "Please reference a PDF file to send via email.",
                                "IMPORTANT", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            });
            rdbtnEmailAsPDF.setFont(new Font("Consolas", Font.PLAIN, 11));
            rdbtnEmailAsPDF.setBounds(229, 294, 300, 23); //10, 294, 209, 23
            panel.add(rdbtnEmailAsPDF);

            rdbtnEmailAsPic = new JRadioButton("Send Schedule as Picture Format and PDF Format");
            rdbtnEmailAsPic.addActionListener(new ActionListener()
            {
                public void actionPerformed(ActionEvent e)
                {
                    rdbtnEmailAsPic.setSelected(true);
                    rdbtnEmailAsTxt.setSelected(false);
                    rdbtnEmailAsPDF.setSelected(false);

                    JOptionPane.showMessageDialog(null, "Make sure all screenshots have been taken for each employee that\n" +
                                    "you would like to send their schedule as a picture to.",
                            "IMPORTANT", JOptionPane.INFORMATION_MESSAGE);
                }
            });
            rdbtnEmailAsPic.setFont(new Font("Consolas", Font.PLAIN, 11));
            rdbtnEmailAsPic.setBounds(10, 330, 350, 23); //10, 330, 238, 23
            panel.add(rdbtnEmailAsPic);

            try
            {
                initializeEmailFile();
            }
            catch(IOException ioe)
            {
                ioe.printStackTrace();
                System.err.println("Attempt failed.");
            }
        }

        return panel;
    }

    private void initializeEmailFile() throws IOException
    {
        emails = emailFile.getEmailList();

        //If the emails ArrayList is empty try loading in the emails from the "emails.txt" file
        if(emails.isEmpty()) {
            System.err.println("No registry file found containing the email list.");
            System.out.println("Now attempting to load emails from \"emails.txt\"...");

            BufferedReader reader = new BufferedReader(new InputStreamReader(BJsWholesaleScheduler.BACKUP_EMAIL_FILE));

            String line;

            while ((line = reader.readLine()) != null)
            {
                System.out.println("Adding " + line + " to email list");
                emailFile.add(line.split(DEFAULT_SEPARATOR)[0].trim(), line.split(DEFAULT_SEPARATOR)[1].trim());
                emails.add(line);
            }

            reader.close();

            System.out.println("Emails have been loaded successfully.");
        }

        //Write all emails to the JList email model
        for(String email : emails)
        {
            //noinspection unchecked
            emailingListModel.addElement(email);
        }
    }

    private boolean isEmployee(Shift[][] myArray, int rows, int cols, String name)
    {
        btnScreenshot.setEnabled(false);

        for(int x = 0; x < rows; x++)
        {
            for(int y = 0; y < cols; y++)
            {
                if(myArray[x][y] != null)
                {
                    if(myArray[x][y].getFullName().toLowerCase().contains(name.toLowerCase())
                            || myArray[x][y].getName().toLowerCase().contains(name.toLowerCase()))
                    {
                        txtEmployee.setText(myArray[x][y].getFullName());
                        currentEmployee = myArray[x][y].getFullName();
                        btnScreenshot.setEnabled(true);
                        return true;
                    }
                }
            }
        }
        return false;
    }

    private String getEmployeeShift(String day, String name)
    {
        String shift = "";

        //Split full date into single day
        String[] truncDay = day.split(",");

        for(int x = 0; x < rows; x++)
        {
            for(int y = 0; y < cols; y++)
            {
                if(employeeSchedule[x][y] != null && employeeSchedule[x][y].day.contains(truncDay[0]))
                {
                    if(employeeSchedule[x][y].getFullName().contains(name) || employeeSchedule[x][y].getName().contains(name))
                    {
                        //If the employee has a special shift simply set the shift to their special shift and return
                        if(employeeSchedule[x][y].isSpecial())
                            return employeeSchedule[x][y].getShiftTime();

                        //Only prints a separator if the shift is a multi-shift
                        shift += employeeSchedule[x][y].getSeparator() + employeeSchedule[x][y].getShiftTime() + " " + employeeSchedule[x][y].position;
                    }
                }
            }
        }

        if(shift.equals(""))
            return "Off";
        return shift;
    }

    private void resetDays()
    {
        txtSunday.setText("");
        txtMonday.setText("");
        txtTuesday.setText("");
        txtWednesday.setText("");
        txtThursday.setText("");
        txtFriday.setText("");
        txtSaturday.setText("");
    }

    private void resetDayButtonTexts()
    {
        //Reset buttons on Create Tab
        btnSunday.setText("Sunday");
        btnMonday.setText("Monday");
        btnTuesday.setText("Tuesday");
        btnWednesday.setText("Wednesday");
        btnThursday.setText("Thursday");
        btnFriday.setText("Friday");
        btnSaturday.setText("Saturday");

        //Reset labels on Search Tab
        lblSunday.setText("Sunday");
        lblMonday.setText("Monday");
        lblTuesday.setText("Tuesday");
        lblWednesday.setText("Wednesday");
        lblThursday.setText("Thursday");
        lblFriday.setText("Friday");
        lblSaturday.setText("Saturday");
    }

    private void resetSearchFields()
    {
        txtSearch.setText("");
        txtEmployee.setText("");
    }

    private void resetEmailFields()
    {
        txtEmailName.setText("");
        txtEmailAddress.setText("");
        txtManualEmailName.setText("");
        txtManualEmailAddress.setText("");
    }

    public String getXlsPath()
    {
        if(txtXLSPath.getText() != null)
            return txtXLSPath.getText();
        return "";
    }

    public boolean getCbxFrontLineState()
    {
        return cbxFrontLine.isSelected();
    }

    public boolean getCbxCashierState()
    {
        return cbxCashier.isSelected();
    }

    public boolean getCbxOvernightState()
    {
        return cbxOvernight.isSelected();
    }

    public void transferVariables(Shift[][] myArray, int rows, ArrayList<String> truncDates)
    {
        this.employeeSchedule = myArray;
        this.rows = rows;
        this.cols = BJsWholesaleScheduler.DAYS;
        this.truncDates = truncDates;
    }

    private void openFileDialog(FileNameExtensionFilter filter) throws BadLocationException, BiffException
    {
        JFileChooser myPath = new JFileChooser();

        //Provides organization to dialog box: i.e. sorting options
        /*Action details = myPath.getActionMap().get("viewTypeDetails");
        details.actionPerformed(null);*/

        myPath.setFileFilter(filter);

        myPath.setCurrentDirectory(new File("C:/"));
        myPath.setDialogTitle("Open");

        String absolutePath;

        //Perform the correct action based on filter description
        //First determine if pdf button has been pressed
        if(filter.getDescription().toLowerCase().equals("pdf (*.pdf)"))
        {
            if(myPath.showOpenDialog(btnOpenPdfFile) != JFileChooser.APPROVE_OPTION)
            {
                //Only reset the txtPDFPath if the path was set to reference something other than a PDF file
                if(!txtPDFPath.getText().contains(".pdf"))
                    txtPDFPath.setText("");
            }
            else
            {
                absolutePath = myPath.getSelectedFile().getAbsolutePath();

                //Determine if the selected file was a pdf file
                if(absolutePath.contains(".pdf"))
                {
                    txtPDFPath.setText(absolutePath);
                    currentFilePath = txtPDFPath.getText();
                    JOptionPane.showMessageDialog(null, "PDF file has been referenced successfully", "INFORMATION", JOptionPane.INFORMATION_MESSAGE);
                }
                else if(txtPDFPath.getText().isEmpty())
                    txtPDFPath.setText(currentFilePath);
                else
                    JOptionPane.showMessageDialog(null, "You must choose a PDF File (.pdf)", "ERROR", JOptionPane.ERROR_MESSAGE);
            }
        }
        else if(filter.getDescription().toLowerCase().equals("excel 97-2003 workbook (*.xls)")) //Determine if xls button has been pressed
        {
            if(myPath.showOpenDialog(btnOpenXlsFile) != JFileChooser.APPROVE_OPTION)
            {
                //Only reset the txtXLSPath if the path was set to reference something other than an Excel file
                if(!txtXLSPath.getText().contains(".xls"))
                    txtXLSPath.setText("");
            }
            else
            {
                absolutePath = myPath.getSelectedFile().getAbsolutePath();

                //Enable processing buttons
                if(absolutePath.contains(".xls"))
                {
                    txtXLSPath.setText(absolutePath);
                    currentFilePath = txtXLSPath.getText();

                    //Set cursor to loading symbol
                    setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

                    //Call the main classes mainDelegate method when a new file has been chosen in order to process that file
                    BJsWholesaleScheduler.mainDelegate(menu);

                    //Initially sort the schedule in ascending order so that the schedules that are searched appear in the
                    //correct order
                    BJsWholesaleScheduler.sortAscending(employeeSchedule, rows, menu);

                    //Set the currentOrderState to ASCENDING order so that if the overnight checkbox is selected the program will
                    //know to sort in DESCENDING order
                    setOrderState(SortedOrder.ASCENDING);

                    //Toggle everything to true and uncheck the checkboxes
                    toggleSearch(true);
                    toggleEmail(true);
                    toggleCheckboxes(true);
                    uncheckCheckboxState();
                    btnReset.setEnabled(true);

                    //Turn off wait cursor
                    setCursor(null);

                    JOptionPane.showMessageDialog(null, "Schedules have been successfully processed!",
                            "PROCESS COMPLETED", JOptionPane.INFORMATION_MESSAGE);
                }
                else if(txtXLSPath.getText().isEmpty())
                    txtXLSPath.setText(currentFilePath);
                else
                    JOptionPane.showMessageDialog(null, "You must choose an Excel File (.xls)", "ERROR", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void toggleButtons(boolean b)
    {
        //Only enable the processing buttons when at least one type of schedule to be
        //process is selected
        btnSunday.setEnabled(b);
        btnMonday.setEnabled(b);
        btnTuesday.setEnabled(b);
        btnWednesday.setEnabled(b);
        btnThursday.setEnabled(b);
        btnFriday.setEnabled(b);
        btnSaturday.setEnabled(b);
        btnAll.setEnabled(b);
    }

    private void toggleEmail(boolean b)
    {
        //Emailing tab
        txtManualEmailName.setEnabled(b);
        txtManualEmailAddress.setEnabled(b);
        btnClearManualName.setEnabled(b);
        btnClearManualEmail.setEnabled(b);
        btnSendManualEmail.setEnabled(b);
        btnResetManualEmail.setEnabled(b);
        rdbtnEmailAsTxt.setEnabled(b);
        rdbtnEmailAsPic.setEnabled(b);
    }

    private void toggleSearch(boolean b)
    {
        btnSearch.setEnabled(b);
        txtSearch.setEnabled(b);
        btnScreenshotAll.setEnabled(b);
    }

    private void toggleCheckboxes(boolean b)
    {
        cbxFrontLine.setEnabled(b);
        cbxCashier.setEnabled(b);
        cbxOvernight.setEnabled(b);
    }

    private void uncheckCheckboxState()
    {
        cbxFrontLine.setSelected(false);
        cbxCashier.setSelected(false);
        cbxOvernight.setSelected(false);
    }

    private void displayInstructions()
    {
        JOptionPane.showMessageDialog(null,
                "Please complete and acknowledge the following before using this program:\n\n"
                        + "    1) Save the schedule for the week by department (All labor) as an Excel file (.xls).\n"
                        + "             - Conventionally, you should rename the Excel file as the week ending date.\n"
                        + "             - Open the Excel file after it has been saved successfully.\n"
                        + "             - You may be prompted with an error message stating that the Excel file is in\n"
                        + "               an incorrect format. Just click \"yes\"\n"
                        + "             - Click File, Save As, and change the file format by clicking the down arrow next to \"Save as type:\"\n"
                        + "               and select the option \"Excel 97-2003 Workbook (*.xls)\"\n"
                        + "             - Click the file that you saved and then click Save.\n"
                        + "             - You will be prompted if you would like to overwrite the original file. Click \"Yes\".\n"
                        + "             - You can now exit Excel.\n\n"
                        + "Using this program:\n"
                        + "    - Begin by clicking \"Open XLS\"\n"
                        + "    - Navigate to the Excel file (file_name.xls) previously saved.\n"
                        + "    - Open the Excel file by clicking \"open\"\n"
                        + "    - If the correct file was chosen the checkboxes will enable, allowing you to select schedule types\n"
                        + "      to be processed.\n"
                        + "    - At least one checkbox must be checked in order for the action buttons to enable. Once a checkbox\n"
                        + "      is checked, the action buttons will enable and you can select which day(s) to create a schedule for.\n"
                        + "    - Select any option by clicking the button designated for the desired operation.\n"
                        + "    - Click \"Reset\" to clear everything (console, file txtXLSPath, disable checkboxes and buttons)\n"
                        + "    - If you choose to reset everything you must choose another file in order to re-enable the checkboxes\n\n\n\n"
                        + "Created By: Brett Allen (443-812-0896)",
                "Instructions", JOptionPane.INFORMATION_MESSAGE
        );
    }

    public void setDayButtonTexts(ArrayList<String> truncDates)
    {
        String[] tokens;
        String date;
        String year = Integer.toString(Calendar.getInstance().get(Calendar.YEAR));

        for(int i = 0; i < truncDates.size(); i++)
        {
            tokens = truncDates.get(i).split("-");
            date = BJsWholesaleScheduler.transformDate(tokens, year);

            switch(i)
            {
                case 0:
                    btnSunday.setText("Sunday, " + date);
                    lblSunday.setText("Sunday, " + date);
                    break;
                case 1:
                    btnMonday.setText("Monday, " + date);
                    lblMonday.setText("Monday, " + date);
                    break;
                case 2:
                    btnTuesday.setText("Tuesday, " + date);
                    lblTuesday.setText("Tuesday, " + date);
                    break;
                case 3:
                    btnWednesday.setText("Wednesday, " + date);
                    lblWednesday.setText("Wednesday, " + date);
                    break;
                case 4:
                    btnThursday.setText("Thursday, " + date);
                    lblThursday.setText("Thursday, " + date);
                    break;
                case 5:
                    btnFriday.setText("Friday, " + date);
                    lblFriday.setText("Friday, " + date);
                    break;
                case 6:
                    btnSaturday.setText("Saturday, " + date);
                    lblSaturday.setText("Saturday, " + date);
                    break;
            }
        }
    }

    private void render(String day, Shift[][] myArray, int rows, int cols,
                        int choice, ArrayList<String> truncDates) throws BiffException, WriteException,
            IOException, BadLocationException
    {
        BJsWholesaleScheduler.renderChoice(day, myArray, rows, cols, choice, truncDates, menu);
    }

    public void print(String text) throws BadLocationException
    {
        Document doc = logArea.getDocument();
        logArea.getDocument().insertString(doc.getLength(), text, null);
    }

    /*Search for employee within provided schedule and fill in the employee's schedule
    * onto the search tab*/
    private boolean findEmployeeSchedule(boolean showMessages)
    {
        resetDays();

        if(!txtSearch.getText().equals("") && !txtSearch.getText().equals(""))
        {
            if(isEmployee(employeeSchedule, rows, cols, txtSearch.getText()))
            {
                txtSunday.setText(getEmployeeShift(lblSunday.getText(), txtEmployee.getText()));
                txtMonday.setText(getEmployeeShift(lblMonday.getText(), txtEmployee.getText()));
                txtTuesday.setText(getEmployeeShift(lblTuesday.getText(), txtEmployee.getText()));
                txtWednesday.setText(getEmployeeShift(lblWednesday.getText(), txtEmployee.getText()));
                txtThursday.setText(getEmployeeShift(lblThursday.getText(), txtEmployee.getText()));
                txtFriday.setText(getEmployeeShift(lblFriday.getText(), txtEmployee.getText()));
                txtSaturday.setText(getEmployeeShift(lblSaturday.getText(), txtEmployee.getText()));

                txtSearch.setSelectionEnd(160);

                //Reset txtSearch field and give it focus
                txtSearch.setText("");
                txtSearch.requestFocus();

                //txtSearch.setSelectionStart(0);


                if(showMessages)
                {
                    JOptionPane.showMessageDialog(null, "Successfully found " + txtEmployee.getText() + "!",
                            "OPERATION COMPLETE", JOptionPane.INFORMATION_MESSAGE);
                }
            }
            else
            {
                resetDays();

                if(showMessages)
                {
                    JOptionPane.showMessageDialog(null, txtSearch.getText() + " is either not an existing employee or does not work this week!",
                            "ERROR", JOptionPane.ERROR_MESSAGE);
                }

                resetSearchFields();

                return false;
            }
        }
        else
        {
            resetDays();

            if(showMessages)
            {
                JOptionPane.showMessageDialog(null, "Search string empty!", "ERROR", JOptionPane.ERROR_MESSAGE);
            }

            resetSearchFields();

            return false;
        }

        return true;
    }

    /*Find any existing screenshot files if the exist and place them in the screenshots ArrayList*/
    private void getScreenshotFiles()
    {
        /*Try to determine if a schedule is loaded. If there is a loaded schedule, determine the week ending.
        * Use the week ending as a way to specify which screenshots to email; in other words: don't send previous
        * week ending schedules to people if there is a different week ending schedule currently loaded*/

        //Make it so that if there are more screenshots than the number of email recipients (more than one week's
        //worth of screenshots) the user will be notified that they need to move or delete the week that they
        //do not want to send (program can only send one schedule at a time)

        String formattedName;

        //Create a list of all the files within the default directory that newly created files with be stored
        File projectFilePath = new File(new File("").getAbsolutePath());
        String[] fileList = projectFilePath.list();

        //Keep track of the number of screenshots to determine if the number of screenshots is greater than the
        //number of emails
        int numScreenshots = 0;

        //Vary processing based on the currentFileLoadState
        if(currentFileLoadState == FileLoadState.INIT)
        {
            //TODO Handle INIT State
        }
        else if (currentFileLoadState == FileLoadState.LOADED)
        {
            //TODO Handle LOADED State
        }
        else if(currentFileLoadState == FileLoadState.RESET)
        {
            //TODO Handle RESET State
        }

        //Traverse the email list and get the screenshots that may have been taken already for those recipients
        for(String recipient : emails)
        {
            formattedName = recipient.split(DEFAULT_SEPARATOR)[0].trim();

            //Traverse the fileList as long as there are files and add any screenshots that contain the name of
            // the employee from the email list
            if(fileList != null)
            {
                for(String fileName : fileList)
                {
                    if(fileName.contains(formattedName)) {
                        screenshots.add(formattedName + DEFAULT_SEPARATOR + fileName);
                        numScreenshots++;
                    }
                }
            }
        }

        System.out.println("Num screenshots: " + numScreenshots);
        System.out.println("Num emails: " + (emails.size()));

        //Log the screenshots that are available to send
        for(String screenshot : screenshots)
            System.out.println(screenshot.split(":")[1].trim() + " is ready to be sent.");
    }

    /*Take a screenshot of the search tab*/
    private void takeScreenshot(boolean showMessages)
    {
        Rectangle rect = getContentPane().getBounds();

        BufferedImage screenshot = new BufferedImage(rect.width, rect.height, BufferedImage.TYPE_INT_ARGB);

        getContentPane().paint(screenshot.getGraphics());

        String screenshotFileName = "Schedule for " + currentEmployee + " - Week Ending " + lblSaturday.getText() + ".png";

        try
        {
            ImageIO.write(screenshot, "png", new File(screenshotFileName));

            //Open the screenshot and crop it accordingly
            Image src = ImageIO.read(new File(screenshotFileName));

            //Dimensions of cropped image
            int x = 0, y = 100, w = 670, h = 502;

            BufferedImage croppedImage = new BufferedImage(w, h, BufferedImage.TYPE_INT_ARGB);
            croppedImage.getGraphics().drawImage(src, 0, 0, w, h, x, y, x + w, y + h, null);

            //Save the cropped screenshot in the same location in order to overwrite the original
            ImageIO.write(croppedImage, "png", new File(screenshotFileName));

            if(showMessages)
                JOptionPane.showMessageDialog(null, "Screenshot taken!", "OPERATION COMPLETE", JOptionPane.INFORMATION_MESSAGE);
        }
        catch(IOException e)
        {
            e.printStackTrace();
        }

        screenshots.add(currentEmployee + ": " + screenshotFileName);
    }

    /*Prompts the user for authentication*/
    private String getPassword()
    {
        JPasswordField passwordField = new JPasswordField();
        String passwordString;

        int confirmationCode = JOptionPane.showConfirmDialog(null, passwordField, AUTHENTICATION_MESSAGE,
                JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        if(confirmationCode == JOptionPane.OK_OPTION)
        {
            passwordString = new String(passwordField.getPassword());
            return passwordString;
        }

        return "";
    }

    private boolean sendScheduleAsPic()
    {
        boolean found = false;

        //Only require authentication one time
        if(password.equals(""))
            password = getPassword();

        if(!password.equals(""))
        {
            //Create an email object for sending the schedule as a picture
            Email email = new Email(password);

            //Find matching employee name in screenshots ArrayList
            for(String screenshot : screenshots)
            {
                System.out.println(screenshot);

                //Get the file path of the screenshot and pdf file (pdf file may not exist)
                String[] fileNames = new String[2];
                fileNames[0] = new File("").getAbsolutePath() + "\\" + screenshot.split(":")[1].trim();
                fileNames[1] = txtPDFPath.getText();

                String subject = screenshot.split(":")[1].split(".png")[0].trim();

                //Determine if the employee name within the screenshot file name matches txtManualEmailName field
                if(screenshot.split(":")[0].equals(txtManualEmailName.getText()))
                {
                    found = true;

                    System.out.println("Sending email for " + txtManualEmailName.getText() + " to "
                            + txtManualEmailAddress.getText() + "...");

                    try
                    {
                        email.sendAttachmentMessage(txtManualEmailAddress.getText(), txtManualEmailName.getText(), subject, fileNames);
                    }
                    catch(MessagingException me)
                    {
                        email.errorInEmailProcess = true;
                        found = false;

                        System.err.println("Messaging Exception was Thrown!");
                        JOptionPane.showMessageDialog(null, Email.DEV_INFO + me.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
                        me.printStackTrace();
                    }

                    break;
                }
            }
        }

        return found;
    }

    private void writeToEmailFile(EmailMode mode)
    {
        if(mode == EmailMode.ADDING)
        {
            emailingListModel.removeAllElements();

            emails.add(txtEmailName.getText() + DEFAULT_SEPARATOR + txtEmailAddress.getText());

            //Write to email preferences file
            emailFile.add(txtEmailName.getText(), txtEmailAddress.getText());

            txtEmailName.setText("");
            txtEmailAddress.setText("");

            //Traverse the emails ArrayList and write the contents to both the emailingListModel and emailOutFile
            for (String recipient : emails)
            {
                //noinspection unchecked
                emailingListModel.addElement(recipient);

                //Append each email entry to emails.txt file
                //emailOutFile.println(recipient);
            }
        }

        //Only delete the selected email
        if(mode == EmailMode.DELETING_SELECTED)
        {
            //Remove employee from emailList
            for(int i = 0; i < emails.size(); i++)
            {
                //Determine if the email entry contains the name of the employee
                if(emails.get(i).contains(emailSpliter[0]))
                {
                    emailFile.eraseEmail(emailSpliter[0].trim());
                    emails.remove(i);
                    emailingListModel.remove(i);
                    txtManualEmailName.setText("");
                    txtManualEmailAddress.setText("");
                }
            }
        }

        //Delete all emails in the email list
        if(mode == EmailMode.DELETING_ALL)
        {
            do
            {
                for(int i = 0; i < emails.size(); i++)
                {
                    emailFile.eraseEmail(emails.get(i).split(DEFAULT_SEPARATOR)[0].trim());
                    emails.remove(i);
                    emailingListModel.remove(i);
                }
            }
            while(!emails.isEmpty());

            JOptionPane.showMessageDialog(null, "Email list cleared.\n" +
                    "The app will exit after acknowledging this message.\n" +
                    "Please restart application to refresh the email list.", "IMPORTANT", JOptionPane.INFORMATION_MESSAGE);

            //Exit the app so the user can restart it.
            System.exit(0);
        }
    }

    private String getEmailMessage(String employeeName)
    {
        String message = "";

        message += lblSunday.getText() + " | " + getEmployeeShift(lblSunday.getText(), employeeName) + "\n\n";
        message += lblMonday.getText() + " | " + getEmployeeShift(lblMonday.getText(),  employeeName) + "\n\n";
        message += lblTuesday.getText() + " | " + getEmployeeShift(lblTuesday.getText(),  employeeName) + "\n\n";
        message += lblWednesday.getText() + " | " + getEmployeeShift(lblWednesday.getText(),  employeeName) + "\n\n";
        message += lblThursday.getText() + " | " + getEmployeeShift(lblThursday.getText(),  employeeName) + "\n\n";
        message += lblFriday.getText() + " | " + getEmployeeShift(lblFriday.getText(),  employeeName) + "\n\n";
        message += lblSaturday.getText() + " | " + getEmployeeShift(lblSaturday.getText(),  employeeName) + "\n\n";

        return message;
    }

    private String getEmailSubject()
    {
        return "Schedule for " + currentEmployee + " - Week Ending " + lblSaturday.getText();
    }

    /**Determines when to toggle the btnAddEmail and btnResetAddEmail based on if there
     * is text in both the txtEmailName and txtEmailAddress components.*/
    private void determineIfReady()
    {
        if(!txtEmailName.getText().trim().equals("") && !txtEmailAddress.getText().trim().equals(""))
        {
            btnAddEmail.setEnabled(true);
            btnResetAddEmail.setEnabled(true);
        }
        else
        {
            btnAddEmail.setEnabled(false);
            btnResetAddEmail.setEnabled(false);
        }
    }

    //Set the current order state to orderState and update the displayed order state
    private void setOrderState(SortedOrder orderState)
    {
        currentOrderState = orderState;
        lblSortedOrder.setText("Sorted Order: " + currentOrderState);
    }

    @Override
    public void actionPerformed(ActionEvent e)
    {
    }
}