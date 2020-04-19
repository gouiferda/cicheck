import java.awt.*;
import java.awt.Font;
import javax.swing.*;
import javax.swing.table.*;
import javax.swing.border.TitledBorder;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import processing.serial.*;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.File;
import java.io.IOException;
import java.io.FileInputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.nio.file.Files;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.format.Alignment;
import jxl.format.CellFormat;
import jxl.format.Colour;
import jxl.format.VerticalAlignment;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import javax.mail.*;
import javax.mail.internet.*;
import javax.activation.*;
import javax.swing.Timer;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.image.BufferedImage;
import javax.imageio.ImageIO;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;


//French
String lbl1 = "ciCHECK box", 
  lbl2 = "Connecté, En attente des cartes à scanner... ", 
  lbl3 = "Déconnecté, Veuillez connecter ciCHECK box à l'ordinateur!", 
  lbl4 = "Fichier de présence Excel", 
  lbl5 = "Gestion des cartes", 
  lbl6 = "Temps d'envoi:", 
  lbl7 = "Enregistrer", 
  lbl8 = "Vider champ", 
  lbl9 = "Envoyer fichier Excel", 
  lbl10 = "Auto envoi", 
  lbl11 = "Rapport de date", 
  lbl12 = "Fichier Excel de présence a ete envoyé à l'e-mail", 
  lbl13 = " tenter de glisser la carte 2 fois dans une courte période de temps.", 
  lbl14 = " tenter de glisser la carte plus de 2 fois.", 
  lbl15 = "Fichier Excel de présence est programmé pour être envoyé à ", 
  lbl16 = " à ", 
  lbl17 = "Fichier de présence Text", 
  lbl18 = "Nom complet (num carte):\t\t\t\tSortie/Entrée", 
  lblErrorTitle = "Erreur", 
  lbl20 = "Erreur en écrivant dans le fichier texte", 
  lbl21 = "Erreur en fermeture du fichier texte ", 
  lbl22 = "Erreur loadConfig ", 
  lbl23 = "Erreur saveConfig ", 
  lbl24 = "Impossible de traiter le message \n\n", 
  lbl25 = "Erreur en écrivant dans le fichier texte ", 
  lbl26 = "Erreur en fermeture du fichier texte", 
  lbl27 = "Impossible de charger les utilisateurs à partir de la base de données (fichier manquant ou inaccessible) usersTxtFilePath:", 
  lbl28 = "Impossible de créer le fichier Excel", 
  lbl29 = "ciCHECK Reports Service", 
  lbl30 = "ciCHECK fichier de présence", 
  lbl31 = "Bonjour\nCeci est un message automatique du service ciCHECK d'envoi des rapports de présence,\n\nn.b: Le fichier excel de présence est ci-joint, merci.\n\n", 
  lblMessage = "Message", 
  lbl33 = "Erreur d'envoie le mail!\nVérifiez l'email utilisé et votre connexion internet.\n\n", 
  lbl34 = "Authentification... ", 
  lbl35 = "Annuler envoi auto", 
  lbl36 = "Envoi automatique annulé", 
  lbl37 = "Saisir un email d'abords, et une heure valid", 
  lbl38 = "ciCHECK: Système de gestion du temps de présence", 
  lbl39 = "Nom complet", 
  lbl40 = "Sortie", 
  lbl41 = "Entrée", 
  lbl42 = "Événement de jour:", 
  lbl43 = "Email:", 
  lbl44 = "Activer autoSync", 
  lbl45 = "Activer autoSync", 
  lbl46 = "Desactiver autoSync", 
  lbl47 = "Il faut redemarer l'application pour voir les changements!", 
  lbl48 = "", 
  lbl49 = "", 
  lblcopyrightsTxt = "ciCHECK 2.0 (c) 2018 \n Réalisé par Soufiane Gouiferda, Grâce à Connect Institute \n https://github.com/notsoufiane/cicheck";

String frameTitle = "ciCHECK 2.0", txtFrameUpText="ciCHECK 2.0", txtFrameUpTime="Time here", cicheckBoxStatus = "Déconnecté", 
  txtTitlePanelLeft="Options generale", txtTitlePanelRight="Options d'envoi", 
  txtTitlePanelCenter="Tableau du presence", txtFrameDownStatus="Statu: deconnecter";
String dateOnly, homePathUser, homePathShared, fileNameTextRecord, fileNameExcel, appFilesPath, filePathTextRecord, filePathExcelAcc, 
  filePathExcel, recordFilePath = "", usersTxtFilePath, currentTimeFrameUp, currentTime = "", txtEmailTitle="Email:", 
  txtEventTitle="Événement de jour:", txtActionsTitle="Actions:", txtMailTime="", sendToMail = "", 
  newFullName = "", newSex = "", currentCardCode = "Unknown", txtAutoSync2="http://qadrine.eb2a.com/cicheck/", txtAutoSync1="Siteweb de Synchronisation:"
  ;

int frameWidth=1000, frameHeight=600, txtFieldsWidth=16;

JFrame frame;
JPanel panelMain, panelNorth, panelCenter, panelSouth, panelLeft, panelRight;
Font font1= new Font("Dialog", Font.PLAIN, 14);
;
JLabel lblFrameUpText, lblFrameDownTime, lblFrameDownStatus, lblEmailTitle, lblEventTitle, lblActionsTitle, 
  lblAutoSync1, lblAutoSync2, lblLogoConnect, lblLogocicheck, lblStatusIcon, lblIncomingCards;
JTextField txtEvent, txtEmail;
JButton btnSaveExcel, btnManageCards, btnSendEmail, btnNewCard;
JCheckBox cbAutoSend, cbAutoSync;
TitledBorder titlePanelLeft, titlePanelCenter, titlePanelRight;
DefaultTableModel tableModel1;
JTable tablePresence;

JMenuBar menubar = new JMenuBar();
JMenu menuFile = new JMenu("Fichier");
JMenu menuOptions = new JMenu("Options");
JMenu menuHelp = new JMenu("Aider");
JMenu menuLanguage = new JMenu("Langue");
JMenu menuVariables = new JMenu("Variables");


JMenuItem menuInSaveExcelFile = new JMenuItem("Enregistrer fichier excel");
JMenuItem menuInExit = new JMenuItem("Fermer");

JMenuItem menuLanguageEn = new JMenuItem("English");
JMenuItem menuLanguageFr = new JMenuItem("Français");
JMenuItem menuAllowedMinutesCycle = new JMenuItem("Période de minutes autorisée");
JMenuItem menuInReference = new JMenuItem("Reference");
JMenuItem menuInCheckUpdate = new JMenuItem("Mise a jour");
JMenuItem menuInAbout = new JMenuItem("A propos");

int portPad = -1;
Serial myPort;

public static Timer timerNewCard;
int timerSeconds = 10;
int timerMs = 1000*timerSeconds;
String msgTitleTimer = "Message";
String message3Timer= "Glisser la nouvelle carte de nouveau utilisateur maintenant";
final JOptionPane optionPaneTimer = new JOptionPane(message3Timer, JOptionPane.INFORMATION_MESSAGE, JOptionPane.DEFAULT_OPTION, null, new Object[] {}, null);
int timerCountDown = timerSeconds;

static ArrayList < Person > personsList = new ArrayList < Person > ();
static ArrayList < String > recordedList = new ArrayList < String > ();

int allowedMinutesCycle = 30; //in minutes

BufferedWriter output;

boolean autoSyncProgrammed=false, autoSendProgrammed=false;

public static WritableWorkbook wworkbook, workbookFrom, activeWorksheet;
public static WritableSheet wsheet;

final JDialog dialogWithTimer= new JDialog();

BufferedImage imgLogoConnect, imgLogocicheck, imgConnected, imgDisconnected, imgExcel, imgNewCard, imgCards, imgSendMail, imgScanned, imgDelete, imgSave, imgDown, imgUp;

JFrame frameManageCards;
JPanel panelManageCardsMain, panelManageCardsRight;
DefaultTableModel tableModelManageCards;
JTable tableManageCards;
JTextField txtManageFullName, txtManageCardCode, txtManageSex;
JButton btnManageSave, btnManageDelete, btnManageUp, btnManageDown;

int selectedRowManageCards = -1;

void setup() {
  frameRate(12);
  size(0, 0);
  surface.setVisible(false);


  dialogWithTimer.setTitle(msgTitleTimer);
  dialogWithTimer.setModal(true);
  dialogWithTimer.setContentPane(optionPaneTimer);
  dialogWithTimer.setLocationRelativeTo(null);
  dialogWithTimer.setModalityType(JDialog.ModalityType.MODELESS);
  dialogWithTimer.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
  dialogWithTimer.pack();
  //create timer to dispose of dialog after 5 seconds
  timerNewCard = new Timer(timerMs, new ActionListener() {
    @Override
      public void actionPerformed(ActionEvent ae) {
      dialogWithTimer.dispose();
      timerCountDown = timerSeconds;
    }
  }
  );
  timerNewCard.setRepeats(false); //the timer should only go off once


  dateOnly = "Date-" + day() + "-" + month() + "-" + year();
  homePathUser = System.getProperty("user.home");
  homePathShared = "/Users/Shared";
  fileNameTextRecord = "Record_" + dateOnly + ".txt";
  fileNameExcel = "Record_" + dateOnly + ".xls";
  appFilesPath = "/cicheckFiles";
  filePathTextRecord = appFilesPath + "/records/" + fileNameTextRecord;
  filePathExcel = appFilesPath + "/records/";


  frame = new JFrame(frameTitle);
  menuFile.add(menuInSaveExcelFile);
  menuFile.addSeparator();
  menuFile.add(menuInExit);
  menuOptions.add(menuLanguage);
  menuOptions.add(menuVariables);

  menuLanguage.add(menuLanguageEn);
  menuLanguage.addSeparator();
  menuLanguage.add(menuLanguageFr);
  menuVariables.add(menuAllowedMinutesCycle);
  menuHelp.add(menuInReference);
  menuHelp.addSeparator();
  menuHelp.add(menuInCheckUpdate);
  menuHelp.addSeparator();
  menuHelp.add(menuInAbout);
  menubar.add(menuFile);
  menubar.add(menuOptions);
  menubar.add(menuHelp);

  menuInSaveExcelFile.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      saveExcel();
    }
  }
  );
  menuInExit.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      System.exit(0);
    }
  }
  );
  menuInReference.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      openFile("https://github.com/notsoufiane/cicheck");
    }
  }
  );
  menuInCheckUpdate.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      popOut("Soon", lblMessage);
    }
  }
  );
  menuInAbout.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      popOut(lblcopyrightsTxt, lblMessage);
    }
  }
  );
  menuLanguageFr.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      //popOut("Soon", lblMessage);
    }
  }
  );
  menuLanguageEn.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      popOut("Soon", lblMessage);
    }
  }
  );
  menuAllowedMinutesCycle.addActionListener(new ActionListener() {
    public void actionPerformed(ActionEvent ev) {
      try {
        //println("allowedMinutesCycle1: "+allowedMinutesCycle);
        String newAllowedMinutesCycle = "";
        newAllowedMinutesCycle = JOptionPane.showInputDialog("Taper une période de minutes autorisée en Minutes (ex: 30) ");
        if (isValidString(newAllowedMinutesCycle)) {
          allowedMinutesCycle = Integer.parseInt(newAllowedMinutesCycle);
          saveConfig("allowedMinutesCycle", newAllowedMinutesCycle);
        }
      }
      catch(Exception ex) {
      }
    }
  }
  );

  frame.setJMenuBar(menubar);

  panelMain = new JPanel();
  panelNorth = new JPanel(new BorderLayout());
  panelCenter = new JPanel(new BorderLayout());
  panelSouth = new JPanel(new BorderLayout());
  panelLeft = new JPanel(new GridLayout(8, 1)); //rows //collums
  panelRight = new JPanel(new GridLayout(8, 1));

  lblFrameUpText = new JLabel(txtFrameUpText, SwingConstants.CENTER);
  //lblFrameUpText.setFont(font1);
  lblFrameDownTime= new JLabel(txtFrameUpTime);
  //lblFrameDownTime.setFont(font1);
  txtEvent = new JTextField("", txtFieldsWidth);



  btnSaveExcel = new JButton("Enregistrer fichier Excel");
  try {
    imgExcel = ImageIO.read(new File(getFullImgPath("excel.png")));
    btnSaveExcel.setIcon(new ImageIcon(imgExcel));
  }
  catch(Exception ex) {
  }
  btnSaveExcel.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) { 
      saveExcel();
    }
  } 
  );
  btnManageCards = new JButton(lbl5);
  try {
    imgCards = ImageIO.read(new File(getFullImgPath("cards.png")));
    btnManageCards.setIcon(new ImageIcon(imgCards));
  }
  catch(Exception ex) {
  }
  btnManageCards.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) { 
      //popOut("Apres modifier le fichier, Il faut redemarer l'application pour voir les changements!", lblMessage);
      //openFile(usersTxtFilePath);
      manageCards();
    }
  } 
  );
  btnNewCard = new JButton("Nouvelle carte");
  try {
    imgNewCard = ImageIO.read(new File(getFullImgPath("newcard.png")));
    btnNewCard.setIcon(new ImageIcon(imgNewCard));
  }
  catch(Exception ex) {
  }
  btnNewCard.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) { 
      saveNewUser();
    }
  } 
  );
  cbAutoSend = new JCheckBox(lbl10);
  cbAutoSend.addItemListener(new ItemListener() {
    public void itemStateChanged(ItemEvent e) {
      //println("cbAutoSync Checked? " + cbAutoSync.isSelected());
      if (cbAutoSend.isSelected()) {
        int dialogButton = JOptionPane.YES_NO_OPTION;
        int dialogResult = JOptionPane.showConfirmDialog (null, "Voullez-vous saisir une nouvelle heure d'auto envoi? (ancien: "+txtMailTime+"h)", lblMessage, dialogButton);
        if (dialogResult == JOptionPane.YES_OPTION) {
          String txtSendTime = "";
          txtSendTime=JOptionPane.showInputDialog("Taper une heure, exemple: 19:00"); 
          txtMailTime = txtSendTime;
        }


        if (!txtEmail.getText().equals("") && isValidTime(txtMailTime)) {
          cbAutoSend.setText(lbl10+" ("+txtMailTime + "h)");
          autoSendProgrammed = true;
          showMsg(lbl15 + txtEmail.getText() + lbl16 + txtMailTime + "h", lblMessage);
          saveConfig("email", txtEmail.getText());
          saveConfig("autoSendTime", txtMailTime);
          saveConfig("autoSendOperation", "1");
        } else {
          popOut(lbl37, lblMessage);
          cbAutoSend.setSelected(false);
          cbAutoSend.setText(lbl10);
        }
      } else {
        cbAutoSend.setText(lbl10);
        autoSendProgrammed = false;
        popOut(lbl36, "Message");
        saveConfig("autoSendTime", txtMailTime);
        saveConfig("autoSendOperation", "0");
      }
    }
  }
  );
  cbAutoSync = new JCheckBox("Auto Synchronisation");
  cbAutoSync.addItemListener(new ItemListener() {
    public void itemStateChanged(ItemEvent e) {
      //println("cbAutoSend Checked? " + cbAutoSend.isSelected());
      if (cbAutoSend.isSelected()) {
        autoSyncProgrammed = true;
        saveConfig("autoSyncOperation", "1");
      } else {
        autoSyncProgrammed = false;
        saveConfig("autoSyncOperation", "0");
      }
    }
  }
  );
  JButton btnSendEmail = new JButton("Envoyer email");
  try {
    imgSendMail = ImageIO.read(new File(getFullImgPath("mail.png")));
    btnSendEmail.setIcon(new ImageIcon(imgSendMail));
  }
  catch(Exception ex) {
  }
  btnSendEmail.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) { 
      sendExcelViaMail();
    }
  } 
  );
  txtEmail = new JTextField("email here", txtFieldsWidth);

  try {
    //println("path: "+getFullImgPath("logo.png"));
    imgLogoConnect = ImageIO.read(new File(getFullImgPath("logo.png")));
    imgLogocicheck= ImageIO.read(new File(getFullImgPath("cichecklogo.png")));
    imgConnected = ImageIO.read(new File(getFullImgPath("connected.png")));
    imgDisconnected = ImageIO.read(new File(getFullImgPath("disconnected.png")));
    imgScanned= ImageIO.read(new File(getFullImgPath("scanned.png")));

    lblLogoConnect = new JLabel(new ImageIcon(imgLogoConnect));
    lblLogocicheck= new JLabel(new ImageIcon(imgLogocicheck));
    lblStatusIcon= new JLabel(new ImageIcon(imgConnected));
  }
  catch(Exception ex) {
    println("Error setting up images "+ex.toString());
  }


  // panelNorth.add(lblFrameUpText, BorderLayout.CENTER);

  panelNorth.add(lblLogoConnect, BorderLayout.WEST);
  panelNorth.add(lblLogocicheck, BorderLayout.EAST);
  panelMain.setLayout(new BorderLayout());
  panelMain.add(panelNorth, BorderLayout.NORTH);
  panelMain.add(panelCenter, BorderLayout.CENTER);
  panelMain.add(panelSouth, BorderLayout.SOUTH);
  panelMain.add(panelLeft, BorderLayout.WEST);
  panelMain.add(panelRight, BorderLayout.EAST);
  titlePanelLeft= BorderFactory.createTitledBorder(txtTitlePanelLeft);
  panelLeft.setBorder(titlePanelLeft);
  titlePanelRight = BorderFactory.createTitledBorder(txtTitlePanelRight);
  panelRight.setBorder(titlePanelRight);
  titlePanelCenter = BorderFactory.createTitledBorder(txtTitlePanelCenter);
  panelCenter.setBorder(titlePanelCenter);
  lblEmailTitle  = new JLabel(txtEmailTitle);
  lblEventTitle = new JLabel(txtEventTitle);
  lblActionsTitle= new JLabel(txtActionsTitle);
  lblAutoSync1= new JLabel(txtAutoSync1);
  lblAutoSync2= new JLabel(txtAutoSync2);
  panelLeft.add(lblEventTitle);
  panelLeft.add(txtEvent);
  panelLeft.add(lblActionsTitle);
  panelLeft.add(btnSaveExcel);
  panelLeft.add(btnManageCards);
  panelLeft.add(btnNewCard);
  /*
  panelRight.add(lblAutoSync1);
   panelRight.add(lblAutoSync2);
   panelRight.add(cbAutoSync);
   */
  panelRight.add(lblEmailTitle);
  panelRight.add(txtEmail);
  panelRight.add(btnSendEmail);
  panelRight.add(cbAutoSend);


  lblFrameDownStatus= new JLabel(txtFrameDownStatus);


  panelSouth.add(lblFrameDownStatus, BorderLayout.CENTER);
  panelSouth.add(lblStatusIcon, BorderLayout.WEST);
  panelSouth.add(lblFrameDownTime, BorderLayout.EAST);



  // Here is to load the TableModel
  String[] columnNames = {"Nom complet", "Sortie", "Entree"};

  tableModel1 = new DefaultTableModel(columnNames, 0) {
    public boolean isCellEditable(int row, int column) {
      return false;
    }
  };

  tablePresence = new JTable(tableModel1);
  tablePresence.setSelectionModel(new ForcedListSelectionModel());
  panelCenter.add(new JScrollPane(tablePresence), BorderLayout.CENTER );



  setConnectedPort();

  loadPersonsListFromDB();
  loadConfig();
  recordFilePath = homePathShared + filePathTextRecord;
  deleteFileIfExists(recordFilePath);
  deleteFileIfExists(homePathUser + "/" + fileNameExcel);
  addAnyTextToRecord("\nDate:" + day() + "/" + month() + "/" + year());


  frame.add(panelMain);
  frame.setSize(frameWidth, frameHeight);
  frame.setLocationRelativeTo(null);
  frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
  frame.setVisible(true);


  frameManageCards = new JFrame("Gestion des cartes");
  frameManageCards.addWindowListener(new java.awt.event.WindowAdapter() {
    @Override
      public void windowClosing(java.awt.event.WindowEvent windowEvent) {
      popOut(lbl47, lblMessage);
    }
  }
  );
  panelManageCardsMain = new JPanel(new BorderLayout());
  String[] columnNamesManageCards = {"Nom complet", "code du carte", "Sexe"};
  tableModelManageCards = new DefaultTableModel(columnNamesManageCards, 0) {
    public boolean isCellEditable(int row, int column) {
      return false;
    }
  };
  tableManageCards = new JTable(tableModelManageCards);
  tableManageCards.setSelectionModel(new ForcedListSelectionModel());
  panelManageCardsMain.add(new JScrollPane(tableManageCards), BorderLayout.CENTER );


  panelManageCardsRight = new JPanel(new GridLayout(10, 1));
  txtManageFullName = new JTextField("", 14);
  txtManageCardCode= new JTextField("", 14);
  txtManageSex= new JTextField("", 14);

  txtManageFullName.getDocument().addDocumentListener(new DocumentListener() {
    public void changedUpdate(DocumentEvent e) {
      //showMsg("changedUpdate", "");
    }
    public void removeUpdate(DocumentEvent e) {
      UpdateManageTextFields();
    }
    public void insertUpdate(DocumentEvent e) {
      UpdateManageTextFields();
    }
  }
  );

  txtManageCardCode.getDocument().addDocumentListener(new DocumentListener() {
    public void changedUpdate(DocumentEvent e) {
      //showMsg("changedUpdate", "");
    }
    public void removeUpdate(DocumentEvent e) {
      UpdateManageTextFields();
    }
    public void insertUpdate(DocumentEvent e) {
      UpdateManageTextFields();
    }
  }
  );

  txtManageSex.getDocument().addDocumentListener(new DocumentListener() {
    public void changedUpdate(DocumentEvent e) {
      //showMsg("changedUpdate", "");
    }
    public void removeUpdate(DocumentEvent e) {
      UpdateManageTextFields();
    }
    public void insertUpdate(DocumentEvent e) {
      UpdateManageTextFields();
    }
  }
  );

  btnManageSave = new JButton("Enregistrer");
  try {
    imgSave = ImageIO.read(new File(getFullImgPath("save.png")));
    btnManageSave.setIcon(new ImageIcon(imgSave));
  }
  catch(Exception ex) {
  }
  btnManageDelete= new JButton("Supprimer");
  try {
    imgDelete = ImageIO.read(new File(getFullImgPath("delete.png")));
    btnManageDelete.setIcon(new ImageIcon(imgDelete));
  }
  catch(Exception ex) {
  }
  btnManageUp= new JButton("Haut");
  try {
    imgUp = ImageIO.read(new File(getFullImgPath("up.png")));
    btnManageUp.setIcon(new ImageIcon(imgUp));
  }
  catch(Exception ex) {
  }
  btnManageDown= new JButton("Bas");
  try {
    imgDown = ImageIO.read(new File(getFullImgPath("down.png")));
    btnManageDown.setIcon(new ImageIcon(imgDown));
  }
  catch(Exception ex) {
  }
  btnManageSave.setEnabled(false);
  btnManageDelete.setEnabled(false);
  btnManageUp.setEnabled(false);
  btnManageDown.setEnabled(false);

  tableManageCards.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
    public void valueChanged(ListSelectionEvent event) {
      // do some actions here, for example
      // print first column value from selected row
      try {
        selectedRowManageCards = tableManageCards.getSelectedRow();
        String selectedFullName="", selectedCardCode="", selectedSex="";
        selectedFullName= tableManageCards.getValueAt(selectedRowManageCards, 0).toString();
        selectedCardCode= tableManageCards.getValueAt(selectedRowManageCards, 1).toString();
        selectedSex= tableManageCards.getValueAt(selectedRowManageCards, 2).toString();
        txtManageFullName.setText(selectedFullName);
        txtManageCardCode.setText(selectedCardCode);
        txtManageSex.setText(selectedSex);

        if (isValidString(selectedFullName) && isValidString(selectedCardCode) && isValidString(selectedSex)) {

          btnManageDelete.setEnabled(true);
          btnManageUp.setEnabled(true);
          btnManageDown.setEnabled(true);
          if (isValidString(txtManageFullName.getText()) && isValidString(txtManageCardCode.getText()) && isValidString(txtManageSex.getText())) {
            btnManageSave.setEnabled(true);
          } else {
            btnManageSave.setEnabled(false);
          }
        } else {
          btnManageSave.setEnabled(false);
          btnManageDelete.setEnabled(false);
          btnManageUp.setEnabled(false);
          btnManageDown.setEnabled(false);
        }

        //println(tableManageCards.getValueAt(selectedRow, 0).toString());
      }
      catch(Exception ex) {
      }
    }
  }
  );

  btnManageSave.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) {
      int reply = JOptionPane.showConfirmDialog(null, "Etes-vous sure de modifier la carte?", lblMessage, JOptionPane.YES_NO_OPTION);
      if (reply == JOptionPane.YES_OPTION) {
        updateUserInDB();
      }
    }
  } 
  );


  btnManageDelete.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) {
      int reply = JOptionPane.showConfirmDialog(null, "Etes-vous sure de supprimer la carte?", lblMessage, JOptionPane.YES_NO_OPTION);
      if (reply == JOptionPane.YES_OPTION) {
        deleteUserInDB();
      }
    }
  } 
  );

  btnManageUp.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) {
      updatePositionOfSelected("up");
    }
  } 
  );

  btnManageDown.addActionListener(new ActionListener() { 
    public void actionPerformed(ActionEvent e) {
      updatePositionOfSelected("down");
    }
  } 
  );


  panelManageCardsRight.add(txtManageFullName);
  panelManageCardsRight.add(txtManageCardCode);
  panelManageCardsRight.add(txtManageSex);
  panelManageCardsRight.add(btnManageSave);
  panelManageCardsRight.add(btnManageDelete);
  panelManageCardsRight.add(btnManageUp);
  panelManageCardsRight.add(btnManageDown); //
  panelManageCardsRight.add(new JLabel("\tDerniere code carte entrant:"));
  lblIncomingCards = new JLabel("\tN/A");
  panelManageCardsRight.add(lblIncomingCards);





  frameManageCards.add(panelManageCardsMain, BorderLayout.CENTER );
  frameManageCards.add(panelManageCardsRight, BorderLayout.EAST );
  frameManageCards.setSize(700, 600);
  frameManageCards.setLocationRelativeTo(null);
  frameManageCards.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);



  loadManageCards();

  //end setup
}

String formatDateToStr(Date date) {
  String output="";
  try {
    output = new SimpleDateFormat("dd-MM-yyyy").format(date);
  }
  catch(Exception ex) {
  }
  return output;
}

boolean checkManageTextFields(JTextField tf) {
  if (!isValidString(tf.getText())) {
    return false;
  } else {
    return true;
  }
}

String updatePositionOfSelected(String movement) {
  try {
    int destination = selectedRowManageCards;
    switch(movement) {
    case "up":
      if ((selectedRowManageCards-1) >-1) {
        destination -=1;
      } else {
        return "";
      }
      break;
    case "down":
      //println("(selectedRowManageCards)+1 "+(selectedRowManageCards+1));
      //println("tableModelManageCards.getRowCount() "+(tableModelManageCards.getRowCount()));
      if ((selectedRowManageCards+1) < tableModelManageCards.getRowCount()) {
        destination +=1;
      } else {
        return "";
      }
      break;
    }

    if (selectedRowManageCards<=-1 || selectedRowManageCards+1>tableModelManageCards.getRowCount()) {
      return "";
    }

    //println("selectedRowManageCards: "+selectedRowManageCards);
    //println("Destination: "+destination);

    String oldSelectedFullName="", oldSelectedCardCode="", oldSelectedSex="";
    oldSelectedFullName= tableManageCards.getValueAt(selectedRowManageCards, 0).toString();
    oldSelectedCardCode= tableManageCards.getValueAt(selectedRowManageCards, 1).toString();
    oldSelectedSex= tableManageCards.getValueAt(selectedRowManageCards, 2).toString();

    String destinationFullName="", destinationCardCode="", destinationSex="";
    destinationFullName=tableManageCards.getValueAt(destination, 0).toString();
    destinationCardCode=tableManageCards.getValueAt(destination, 1).toString();
    destinationSex=tableManageCards.getValueAt(destination, 2).toString();

    tableModelManageCards.setValueAt(new String(oldSelectedFullName), destination, 0);
    tableModelManageCards.setValueAt(new String(oldSelectedCardCode), destination, 1);
    tableModelManageCards.setValueAt(new String(oldSelectedSex), destination, 2);

    tableModelManageCards.setValueAt(new String(destinationFullName), selectedRowManageCards, 0);
    tableModelManageCards.setValueAt(new String(destinationCardCode), selectedRowManageCards, 1);
    tableModelManageCards.setValueAt(new String(destinationSex), selectedRowManageCards, 2);

    tableManageCards.setRowSelectionInterval(destination, destination);
    saveCurrentManageCardsModel();

    return "";
  }
  catch(Exception ex) {
    println("Error updatePositionOfSelected "+ex.toString());
    return "";
  }
}

void UpdateManageTextFields() {
  try {
    if (checkManageTextFields(txtManageFullName) && checkManageTextFields(txtManageCardCode) &&
      checkManageTextFields(txtManageSex) && (txtManageSex.getText().equals("f") || txtManageSex.getText().equals("m"))) {
      btnManageSave.setEnabled(true);
    } else {
      btnManageSave.setEnabled(false);
    }
    if (selectedRowManageCards==-1) {
      btnManageSave.setEnabled(false);
      btnManageDelete.setEnabled(false);
      btnManageUp.setEnabled(false);
      btnManageDown.setEnabled(false);
    }
  }
  catch(Exception ex) {
    println("error UpdateManageTextFields "+ex.toString());
  }
}

boolean isValidString(Object str) {
  if (str==null) {
    return false;
  }
  if (str.toString().equals("")) {
    return false;
  }
  return true;
}

void deleteUserInDB() {

  try {
    tableModelManageCards.removeRow(selectedRowManageCards);
    txtManageFullName.setText("");
    txtManageCardCode.setText("");
    txtManageSex.setText("");
    btnManageSave.setEnabled(false);
    btnManageDelete.setEnabled(false);
    btnManageUp.setEnabled(false);
    btnManageDown.setEnabled(false);
    selectedRowManageCards=-1;
    saveCurrentManageCardsModel();
    // popOut(lbl47,lblMessage);
  }
  catch(Exception ex) {
    println("Error deleteUserInDB "+ex.toString());
  }
}

void saveCurrentManageCardsModel() {
  try {
    String newUsersFileContent="";
    if (tableModelManageCards.getRowCount() > 0) {
      for (int i = 0; i < tableModelManageCards.getRowCount(); i++) {
        newUsersFileContent+=tableModelManageCards.getValueAt(i, 0)+":"+tableModelManageCards.getValueAt(i, 1)+":"+tableModelManageCards.getValueAt(i, 2);
        if (i<tableModelManageCards.getRowCount()-1) {
          newUsersFileContent+="\n";
        }
      }
    }
    String[] list = {
      newUsersFileContent
    };
    if (isValidString(newUsersFileContent)) {
      saveStrings(usersTxtFilePath, list);
      //println("newUsersFileContent:\n\n"+newUsersFileContent);
    }
  }
  catch(Exception ex) {
    println("Error saveCurrentManageCardsModel "+ex.toString());
  }
}

void updateUserInDB() {
  try {
    String fullNameToSave, cardCodeToSave, sexToSave;
    fullNameToSave = txtManageFullName.getText();
    cardCodeToSave= txtManageCardCode.getText();
    sexToSave = txtManageSex.getText();
    if (isValidString(fullNameToSave) && isValidString(cardCodeToSave) && isValidString(sexToSave)) {
      tableManageCards.setValueAt(new String(fullNameToSave), selectedRowManageCards, 0 );
      tableManageCards.setValueAt(new String(cardCodeToSave), selectedRowManageCards, 1);
      tableManageCards.setValueAt(new String(sexToSave), selectedRowManageCards, 2 );
    }
    saveCurrentManageCardsModel();
    //popOut(lbl47,lblMessage);
  }
  catch(Exception ex) {
    println("Error updateUserInDB "+ex.toString());
  }
}


void loadManageCards() {
  try {
    loadPersonsListWithoutAffectingHours();
    if (tableModelManageCards.getRowCount() > 0) {
      for (int i = tableModelManageCards.getRowCount() - 1; i > -1; i--) {
        tableModelManageCards.removeRow(i);
      }
    }
    for (int i = 0; i < (personsList.size()); i++) {
      //println(i);
      Person p = personsList.get(i);
      Object obj[] = new Object[]{new String(p.fullName), new String(p.cardCode), new String(p.sex)};
      tableModelManageCards.addRow(obj);
    }
    tableManageCards.setModel(tableModelManageCards);
  }
  catch(Exception ex) {
    println("loadManageCards "+ex.toString());
  }
}

void prepExcelAcc() {
  try {
    String dateString = day() + "-" + month() + "-" + year();
    //println(dateString);
    Date date1 = new SimpleDateFormat("dd-MM-yyyy").parse(dateString);
    String dayOfWeek = new SimpleDateFormat("EEEE", Locale.ENGLISH).format(date1);
    dayOfWeek = dayOfWeek.substring(0, 1).toLowerCase() + dayOfWeek.substring(1);
    //println(dayOfWeek);
    Calendar cal = Calendar.getInstance();
    Date dateBefore1Day, dateBefore2Day, dateBefore3Day, dateBefore4Day, dateBefore5Day;
    switch(dayOfWeek) {
    case "monday":
      filePathExcelAcc = filePathExcel;

      break;
    case "tuesday":


      cal.setTime(date1);
      cal.add(Calendar.DATE, -1);
      dateBefore1Day = cal.getTime();
      String[] strArray = {formatDateToStr(dateBefore1Day), formatDateToStr(date1)};
      println(generateExcelAcc(strArray));

      break;
    case "wednesday":
      cal.setTime(date1);
      cal.add(Calendar.DATE, -1);
      dateBefore1Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -2);
      dateBefore2Day = cal.getTime();
      String[] strArray2 = {formatDateToStr(dateBefore2Day), formatDateToStr(dateBefore1Day), formatDateToStr(date1)};
      println(generateExcelAcc(strArray2));
      break;
    case "thursday":
      cal.setTime(date1);
      cal.add(Calendar.DATE, -1);
      dateBefore1Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -2);
      dateBefore2Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -3);
      dateBefore3Day = cal.getTime();
      String[] strArray3 = { formatDateToStr(dateBefore3Day), formatDateToStr(dateBefore2Day), formatDateToStr(dateBefore1Day), formatDateToStr(date1)};
      println(generateExcelAcc(strArray3));
      break;
    case "friday":
      cal.setTime(date1);
      cal.add(Calendar.DATE, -1);
      dateBefore1Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -2);
      dateBefore2Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -3);
      dateBefore3Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -4);
      dateBefore4Day = cal.getTime();
      String[] strArray4 = { formatDateToStr(dateBefore4Day), formatDateToStr(dateBefore3Day), formatDateToStr(dateBefore2Day), formatDateToStr(dateBefore1Day), formatDateToStr(date1)};
      println(generateExcelAcc(strArray4));
      break;
    case "saturday":
      cal.setTime(date1);
      cal.add(Calendar.DATE, -1);
      dateBefore1Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -2);
      dateBefore2Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -3);
      dateBefore3Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -4);
      dateBefore4Day = cal.getTime();
      cal.setTime(date1);
      cal.add(Calendar.DATE, -5);
      dateBefore5Day = cal.getTime();
      String[] strArray5 = { formatDateToStr(dateBefore5Day), formatDateToStr(dateBefore4Day), formatDateToStr(dateBefore3Day), formatDateToStr(dateBefore2Day), formatDateToStr(dateBefore1Day), formatDateToStr(date1)};
      println(generateExcelAcc(strArray5));
      break;
    }
  } 
  catch (Exception ex) {
    showMsg("Error prepExcelAcc "+ex.toString(), lblErrorTitle);
  }
}

void manageCards() {
  frameManageCards.setVisible(true);
}

String generateExcelAcc(String filesArray[]) {
  String output = "";
  output = Arrays.toString(filesArray);



  /*

   //Create the new workbook
   File file = new File(fileNameExcel);
   WorkbookSettings wbs = new WorkbookSettings();
   wbs.setLocale(new Locale("en", "EN"));
   
   
   workbookFrom = Workbook.createWorkbook(file, wbs);
   
   //Mark the existing worksheet
   activeWorksheet = workbookFrom.getActiveWorksheet();
   Cell a1 = activeWorksheet.getCell("A1");
   a1.setValue("First");
   //Cell object must be released explicitly
   a1.release();
   
   //Merge the current workbook with the existing one
   workbookFrom.mergeWorkbook(new File("Other.xls"));
   
   
   //  Add picture. Since interfaces returned by native peers are bound to current thread but not
   // to Ole thread we need to explicitly run the following action in Ole Message Loop
   application.getOleMessageLoop().doInvokeAndWait(new Runnable() {
   public void run() {
   
   //Get the active worksheet again because it has been changed after merge
   Worksheet activeWorksheet = workbook.getActiveWorksheet();
   
   //Get its native peer to address MS Excel Object model
   _Worksheet worksheet = activeWorksheet.getPeer();
   
   //Retrieve the Pictures object
   Variant unspecifiedParameter = Variant.createUnspecifiedParameter();
   Int32 locale = new Int32(LocaleID.LOCALE_USER_DEFAULT);
   IDispatch pictures = worksheet.pictures(unspecifiedParameter, locale);
   PicturesImpl iPictures = new PicturesImpl(pictures);
   
   //Insert the image
   File file = new File("logo.jpg");
   //You have to specify absolute path to image file
   BStr bStr = new BStr(file.getAbsolutePath());
   iPictures.insert(bStr,unspecifiedParameter);
   }
   });
   
   //  Save workbook. WORKBOOKNORMAL file format constant allows to save workbook in
   // MS Excel 2003 format embedding either MS Excel 2007 or MS Excel 2003
   workbook.saveAs(new File("Result.xls"), FileFormat.WORKBOOKNORMAL, true);
   
   //Close workbook and quit application
   workbook.close(false);
   
   */

  return output;
}

void moveUserToUp(String newPersonCardCode) {
  try {
    int newUserRow = -1;
    int destination = -1;
    newUserRow = tableModelManageCards.getRowCount()+1;
    for (int i = 0; i < (personsList.size()); i++) {
      Person p = personsList.get(i);
      if (p.sex.equals("m") && !p.cardCode.equals(newPersonCardCode)) {
        //println("p.cardCode: "+p.cardCode+"  , newPersonCardCode: "+newPersonCardCode);
        destination = i;
      }
    }
    destination+=2;
    //println("newUserRow: "+newUserRow);
    //println("destination: "+destination);
    //println("ok1");
    tableManageCards.setRowSelectionInterval(newUserRow-2, newUserRow-2);
    //println("tableManageCards.getSelectedRow(): "+tableManageCards.getSelectedRow());
    //println("getValueAt: "+tableManageCards.getValueAt(newUserRow-2,0));
    //println("ok2");
    for (int i=newUserRow-2; i>=destination; i--) {
      //println("i: "+i);
      updatePositionOfSelected("up");
    }
  }
  catch(Exception ex) {
    println("Error moveUserToUp "+ex.toString());
  }
}

boolean addNewUserToDB(String name, String sex, String cardCode) {
  try {
    String usersDbPath  = homePathShared + appFilesPath + "/database/users.txt";
    output = new BufferedWriter(new FileWriter(usersDbPath, true)); 
    String messageToSave = name +":"+cardCode +":"+sex+"\n";
    output.write(messageToSave);
    return true;
  } 
  catch (Exception ex) {
    showMsg("Error1 addNewUserToDB", lblErrorTitle);
    //e.printStackTrace();
    return false;
  } 
  finally {
    if (output != null) {
      try {
        output.close();
      } 
      catch (IOException e) {
        showMsg("Error2 addNewUserToDB", lblErrorTitle);
        return false;
      }
    }
  }
}

void saveNewUserNow() {
  try {
    //println("currentCardCode: "+currentCardCode+" ,getCardOwnerNameByCardID(currentCardCode): "+getCardOwnerNameByCardID(currentCardCode)+", fullName: "+newFullName+", sex: "+newSex);
    if (!currentCardCode.equals("Unknown") && getCardOwnerNameByCardID(currentCardCode).equals("Unknown") && !newFullName.equals("") && (newSex.equals("m") || newSex.equals("f"))) {
      if (addNewUserToDB(newFullName, newSex, currentCardCode)) {
        dialogWithTimer.dispose();
        popOut("Le nouveau utilisateur a ete ajoute avec succes!", "Message");
        loadManageCards();
        if (newSex.equals("m")) {
          moveUserToUp(currentCardCode);
        }
        currentCardCode="Unknown";
      }
    }
  } 
  catch (Exception e) {
    showMsg("Error newUserNow", lblErrorTitle);
  }
}

String getFullImgPath(String str) {
  return homePathShared + appFilesPath + "/assets/imgs/" + str;
}

boolean saveNewUser() {
  try {
    newFullName = JOptionPane.showInputDialog("Taper le nom complet d'utilisateur: ");
    if (newFullName.equals("")) {
      return false;
    }
    String[] values = {"Masculin", "Feminin"};
    Object selected = JOptionPane.showInputDialog(null, "Choisir le sexe d'utilisateur:\n", "Message", JOptionPane.DEFAULT_OPTION, null, values, "0");
    if ( selected != null ) {//null if the user cancels. 
      if (selected.toString().equals("Masculin")) {
        newSex = "m";
      } else if (selected.toString().equals("Feminin")) {
        newSex = "f";
      }
    }
    if (!newSex.equals("m") && !newSex.equals("f")) {
      return false;
    }

    //start timer to close JDialog as dialog modal we must start the timer before its visible
    timerNewCard.start();
    dialogWithTimer.setVisible(true);
    return false;
  } 
  catch (Exception e) {
    showMsg("Error saveNewUser "+e.toString(), lblErrorTitle);
    return false;
  }
}

boolean isValidTime(String timeStr) {
  if (!timeStr.matches("(?:[0-1][1-9]|2[0-4]):[0-5]\\d")) {
    return false;
  } else {
    return true;
  }
}
void deleteFileIfExists(String path) {
  try {
    //println("deleting: "+path);
    File file = new File(path);
    boolean result = Files.deleteIfExists(file.toPath());

    if (result) {
      //println("deleted old text file");
    } else {
      //println("failed to delete old text file path: "+path);
    }
  } 
  catch (Exception ex) {
    println(" deleteFileIfExists error ");
  }
}

void saveExcel() {
  try {
    generateExcelRecord();
    openFileNew(homePathUser + "/" + fileNameExcel);
    //prepExcelAcc();
  }
  catch(Exception ex) {
    println("Error deleteFileIfExists "+ex.toString());
  }
}

void openFile(String filePath) {
  try {
    Runtime.getRuntime().exec("open " + filePath);
  } 
  catch (Exception ex) {
  }
}

void openFileNew(String filePath) {
  try {
    Runtime.getRuntime().exec("open -n " + filePath);
  } 
  catch (Exception ex) {
  }
}



void runCommand(String command) {
  try {
    Runtime.getRuntime().exec(command);
  } 
  catch (Exception ex) {
  }
}


void loadConfig() {
  try {
    String[] lines = loadStrings(homePathShared + appFilesPath + "/database/config.txt");
    //println(homePathShared + appFilesPath + "/database/config.txt");
    for (int i = 0; i < lines.length; i++) {
      if (!lines[i].equals("")) {
        String lineParts[] = lines[i].split(";");
        String param = lineParts[0];
        String value = lineParts[1];
        switch (param) {
        case "email":
          txtEmail.setText(value);
          break;
        case "autoSendTime":
          txtMailTime = value;
          break;
        case "autoSendOperation":
          if (value.equals("1")) {
            cbAutoSend.setSelected(true);
          }
          break;
        case "autoSyncOperation":
          if (value.equals("1")) {
            //autoSync(1);
          }
          break;
        case "allowedMinutesCycle":
          allowedMinutesCycle=Integer.parseInt(value);
          break;
        }
      }
    }
  } 
  catch (Exception e) {
    showMsg(lbl22, lblErrorTitle);
  }
}

void loadPersonsListFromDB() {

  try {
    personsList.removeAll(personsList);
    usersTxtFilePath = homePathShared + appFilesPath + "/database/users.txt";
    String[] lines = loadStrings(usersTxtFilePath);
    if (tableModel1.getRowCount() > 0) {
      for (int i = tableModel1.getRowCount() - 1; i > -1; i--) {
        tableModel1.removeRow(i);
      }
    }
    for (int i = 0; i < lines.length; i++) {
      String line = lines[i];
      String[] parts = line.split(":");
      String part1 = parts[0];
      String part2 = parts[1];
      String part3 = parts[2];
      personsList.add(new Person(part1, part2, part3));
      Object obj[] = new Object[]{part1, "--:--", "--:--"};
      tableModel1.addRow(obj);
    }
  } 
  catch (Exception ex) {
    showMsg(lbl27 + usersTxtFilePath+" "+ex.toString(), lblErrorTitle);
  }
}

void loadPersonsListWithoutAffectingHours() {
  try {
    personsList.removeAll(personsList);
    usersTxtFilePath = homePathShared + appFilesPath + "/database/users.txt";
    String[] lines = loadStrings(usersTxtFilePath);
    if (tableModel1.getRowCount() > 0) {
      for (int i = tableModel1.getRowCount() - 1; i > -1; i--) {
        tableModel1.removeRow(i);
      }
    }
    for (int i = 0; i < lines.length; i++) {
      String line = lines[i];
      String[] parts = line.split(":");
      String part1 = parts[0];
      String part2 = parts[1];
      String part3 = parts[2];
      personsList.add(new Person(part1, part2, part3));
      String timeout="--:--", timein="--:--";
      if (!getTimeOutInByCardOwner(part2, "timeOut").equals("")) {
        timeout=getTimeOutInByCardOwner(part2, "timeOut");
      }
      if (!getTimeOutInByCardOwner(part2, "timeIn").equals("")) {
        timein=getTimeOutInByCardOwner(part2, "timeIn");
      }
      Object obj[] = new Object[]{part1, timeout, timein};
      tableModel1.addRow(obj);
    }
  } 
  catch (Exception ex) {
    showMsg("loadPersonsListWithoutAffectingHours "+ex.toString(), lblErrorTitle);
  }
}

void setConnectedPort() {
  try {
    portPad = (int) getConnectedPadPort();
    if (portPad != -1) {
      cicheckBoxStatus = lbl2;
      lblStatusIcon.setIcon(new ImageIcon(imgConnected));
    } else {
      cicheckBoxStatus = lbl3;
      lblStatusIcon.setIcon(new ImageIcon(imgDisconnected));
    }
    lblFrameDownStatus.setText(lbl1 + ":\t" + cicheckBoxStatus);
    myPort = new Serial(this, Serial.list()[portPad], 9600);
  } 
  catch (Exception ex) {
    //showMsg("Error connecting pad",lblErrorTitle);
    //println("setConnectedPort "+ex.toString());
  }
}

static int getConnectedPadPort() {
  try {
    int i = 0;
    int portNum = -1;
    for (String a : Serial.list()) {
      //println(a);  // Will invoke overrided `toString()` method
      if (Serial.list()[i].indexOf("cu.wchu") != -1) {
        portNum = i;
      }
      i++;
    }
    return portNum;
  }
  catch(Exception ex) {
    println("Error setConnectedPort "+ex.toString());
    return -1;
  }
}

void serialEvent(Serial p) {
  // get message till line break (ASCII > 13)
  String message = myPort.readStringUntil(13);
  if (message != null) {
    //println("serialEvent: " + getGoodMessageFromSerial(message));
    lblStatusIcon.setIcon(new ImageIcon(imgScanned));
    processIncomingMessage(getGoodMessageFromSerial(message));
  }
}

String getGoodMessageFromSerial(String msg) {
  return msg.replaceAll("[^\\d.]", "");
}

void popOut(String msg, String title) {
  JOptionPane.showMessageDialog(null, msg, title, JOptionPane.INFORMATION_MESSAGE);
  //showMsg("popOut: "+msg, title);
}

void processIncomingMessage(String msg) {
  try {
    currentCardCode = msg;
    String cardIdPlane = msg;
    String fulltime = hour() + ":" + minute() + ":" + second();
    String cardCode = "", timeFound = "";
    SimpleDateFormat format = new SimpleDateFormat("HH:mm", Locale.FRENCH);
    String currentTime3 = format.format(Calendar.getInstance().getTime());
    int howMuchThereIsOfCardRec = 0;
    for (int i = 0; i < (recordedList.size()); i++) {
      String[] recordParts = recordedList.get(i).toString().split("-");
      cardCode = recordParts[0];
      //println("cardIdPlane: " + cardIdPlane+" ,timeFound: "+timeFound);
      if (cardCode.equals(msg)) {
        howMuchThereIsOfCardRec++;
        if (howMuchThereIsOfCardRec == 1) {
          timeFound = recordParts[1];
          //remove seconds part
          String[] timeFoundParts = timeFound.split(":");
          timeFound = timeFoundParts[0] + ":" + timeFoundParts[1] + "";
        }
      }
    }
    switch (howMuchThereIsOfCardRec) {
    case 0:
      recordedList.add(cardIdPlane + "-" + fulltime);
      lblIncomingCards.setText("\t"+cardIdPlane+" - Heure: "+fulltime);
      addAnyTextToRecord(cardIdPlane + ":" + fulltime);
      updateTable(cardIdPlane, fulltime, "in");
      break;
    case 1:
      Date dateFound = format.parse(timeFound);
      Date dateNow = format.parse(currentTime3);
      long difference = dateNow.getTime() - dateFound.getTime();
      String howMuchMinutesSpent = ((difference / (1000 * 60)) - (difference / (1000 * 60 * 60))) + "";
      //println("howMuchMinutesSpent: "+howMuchMinutesSpent);
      if (Integer.parseInt(howMuchMinutesSpent) >= allowedMinutesCycle) {
        recordedList.add(cardIdPlane + "-" + fulltime);
        addAnyTextToRecord(cardIdPlane + ":" + fulltime);
        updateTable(cardIdPlane, fulltime, "out");
      } else {
        showMsg(getCardOwnerNameByCardID(msg) + " (" + msg + ")" + lbl13, "Message");
        popOut(getCardOwnerNameByCardID(msg) + " (" + msg + ")" + lbl13, "Message");
      }
      break;
    case 2:
      showMsg(getCardOwnerNameByCardID(msg) + " (" + msg + ")" + lbl14, "Message");
      popOut(getCardOwnerNameByCardID(msg) + " (" + msg + ")" + lbl14, "Message");
      break;
    }
    showRecordedCards();
  }  
  catch (Exception ex) {
    showMsg(lbl24 + " " + ex.toString(), lblErrorTitle);
  }
}

void updateTable(String cardIdPlane, String fulltime, String inOrOut) {
  try {
    String name;
    int positionInLine = -1;
    switch(inOrOut) {
    case "in":
      positionInLine = 2;
      break;
    case "out":
      positionInLine = 1;
      break;
    }
    if (tableModel1.getRowCount() > 0) {
      for (int i = tableModel1.getRowCount() - 1; i > -1; i--) {
        name = tableModel1.getValueAt(i, 0).toString();
        if (getCardOwnerNameByCardID(cardIdPlane).equals(name)) {
          tableModel1.setValueAt(new String(fulltime), i, positionInLine);
        }
      }
    }
    //println("Updated time for "+getCardOwnerNameByCardID(cardIdPlane));
  }  
  catch (Exception ex) {
    showMsg("Erreur updateTable " + ex.toString(), lblErrorTitle);
  }
}

void showRecordedCards() {
  try {
    for (int i = 0; i < (recordedList.size()); i++) {
      String[] recordParts = recordedList.get(i).toString().split("-");
      println(recordParts[0]+" - "+recordParts[1]);
    }
  }
  catch(Exception ex) {
    println("Error showRecordedCards "+ex.toString());
  }
}

String getCardOwnerNameByCardID(String cardID) {
  try {
    String ownerName = "Unknown";
    for (Person person : personsList) {
      String cardIdGotFromList = person.getCardCode() + "";
      if (cardIdGotFromList.indexOf(cardID) != -1) {
        ownerName = person.getFullName();
      }
    }
    return ownerName;
  }
  catch(Exception ex) {
    println("Error getCardOwnerNameByCardID "+ex.toString());
    return "Unknown";
  }
}

void addAnyTextToRecord(String txt) {
  try {
    output = new BufferedWriter(new FileWriter(homePathShared + filePathTextRecord, true)); //the true will append the new data
    txt = txt + "\n";
    output.write(txt);
  } 
  catch (IOException e) {
    showMsg(lbl20, lblErrorTitle);
    //e.printStackTrace();
  } 
  finally {
    if (output != null) {
      try {
        output.close();
      } 
      catch (IOException e) {
        showMsg(lbl21, lblErrorTitle);
      }
    }
  }
}

void showMsg(String Msg, String Title) {
  println(Msg);
  //javax.swing.JOptionPane.showMessageDialog ( null, Msg, Titel, javax.swing.JOptionPane.INFORMATION_MESSAGE  );
}

void uploadFileFTP(String filePath) {

  String server = ""; //replace with your ftp server
  int port = 21;
  String user = "";  //replace with your username
  String pass = ""; //replace with your password
  FTPClient ftpClient = new FTPClient();
  try {
    ftpClient.connect(server, port);
    ftpClient.login(user, pass);
    ftpClient.enterLocalPassiveMode();
    ftpClient.setFileType(FTP.BINARY_FILE_TYPE);

    File firstLocalFile = new File(filePath);
    String firstRemoteFile = "/htdocs/cicheck/records/" + fileNameTextRecord;

    InputStream inputStream = new FileInputStream(firstLocalFile);
    println("Syncing: Start uploading file");
    boolean done = ftpClient.storeFile(firstRemoteFile, inputStream);
    inputStream.close();
    if (done) {
      println("Syncing: The file is uploaded successfully.");
    } else {
      println("Syncing failed!");
    }
  } 
  catch (IOException ex) {
    showMsg("Error1 uploadFileFTP " + ex.toString(), lblErrorTitle);
  } 
  finally {
    try {
      if (ftpClient.isConnected()) {
        ftpClient.logout();
        ftpClient.disconnect();
      }
    } 
    catch (IOException ex) {
      showMsg("Error2 uploadFileFTP " + ex.toString(), lblErrorTitle);
    }
  }
}

public static void addCell(WritableSheet sheet, Label lbl, int col, int row, Alignment alignment, Colour c, Colour c2) throws WriteException {
  try {
    WritableFont arial14font = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
    arial14font.setColour(c2);

    WritableCellFormat cellFormat = new WritableCellFormat(arial14font);
    cellFormat.setAlignment(alignment);

    cellFormat.setBackground(c);
    cellFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
    sheet.addCell(new Label(col, row, lbl.getString(), cellFormat));
  }
  catch (Exception ex) {
    println("Error addCell "+ex.toString());
  }
}

public static void addCellNormalText(WritableSheet sheet, Label lbl, int col, int row, Alignment alignment, Colour c, Colour c2) throws WriteException {
  try {
    WritableFont arial14font = new WritableFont(WritableFont.ARIAL, 14, WritableFont.NO_BOLD);
    arial14font.setColour(c2);

    WritableCellFormat cellFormat = new WritableCellFormat(arial14font);
    cellFormat.setAlignment(alignment);

    cellFormat.setBackground(c);
    cellFormat.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
    sheet.addCell(new Label(col, row, lbl.getString(), cellFormat));
  }
  catch (Exception ex) {
    println("Error addCellNormalText "+ex.toString());
  }
}

String getTimeOutInByCardOwner(String cardNum, String outputType) {
  String output = "";
  String timeIn = "";
  String timeOut = "";
  boolean savedTimeIn = false;

  try {
    for (int i = 0; i < (recordedList.size()); i++) {
      String[] parts = recordedList.get(i).toString().split("-");

      if ((cardNum.equals(parts[0])) && !savedTimeIn) {
        timeIn = parts[1];
        savedTimeIn = true;
      } else if ((cardNum.equals(parts[0])) && savedTimeIn) {
        timeOut = parts[1];
        savedTimeIn = false;
      }
    }
    switch (outputType) {
    case "timeIn":
      output = timeIn;
      break;
    case "timeOut":
      output = timeOut;
      break;
    }
    return output;
  }
  catch (Exception ex) {
    println("Error getTimeOutInByCardOwner "+ex.toString());
    return "";
  }
}


void expandColumn(WritableSheet sheet, int amountOfColumns) {
  try {
    int c = amountOfColumns;
    for (int x = 0; x < c; x++) {
      CellView cell = sheet.getColumnView(x);
      cell.setAutosize(true);
      sheet.setColumnView(x, cell);
    }
  }
  catch (Exception ex) {
    println("Error");
  }
}

void sheetAutoFitColumns(WritableSheet sheet) {
  try {
    for (int i = 0; i < sheet.getColumns(); i++) {
      Cell[] cells = sheet.getColumn(i);
      int longestStrLen = -1;

      if (cells.length == 0)
        continue;

      /* Find the widest cell in the column. */
      for (int j = 0; j < cells.length; j++) {
        if (cells[j].getContents().length() > longestStrLen) {
          String str = cells[j].getContents();
          if (str == null || str.isEmpty())
            continue;
          longestStrLen = str.trim().length();
        }
      }

      /* If not found, skip the column. */
      if (longestStrLen == -1)
        continue;

      /* If wider than the max width, crop width */
      if (longestStrLen > 255)
        longestStrLen = 255;

      CellView cv = sheet.getColumnView(i);
      cv.setSize(longestStrLen * 256 + 100); /* Every character is 256 units wide, so scale it. */
      sheet.setColumnView(i, cv);
    }
  }
  catch (Exception ex) {
    println("Error");
  }
}

void createExcelRecord() {
  try {
    //creating
    // Initial settings
    //File file = new File( filePathExcel,fileNameExcel );  
    File file = new File(fileNameExcel);
    WorkbookSettings wbs = new WorkbookSettings();
    wbs.setLocale(new Locale("en", "EN"));

    // Creates the workbook
    wworkbook = Workbook.createWorkbook(file, wbs);

    wsheet = wworkbook.createSheet("Semaine 1", 0);

    //adding table structure
    String dateString = day() + "/" + month() + "/" + year();
    Date date2 = new SimpleDateFormat("d/M/yyyy").parse(dateString);
    String dayOfWeek = new SimpleDateFormat("EEEE", Locale.FRENCH).format(date2);
    int row = 0;
    dayOfWeek = dayOfWeek.substring(0, 1).toUpperCase() + dayOfWeek.substring(1);
    Label label = new Label(0, row, dayOfWeek + " " + dateString + " " + txtEvent.getText() + " (" + (recordedList.size() / 2) + ")");
    addCell(wsheet, label, 0, 0, Alignment.CENTRE, Colour.GREEN, Colour.WHITE);
    wsheet.mergeCells(0, row, 14, row);

    addCell(wsheet, new Label(2, 1, "P22"), 2, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(3, 1, "Sortie"), 3, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(4, 1, "Entrée"), 4, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(5, 1, "Déj."), 5, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(6, 1, "Propreté"), 6, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);


    for (int i = 3; i < (personsList.size() + 3); i++) {
      //println(i);
      Person p = personsList.get(i - 3);
      //addCellNormalText(wsheet,new Label(0,i,(i-2)+""),0, i,Alignment.CENTRE,jxl.format.Colour.WHITE,Colour.BLACK);

      //println(p.getFullName());
      //println(p.getFullName()+" timeIn: "+getTimeOutInByCardOwner(p.getFullName(),"timeIn")+" , timeOut: "+getTimeOutInByCardOwner(p.getFullName(),"timeOut"));

      wsheet.addCell(new Number(0, i, (i - 2)));
      addCellNormalText(wsheet, new Label(1, i, p.getFullName()), 1, i, Alignment.CENTRE, Colour.ICE_BLUE, Colour.BLACK);


      String timeInToSave = getTimeOutInByCardOwner(p.getCardCode(), "timeIn");
      String timeOutToSave = getTimeOutInByCardOwner(p.getCardCode(), "timeOut");

      if (!timeInToSave.equals("") && !timeOutToSave.equals("")) {

        String howMuchHoursSpent = "XX";
        String howMuchMinutesLeftSpent = "XX";

        //remove seconds part
        String[] timeInParts = timeInToSave.split(":");
        timeInToSave = timeInParts[0] + ":" + timeInParts[1] + "";
        String[] timeOutParts = timeOutToSave.split(":");
        timeOutToSave = timeOutParts[0] + ":" + timeOutParts[1] + "";
        //println(timeInToSave + " "+ timeOutToSave );


        SimpleDateFormat format = new SimpleDateFormat("HH:mm", Locale.FRENCH);
        Date date1ToComp = format.parse(timeInToSave);
        Date date2ToComp = format.parse(timeOutToSave);
        long difference = date2ToComp.getTime() - date1ToComp.getTime();


        howMuchHoursSpent = (difference / (1000 * 60)) + "";
        howMuchMinutesLeftSpent = ((difference / (1000 * 60)) - (difference / (1000 * 60 * 60))) + "";
        //println(howMuchHoursSpent+":"+howMuchMinutesLeftSpent+":00");


        addCellNormalText(wsheet, new Label(2, i, howMuchHoursSpent), 2, i, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);

        timeInToSave += ":00";
        timeOutToSave += ":00";
      }

      addCellNormalText(wsheet, new Label(3, i, timeOutToSave), 3, i, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
      addCellNormalText(wsheet, new Label(4, i, timeInToSave), 4, i, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
    }

    //wsheet.addCell(new Number(3, 4, 1234));

    expandColumn(wsheet, 20);
    wworkbook.write();
    wworkbook.close();

    //reading
    /*
  Workbook workbook = Workbook.getWorkbook(new File(filePathExcel));
     Sheet sheet = workbook.getSheet(0);
     Cell cell1 = sheet.getCell(0, 2);
     System.out.println(cell1.getContents());
     Cell cell2 = sheet.getCell(3, 4);
     System.out.println(cell2.getContents());
     workbook.close();
     */
  } 
  catch (Exception ex) {
    showMsg(lbl28+" Erreur: "+ex.toString(), lblErrorTitle);
  }
}

void generateExcelRecord() {
  try {
    deleteFileIfExists(homePathUser + "/" + fileNameExcel);
    File file = new File(fileNameExcel);
    WorkbookSettings wbs = new WorkbookSettings();
    wbs.setLocale(new Locale("fr", "FR"));
    wworkbook = Workbook.createWorkbook(file, wbs);
    wsheet = wworkbook.createSheet("Semaine Experimental", 0);
    //adding table structure
    String dateString = day() + "/" + month() + "/" + year();
    Date date2 = new SimpleDateFormat("d/M/yyyy").parse(dateString);
    String dayOfWeek = new SimpleDateFormat("EEEE", Locale.FRENCH).format(date2);
    int row = 0;
    dayOfWeek = dayOfWeek.substring(0, 1).toUpperCase() + dayOfWeek.substring(1);
    Label label = new Label(0, row, dayOfWeek + " " + dateString + " " + txtEvent.getText() + " (" + (recordedList.size() / 2) + ")");
    addCell(wsheet, label, 0, 0, Alignment.CENTRE, Colour.GREEN, Colour.WHITE);
    wsheet.mergeCells(0, row, 14, row);

    addCell(wsheet, new Label(2, 1, "P22"), 2, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(3, 1, "Sortie"), 3, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(4, 1, "Entrée"), 4, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);

    addCell(wsheet, new Label(8, 1, "P22"), 8, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(9, 1, "Sortie"), 9, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    addCell(wsheet, new Label(10, 1, "Entrée"), 10, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    //addCell(wsheet, new Label(5, 1, "Déj."), 5, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);
    //addCell(wsheet, new Label(6, 1, "Propreté"), 6, 1, Alignment.CENTRE, jxl.format.Colour.VERY_LIGHT_YELLOW, Colour.BLACK);


    for (int i = 3; i < (personsList.size() + 3); i++) {
      Person p = personsList.get(i - 3);
      if (p.getSex().equals("m")) {
        //println("p.getFullName(): "+p.getFullName());
        // println("i: "+i);
        wsheet.addCell(new Number(0, i, (i - 2)));
        addCellNormalText(wsheet, new Label(1, i, p.getFullName()), 1, i, Alignment.CENTRE, Colour.PALE_BLUE, Colour.BLACK);
        String timeInToSave = getTimeOutInByCardOwner(p.getCardCode(), "timeIn");
        String timeOutToSave = getTimeOutInByCardOwner(p.getCardCode(), "timeOut");
        if (!timeInToSave.equals("") && !timeOutToSave.equals("")) {
          String howMuchHoursSpent = "XX";
          String howMuchMinutesLeftSpent = "XX";
          //remove seconds part
          String[] timeInParts = timeInToSave.split(":");
          timeInToSave = timeInParts[0] + ":" + timeInParts[1] + "";
          String[] timeOutParts = timeOutToSave.split(":");
          timeOutToSave = timeOutParts[0] + ":" + timeOutParts[1] + "";
          SimpleDateFormat format = new SimpleDateFormat("HH:mm", Locale.FRENCH);
          Date date1ToComp = format.parse(timeInToSave);
          Date date2ToComp = format.parse(timeOutToSave);
          long difference = date2ToComp.getTime() - date1ToComp.getTime();
          howMuchHoursSpent = (difference / (1000 * 60)) + "";
          howMuchMinutesLeftSpent = ((difference / (1000 * 60)) - (difference / (1000 * 60 * 60))) + "";
          addCellNormalText(wsheet, new Label(2, i, howMuchHoursSpent), 2, i, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
          timeInToSave += ":00";
          timeOutToSave += ":00";
        }
        addCellNormalText(wsheet, new Label(3, i, timeOutToSave), 3, i, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
        addCellNormalText(wsheet, new Label(4, i, timeInToSave), 4, i, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
      }
    }

    int j=0;
    for (int i = 3; i < (personsList.size() + 3); i++) {

      Person p = personsList.get(i - 3);
      if (p.getSex().equals("m")) {
        j++;
      }
      if (p.getSex().equals("f")) {
        //println(i-j);
        wsheet.addCell(new Number(6, i-j, (i - 2-j)));
        addCellNormalText(wsheet, new Label(7, i-j, p.getFullName()), 7, i-j, Alignment.CENTRE, Colour.CORAL, Colour.BLACK);
        String timeInToSave = getTimeOutInByCardOwner(p.getCardCode(), "timeIn");
        String timeOutToSave = getTimeOutInByCardOwner(p.getCardCode(), "timeOut");
        if (!timeInToSave.equals("") && !timeOutToSave.equals("")) {
          String howMuchHoursSpent = "XX";
          String howMuchMinutesLeftSpent = "XX";
          //remove seconds part
          String[] timeInParts = timeInToSave.split(":");
          timeInToSave = timeInParts[0] + ":" + timeInParts[1] + "";
          String[] timeOutParts = timeOutToSave.split(":");
          timeOutToSave = timeOutParts[0] + ":" + timeOutParts[1] + "";
          SimpleDateFormat format = new SimpleDateFormat("HH:mm", Locale.FRENCH);
          Date date1ToComp = format.parse(timeInToSave);
          Date date2ToComp = format.parse(timeOutToSave);
          long difference = date2ToComp.getTime() - date1ToComp.getTime();
          howMuchHoursSpent = (difference / (1000 * 60)) + "";
          howMuchMinutesLeftSpent = ((difference / (1000 * 60)) - (difference / (1000 * 60 * 60))) + "";
          addCellNormalText(wsheet, new Label(8, i-j, howMuchHoursSpent), 8, i-j, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
          timeInToSave += ":00";
          timeOutToSave += ":00";
        }
        addCellNormalText(wsheet, new Label(9, i-j, timeOutToSave), 9, i-j, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
        addCellNormalText(wsheet, new Label(10, i-j, timeInToSave), 10, i-j, Alignment.CENTRE, Colour.WHITE, Colour.BLACK);
      }
    }

    expandColumn(wsheet, 20);
    wworkbook.write();
    wworkbook.close();
  } 
  catch (Exception ex) {
    showMsg(lbl28+" Erreur: "+ex.toString(), lblErrorTitle);
  }
}

void sendExcelViaMail() {
  try {
    saveConfig("email", txtEmail.getText());
    generateExcelRecord();
    //openFile(homePathOrigin+"/"+fileNameExcel);
    sendToMail = txtEmail.getText();
    sendAttachementMail();
  }
  catch (Exception ex) {
    println("Error");
  }
}

void saveConfig(String paramToSave, String valueToSave) {
  String configFilePath = homePathShared + appFilesPath + "/database/config.txt";
  String fullNewConfigText = "", email = "", autosendtime = "", autosendoperation = "", param = "", value = "", autosyncoperation = "", allowedminutescycle="";
  try {
    String[] lines = loadStrings(configFilePath);
    for (int i = 0; i < lines.length; i++) {
      if (!lines[i].equals("")) {
        String lineParts[] = lines[i].split(";");
        param = lineParts[0];
        value = lineParts[1];

        switch (param) {
        case "email":
          email = value;
          break;
        case "autoSendTime":
          autosendtime = value;
          break;
        case "autoSendOperation":
          autosendoperation = value;
          break;
        case "autoSyncOperation":
          autosyncoperation = value;
          break;
        case "allowedMinutesCycle":
          allowedminutescycle = value;
          break;
        }
      }
    }
    switch (paramToSave) {
    case "email":
      email = valueToSave;
      break;
    case "autoSendTime":
      autosendtime = valueToSave;
      break;
    case "autoSendOperation":
      autosendoperation = valueToSave;
      break;
    case "autoSyncOperation":
      autosyncoperation = valueToSave;
      break;
    case "allowedMinutesCycle":
      allowedminutescycle = valueToSave;
      break;
    }

    fullNewConfigText += "email;" + email + "\n";
    fullNewConfigText += "autoSendTime;" + autosendtime + "\n";
    fullNewConfigText += "autoSendOperation;" + autosendoperation + "\n";
    fullNewConfigText += "autoSyncOperation;" + autosyncoperation + "\n";
    fullNewConfigText += "allowedMinutesCycle;" + allowedminutescycle + "\n";

    String[] list = {
      fullNewConfigText
    };
    saveStrings(configFilePath, list);
  } 
  catch (Exception e) {
    showMsg(lbl23, lblErrorTitle);
  }
}


void sendAttachementMail() {
  // Create a session
  String host = "smtp.gmail.com";
  Properties props = new Properties();

  // SMTP Session
  props.put("mail.transport.protocol", "smtp");
  props.put("mail.smtp.host", host);
  props.put("mail.smtp.port", "587");
  props.put("mail.smtp.auth", "true");
  // We need TTLS, which gmail requires
  props.put("mail.smtp.starttls.enable", "true");

  // Create a session
  Session session = Session.getDefaultInstance(props, new Auth());

  try {
    MimeMessage msg = new MimeMessage(session);
    msg.setFrom(new InternetAddress("attendancePad@gmail.com", lbl29));
    msg.addRecipient(Message.RecipientType.TO, new InternetAddress(sendToMail));
    String date = year() + "/" + month() + "/" + day();
    msg.setSubject(lbl30 + " (" + date + ")");
    BodyPart messageBodyPart = new MimeBodyPart();
    // Fill the message
    messageBodyPart.setText(lbl31);
    Multipart multipart = new MimeMultipart();
    multipart.addBodyPart(messageBodyPart);
    // Part two is attachment
    messageBodyPart = new MimeBodyPart();
    DataSource source = new FileDataSource(fileNameExcel);
    messageBodyPart.setDataHandler(new DataHandler(source));
    messageBodyPart.setFileName(fileNameExcel);
    multipart.addBodyPart(messageBodyPart);
    msg.setContent(multipart);
    msg.setSentDate(new Date());
    Transport.send(msg);
    popOut(lbl12 + ": " + sendToMail, lblMessage);
  } 
  catch (Exception e) {
    popOut(lbl33, lblErrorTitle);
  }
}

void draw() {
  if (frameCount % 4 == 0) {
    SimpleDateFormat format4 = new SimpleDateFormat("HH:mm:ss");
    currentTimeFrameUp = format4.format(Calendar.getInstance().getTime());
    lblFrameDownTime.setText("Heure: "+currentTimeFrameUp+"   \t\tDate: "+year()+ "/"+month() + "/"+day()+"\t" );
  }

  if (frameCount % 14 == 0) {
    if (timerNewCard.isRunning()) {
      timerCountDown--;
      //println("Timer: "+timerCountDown);
      optionPaneTimer.setMessage("Glisser la nouvelle carte maintenant ("+timerCountDown+")");
    }
  }

  // Every 30 frames request new data
  if (frameCount % 30 == 0) {
    setConnectedPort();
  }

  // Every 30 frames request new data
  if (frameCount % 30 == 0) {
    //thread("requestData");
    setConnectedPort();

    SimpleDateFormat format = new SimpleDateFormat("HH:mm");
    currentTime = format.format(new Date());

    if (autoSendProgrammed && txtMailTime.equals(currentTime)) {
      generateExcelRecord();
      //openFile(homePath+"/"+fileNameExcel);
      sendToMail = txtEmail.getText();
      sendAttachementMail();
      autoSendProgrammed = false;
    }

    //loadPersonsListFromDB();

    String[] currentTimeParts = currentTimeFrameUp.split(":");

    if (autoSyncProgrammed) {
      if ((currentTimeParts[1].equals("00") && (currentTimeParts[2].equals("00"))) || (currentTimeParts[1].equals("0")) && (currentTimeParts[2].equals("00"))) {
        popOut("Les fichiers de presences seront synchroniser \navec le siteweb qadrine au lien: http://qadrine.eb2a.com/cicheck/", "Message");
        uploadFileFTP(recordFilePath);
      }
    }
    //println("currentCardCode: "+currentCardCode);
    saveNewUserNow();
  }
}

public class Auth extends Authenticator {

  public Auth() {
    super();
  }

  public PasswordAuthentication getPasswordAuthentication() {
    String username, password;
    username = ""; //your robot email
    password = ""; //your robot email password
    println(lbl34);
    return new PasswordAuthentication(username, password);
  }
}

public class ForcedListSelectionModel extends DefaultListSelectionModel {

  public ForcedListSelectionModel () {
    setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
  }

  @Override
    public void clearSelection() {
  }

  @Override
    public void removeSelectionInterval(int index0, int index1) {
  }
}

class Person {
  String fullName;
  String sex;
  String cardCode;
  Person(String f, String c, String s) {
    fullName = f;
    cardCode = c;
    sex = s;
  }
  String getFullName() {
    return fullName;
  }
  String getSex() {
    return sex;
  }
  String getCardCode() {
    return cardCode;
  }
}
