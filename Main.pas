unit Main;

interface

uses
  SysUtils, Windows, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, StdCtrls, Buttons, ExtCtrls, Menus, ComCtrls, DB, ADODB,
  Grids, DBGrids, DBCtrls, XPStyleActnCtrls, ActnList, ActnMan, ToolWin,
  ActnCtrls, Access2000, OleServer, DateUtils,
  RichEdit, Printers, ImgList, Math, System.ImageList;

type
  TMainForm = class(TForm)
    MainMenu: TMainMenu;
    FileMenu: TMenuItem;
    FileNewItem: TMenuItem;
    FileOpenItem: TMenuItem;
    FileSaveItem: TMenuItem;
    FileSaveAsItem: TMenuItem;
    N1: TMenuItem;
    FilePrintItem: TMenuItem;
    FilePrintSetupItem: TMenuItem;
    N4: TMenuItem;
    FileExitItem: TMenuItem;
    EditMenu: TMenuItem;
    EditUndoItem: TMenuItem;
    N2: TMenuItem;
    EditCutItem: TMenuItem;
    EditCopyItem: TMenuItem;
    EditPasteItem: TMenuItem;
    HelpMenu: TMenuItem;
    HelpContentsItem: TMenuItem;
    HelpSearchItem: TMenuItem;
    HelpHowToUseItem: TMenuItem;
    N3: TMenuItem;
    HelpAboutItem: TMenuItem;
    OpenDialog: TOpenDialog;
    SaveDialog: TSaveDialog;
    PrintDialog: TPrintDialog;
    PrintSetupDialog: TPrinterSetupDialog;
    ADOConnection1: TADOConnection;
    DsViolations: TDataSource;
    AdoTableGenVioLetters: TADOTable;
    dsGenVioLetters: TDataSource;
    adoTblHouses: TADOTable;
    dsHouses: TDataSource;
    adoTblAllApprovalLetters: TADOTable;
    DataSource2: TDataSource;
    dsAllApprovalLetters: TDataSource;
    dsOffsiteOwners: TDataSource;
    adoTblOffsiteOwners: TADOTable;
    adoTblWelcomeLetters: TADOTable;
    dsWelcomeLetters: TDataSource;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    pnlCurrentOwners: TPanel;
    pnlAllOwners: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    sbAcctSort: TSpeedButton;
    sbOwnerSort: TSpeedButton;
    sbStrNumSort: TSpeedButton;
    sbStrNameSort: TSpeedButton;
    eCurrentOwnerSearch: TEdit;
    eCurrentAcctSearch: TEdit;
    eCurrentStreetSearch: TEdit;
    eCurrentAddrSearch: TEdit;
    bdGridCurrentOwners: TDBGrid;
    dsCurrentOwners: TDataSource;
    AdoTableCurrentOwners: TADOTable;
    sbCurrentOwners: TStatusBar;
    pnlAcAppEnter: TPanel;
    dbGridApprovalLetters: TDBGrid;
    dbMemoAcAppApprovalWords: TDBMemo;
    pnlViolations: TPanel;
    pnlVioStatus: TPanel;
    dbGridViolations: TDBGrid;
    dbGridVioStatus: TDBGrid;
    dbMemoVioReason: TDBMemo;
    dbMemoVioStatAction: TDBMemo;
    AdoTableViolations: TADOTable;
    dbMemoVioStatStatus: TDBMemo;
    pnlGenVioLetterEnter: TPanel;
    lblVioStatStatus: TLabel;
    lblVioStatAction: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    lblVioStatus: TLabel;
    dbGridGenVioLetters: TDBGrid;
    dbMemoGenVioLettersSpecText: TDBMemo;
    dbMemoGenVioLettersVioText: TDBMemo;
    dbMemoGenVioLettersRemedyText: TDBMemo;
    sbApproveApplication: TSpeedButton;
    lblGenVioLettersSpecText: TLabel;
    lblGenVioLettersRemedyText: TLabel;
    lblGenVioLettersVioText: TLabel;
    Label13: TLabel;
    sbNewVioLetter: TSpeedButton;
    sbNewViolation: TSpeedButton;
    sbCopyVioLetter: TSpeedButton;
    pnlHousesEnter: TPanel;
    pnlOwnersEnter: TPanel;
    dbGridHouses: TDBGrid;
    DBGrid6: TDBGrid;
    pnlWelcomeEnter: TPanel;
    DBGrid7: TDBGrid;
    pnlAllGenVioLetters: TPanel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    dbGridBrowseGenVioLetters: TDBGrid;
    DBMemo8: TDBMemo;
    DBMemo9: TDBMemo;
    DBMemo10: TDBMemo;
    adoTableAllGenVioLetters: TADOTable;
    dsAllGenVioLetters: TDataSource;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    AdoTableGenVioLettersID: TAutoIncField;
    AdoTableGenVioLettersviolationId: TIntegerField;
    AdoTableGenVioLettershouseAcct: TIntegerField;
    AdoTableGenVioLettersletterNumber: TWideStringField;
    AdoTableGenVioLettersletterType: TWideStringField;
    AdoTableGenVioLettersspecificViolationTitle: TMemoField;
    AdoTableGenVioLetterstextViolationWords: TMemoField;
    AdoTableGenVioLettersletterDate: TDateTimeField;
    AdoTableGenVioLetterssenderInitials: TWideStringField;
    AdoTableGenVioLettersremedyWords: TMemoField;
    AdoTableGenVioLettersoriginalLetterDate: TDateTimeField;
    AdoTableGenVioLetterssignatureLine: TWideStringField;
    DataSource1: TDataSource;
    ADOTable1: TADOTable;
    DBGrid3: TDBGrid;
    ADOQuery1: TADOQuery;
    DataSource3: TDataSource;
    sbAppAcctSort: TSpeedButton;
    sbAppPermitSort: TSpeedButton;
    sbAppIssueDateSort: TSpeedButton;
    sbAppLetterSort: TSpeedButton;
    sbAppApproveDateSort: TSpeedButton;
    sbAppApplyDateSort: TSpeedButton;
    sbAppMonthSort: TSpeedButton;
    sbAppYearSort: TSpeedButton;
    DBNavigator1: TDBNavigator;
    adoTblPropInLegal: TADOTable;
    adoTblLegalStatus: TADOTable;
    DBGrid4: TDBGrid;
    DBGrid5: TDBGrid;
    dsPropInLegal: TDataSource;
    dsLegalStatus: TDataSource;
    DBMemo12: TDBMemo;
    DBMemo13: TDBMemo;
    PopupMenu1: TPopupMenu;
    DeleteRecord1: TMenuItem;
    DBNavigator3: TDBNavigator;
    DBNavigator2: TDBNavigator;
    DBNavigator4: TDBNavigator;
    DBNavigator5: TDBNavigator;
    lblAcApplications: TLabel;
    sbLegalSort: TSpeedButton;
    sbOffsiteSort: TSpeedButton;
    dsViolationStatus: TDataSource;
    adoTableViolationStatus: TADOTable;
    sbNextVioLetter: TSpeedButton;
    Close1: TMenuItem;
    TabSheet7: TTabSheet;
    DBGrid8: TDBGrid;
    adoTblAllLetters: TADOTable;
    DataSource4: TDataSource;
    adoQryClearLetters: TADOQuery;
    btnClearLetters: TButton;
    openDlgSql: TOpenDialog;
    btnOpenSql: TButton;
    btnSaveLetter: TButton;
    adoTblAllLettershouseAcct: TIntegerField;
    adoTblAllLetterslegal: TWideStringField;
    adoTblAllLettersletterType: TWideStringField;
    adoTblAllLettersacctSectionLine: TWideStringField;
    adoTblAllLettersowner: TWideStringField;
    adoTblAllLettersoffsiteAddress: TWideStringField;
    adoTblAllLettersoffsiteCityStateZip: TWideStringField;
    adoTblAllLetterscloseDate: TDateTimeField;
    adoTblAllLettersoffsite: TWideStringField;
    adoTblAllLettersmailName: TWideStringField;
    adoTblAllLettersaddress: TWideStringField;
    adoTblAllLettersviolationId: TWideStringField;
    adoTblAllLetterspropertySite: TWideStringField;
    adoTblAllLetterscityZip: TWideStringField;
    adoTblAllLettersletterDate: TDateTimeField;
    adoTblAllLettersdear: TWideStringField;
    adoTblAllLettersinRe: TWideStringField;
    adoTblAllLetterssenderInitials: TWideStringField;
    adoTblAllLetterstextViolationWords: TMemoField;
    adoTblAllLettersletterNumber: TWideStringField;
    adoTblAllLettersspecificViolationTitle: TMemoField;
    adoTblAllLettersStreetNumber: TIntegerField;
    adoTblAllLettersStreetName: TWideStringField;
    adoTblAllLettersclosing: TMemoField;
    adoTblAllLettersfirstParagraph: TMemoField;
    adoTblAllLetterspresidentLine: TWideStringField;
    adoTblAllLettersccField: TWideStringField;
    adoTblAllLettersspecificApprovalWords: TMemoField;
    adoTblAllLetterspermitNumber: TIntegerField;
    adoTblAllLetterspermitDateIssued: TDateTimeField;
    adoTblAllLetterspermitMonth: TIntegerField;
    adoTblAllLetterspermitYear: TIntegerField;
    adoTblAllLettersapplicationDate: TDateTimeField;
    adoTblAllLettersapprovalDate: TDateTimeField;
    adoTblAllLetterspermitLine: TWideStringField;
    adoTblAllLettersgreeting: TWideStringField;
    adoTblAllLettersoffsiteCity: TWideStringField;
    adoTblAllLettersfirstLine: TWideStringField;
    adoTblAllLettersownerId: TIntegerField;
    adoTblAllLetterswhoFrom: TWideStringField;
    adoTblAllLetterssellDate: TDateTimeField;
    adoTblAllLettersletterBody: TMemoField;
    FontDialog1: TFontDialog;
    btnFont: TButton;
    DBMemo14: TDBMemo;
    RichEdit1: TRichEdit;
    dbBrowseGenVioLetters: TDataSource;
    adoTblBrowseGenVioLetters: TADOTable;
    sbVioNum: TSpeedButton;
    sbAcctNum: TSpeedButton;
    sbLetterNum: TSpeedButton;
    sbDate: TSpeedButton;
    sbOrigLetterDate: TSpeedButton;
    sbLetterType: TSpeedButton;
    TabSheet8: TTabSheet;
    adoTblMemoToLegal: TADOTable;
    DBGrid1: TDBGrid;
    dsMemoToLegal: TDataSource;
    DBNavigator6: TDBNavigator;
    ImageList1: TImageList;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton9: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton12: TSpeedButton;
    SpeedButton13: TSpeedButton;
    dbMemoAcAppRejectWords: TDBMemo;
    dbMemoAcAppRemedyWords: TDBMemo;
    lblAcAppSpecAppWords: TLabel;
    lblAcAppSpecRejectWords: TLabel;
    lblAcAppSpecRemedyWords: TLabel;
    sbWelcomeId: TSpeedButton;
    sbWelcomeOwner: TSpeedButton;
    sbWelcomeClose: TSpeedButton;
    Label25: TLabel;
    DBMemo17: TDBMemo;
    Label26: TLabel;
    DBMemo18: TDBMemo;
    Label27: TLabel;
    DBMemo19: TDBMemo;
    btnVioStatus: TButton;
    adoVioQry: TADOQuery;
    dsVioQry: TDataSource;
    SpeedButton16: TSpeedButton;
    adoVioQryviolationID: TAutoIncField;
    adoVioQryacctID: TIntegerField;
    adoVioQryhouseAcct: TIntegerField;
    adoVioQryviolationDate: TDateTimeField;
    adoVioQryreason: TMemoField;
    adoVioQryaction: TWideStringField;
    adoVioQryreportedBy: TWideStringField;
    adoVioQryopenDate: TDateTimeField;
    adoVioQrycloseDate: TDateTimeField;
    adoVioQryStartDate: TDateTimeField;
    adoVioQryStopDate: TDateTimeField;
    adoVioQrylegalDate: TDateTimeField;
    AdoDataSetViolations: TADODataSet;
    adsViolations: TDataSource;
    AdoDataSetVioStatus: TADODataSet;
    DataSource6: TDataSource;
    AdoDataSetVioLetters: TADODataSet;
    DataSource5: TDataSource;
    adoQryRunLetters: TADOQuery;
    btnRunLetters: TButton;
    SpeedButton14: TSpeedButton;
    menuLetters: TMenuItem;
    menuLetters7Days: TMenuItem;
    menuLetters2Weeks: TMenuItem;
    menuLetters30Days: TMenuItem;
    menuLetters60Days: TMenuItem;
    menuLetters90Days: TMenuItem;
    N5: TMenuItem;
    menuLettersAll: TMenuItem;
    adoQryCullLetters: TADOQuery;
    statBarGenLetters: TStatusBar;
    Label28: TLabel;
    eCurrentRouteSearch: TEdit;
    sbRouteSort: TSpeedButton;
    menuLettersToday: TMenuItem;
    N6: TMenuItem;
    menuLetters2Days: TMenuItem;
    menuLetters3Days: TMenuItem;
    TabSheet9: TTabSheet;
    sbViolations: TStatusBar;
    StatusBar2: TStatusBar;
    Splitter1: TSplitter;
    Connect1: TMenuItem;
    N7: TMenuItem;
    splitVioLetters: TSplitter;
    splitVioStatus: TSplitter;
    Splitter2: TSplitter;
    Splitter3: TSplitter;
    Splitter4: TSplitter;
    pnlGenLetters1: TPanel;
    pnlGenLetters2: TPanel;
    sbOwners: TStatusBar;
    Splitter5: TSplitter;
    pnlAcApps1: TPanel;
    pnlAcApps2: TPanel;
    pnlMemoToLegal: TPanel;
    AdoTableCurrentOwnersownerID: TAutoIncField;
    AdoTableCurrentOwnershouseAcct: TIntegerField;
    AdoTableCurrentOwnersOwner: TWideStringField;
    AdoTableCurrentOwnerslegal: TWideStringField;
    AdoTableCurrentOwnersPhone: TWideStringField;
    AdoTableCurrentOwnersAltPhone: TWideStringField;
    AdoTableCurrentOwnersOffsite: TWideStringField;
    AdoTableCurrentOwnersAddress: TWideStringField;
    AdoTableCurrentOwnersZip: TIntegerField;
    AdoTableCurrentOwnersLot: TWideStringField;
    AdoTableCurrentOwnersBlock: TWideStringField;
    AdoTableCurrentOwnersSection: TWideStringField;
    AdoTableCurrentOwnersStreetNumber: TIntegerField;
    AdoTableCurrentOwnersStreetName: TWideStringField;
    AdoTableCurrentOwnersgreeting: TWideStringField;
    AdoTableCurrentOwnersmailName: TWideStringField;
    AdoTableCurrentOwnerscloseDate: TDateTimeField;
    AdoDataSetOwners: TADODataSet;
    dsOwners: TDataSource;
    AdoDataAllAppLetters: TADODataSet;
    dbNavOwners: TDBNavigator;
    tsAllOwners: TTabSheet;
    DBGrid2: TDBGrid;
    StatusBar1: TStatusBar;
    adoTableOwners: TADOTable;
    dsAllOwners: TDataSource;
    ColorDialog1: TColorDialog;
    adoTableOwnersownerID: TAutoIncField;
    adoTableOwnershouseAcct: TIntegerField;
    adoTableOwnersOwner: TWideStringField;
    adoTableOwnersOffsite: TWideStringField;
    adoTableOwnersPhone: TWideStringField;
    adoTableOwnersAltPhone: TWideStringField;
    adoTableOwnerscloseDate: TDateTimeField;
    adoTableOwnerssellDate: TDateTimeField;
    adoTableOwnersgreeting: TWideStringField;
    adoTableOwnersmailName: TWideStringField;
    adoTableOwnersAddress: TWideStringField;
    adoTableOwnerscity: TWideStringField;
    adoTableOwnersstate: TWideStringField;
    adoTableOwnersZip: TIntegerField;
    adoTableOwnerscityStateZip: TWideStringField;
    AdoDataSetViolationsviolationID: TAutoIncField;
    AdoDataSetViolationsacctID: TIntegerField;
    AdoDataSetViolationshouseAcct: TIntegerField;
    AdoDataSetViolationsviolationDate: TDateTimeField;
    AdoDataSetViolationsreason: TMemoField;
    AdoDataSetViolationsaction: TWideStringField;
    AdoDataSetViolationsreportedBy: TWideStringField;
    AdoDataSetViolationsopenDate: TDateTimeField;
    AdoDataSetViolationscloseDate: TDateTimeField;
    AdoDataSetViolationsStartDate: TDateTimeField;
    AdoDataSetViolationsStopDate: TDateTimeField;
    AdoDataSetViolationslegalDate: TDateTimeField;
    AdoDataSetVioLettersID: TAutoIncField;
    AdoDataSetVioLettersviolationId: TIntegerField;
    AdoDataSetVioLettershouseAcct: TIntegerField;
    AdoDataSetVioLettersletterNumber: TWideStringField;
    AdoDataSetVioLettersletterType: TWideStringField;
    AdoDataSetVioLettersspecificViolationTitle: TMemoField;
    AdoDataSetVioLetterstextViolationWords: TMemoField;
    AdoDataSetVioLettersletterDate: TDateTimeField;
    AdoDataSetVioLetterssenderInitials: TWideStringField;
    AdoDataSetVioLettersremedyWords: TMemoField;
    AdoDataSetVioLettersoriginalLetterDate: TDateTimeField;
    AdoDataSetVioLetterssignatureLine: TWideStringField;
    AdoDataAllAppLettersID: TAutoIncField;
    AdoDataAllAppLettershouseAcct: TIntegerField;
    AdoDataAllAppLetterspermitNumber: TIntegerField;
    AdoDataAllAppLettersspecificApprovalWords: TMemoField;
    AdoDataAllAppLettersletterDate: TDateTimeField;
    AdoDataAllAppLetterspermitDateIssued: TDateTimeField;
    AdoDataAllAppLetterspermitMonth: TIntegerField;
    AdoDataAllAppLetterspermitYear: TIntegerField;
    AdoDataAllAppLettersapplicationDate: TDateTimeField;
    AdoDataAllAppLettersapprovalDate: TDateTimeField;
    AdoDataAllAppLettersspecificRejectionWords: TMemoField;
    AdoDataAllAppLettersRemedyWords: TMemoField;
    AdoDataSetOwnersownerID: TAutoIncField;
    AdoDataSetOwnershouseAcct: TIntegerField;
    AdoDataSetOwnersOwner: TWideStringField;
    AdoDataSetOwnersOffsite: TWideStringField;
    AdoDataSetOwnersPhone: TWideStringField;
    AdoDataSetOwnersAltPhone: TWideStringField;
    AdoDataSetOwnerscloseDate: TDateTimeField;
    AdoDataSetOwnerssellDate: TDateTimeField;
    AdoDataSetOwnersgreeting: TWideStringField;
    AdoDataSetOwnersmailName: TWideStringField;
    AdoDataSetOwnersAddress: TWideStringField;
    AdoDataSetOwnerscity: TWideStringField;
    AdoDataSetOwnersstate: TWideStringField;
    AdoDataSetOwnersZip: TIntegerField;
    AdoDataSetOwnerscityStateZip: TWideStringField;
    adoTblWelcomeLettersID: TAutoIncField;
    adoTblWelcomeLettersownerId: TIntegerField;
    adoTblWelcomeLettershouseAcct: TIntegerField;
    adoTblWelcomeLettersletterDate: TDateTimeField;
    adoTblWelcomeLetterswhoFrom: TWideStringField;
    adoTblWelcomeLettersInvestor: TWideStringField;
    adoTblHouseshouseAcct: TIntegerField;
    adoTblBrowseGenVioLettershouseAcct: TIntegerField;
    adoTblBrowseGenVioLettersID: TAutoIncField;
    adoTblBrowseGenVioLettersviolationId: TIntegerField;
    adoTblBrowseGenVioLettersletterNumber: TWideStringField;
    adoTblBrowseGenVioLettersletterType: TWideStringField;
    adoTblBrowseGenVioLettersspecificViolationTitle: TMemoField;
    adoTblBrowseGenVioLetterstextViolationWords: TMemoField;
    adoTblBrowseGenVioLettersletterDate: TDateTimeField;
    adoTblBrowseGenVioLetterssenderInitials: TWideStringField;
    adoTblBrowseGenVioLettersremedyWords: TMemoField;
    adoTblBrowseGenVioLettersoriginalLetterDate: TDateTimeField;
    adoTblBrowseGenVioLetterssignatureLine: TWideStringField;
    adoTblHousesacctID: TAutoIncField;
    adoTblHousesOwner: TWideStringField;
    adoTblHousesAddress: TWideStringField;
    adoTblHousesOffsite: TWideStringField;
    adoTblHouseslegal: TWideStringField;
    adoTblHousesZip: TIntegerField;
    adoTblHousesLot: TWideStringField;
    adoTblHousesBlock: TWideStringField;
    adoTblHousesSection: TWideStringField;
    adoTblHousesStreetNumber: TIntegerField;
    adoTblHousesStreetName: TWideStringField;
    menuCompactDb: TMenuItem;
    N8: TMenuItem;
    Label5: TLabel;
    eCurrentPhoneSearch: TEdit;
    AdoTableCurrentOwnersdriveRoute: TFloatField;
    DataSource7: TDataSource;
    ADOTable2: TADOTable;
    Button1: TButton;
    adoTblHousesdriveRoute: TFloatField;
    DBNavigator7: TDBNavigator;
    AdoTableCurrentOwnersmobilePhone1: TWideStringField;
    AdoTableCurrentOwnersmobilePhone2: TWideStringField;
    AdoDataSetOwnersmobilePhone1: TWideStringField;
    AdoDataSetOwnersmobilePhone2: TWideStringField;
    dbNavWelcome: TDBNavigator;
    procedure FormCreate(Sender: TObject);
    procedure ShowHint(Sender: TObject);
    procedure FileNew(Sender: TObject);
    procedure FileOpen1(Sender: TObject);
    procedure FileSave(Sender: TObject);
    procedure FileSaveAs(Sender: TObject);
    procedure FilePrint(Sender: TObject);
    procedure FilePrintSetup(Sender: TObject);
    procedure FileExit(Sender: TObject);
    procedure EditUndo(Sender: TObject);
    procedure EditCut(Sender: TObject);
    procedure EditCopy(Sender: TObject);
    procedure EditPaste(Sender: TObject);
    procedure HelpContents(Sender: TObject);
    procedure HelpSearch(Sender: TObject);
    procedure HelpHowToUse(Sender: TObject);
    procedure HelpAbout(Sender: TObject);
    procedure EditColorChangeOnEnter(Sender: TObject);
    procedure EditColorChangeOnExit(Sender: TObject);
    procedure sbNewViolationClick(Sender: TObject);
    procedure sbNewGenVioLetterClick(Sender: TObject);
    procedure frmCurrentOwners21DbGridCurrentOwnersDblClick(Sender: TObject);
    procedure sbwww(Sender: TObject);
    procedure eCurrentAcctSearchChange(Sender: TObject);
    procedure eCurrentStreetSearchChange(Sender: TObject);
    procedure eCurrentAddrSearchChange(Sender: TObject);
    procedure sbCopyGenVioLetterClick(Sender: TObject);
    procedure sbAcctSortClick(Sender: TObject);
    procedure sbOwnerSortClick(Sender: TObject);
    procedure sbStrNumSortClick(Sender: TObject);
    procedure sbStrNameSortClick(Sender: TObject);
    procedure eOwnersOwnerOnChange(Sender: TObject);
    procedure sbApproveApplicationClick(Sender: TObject);
    procedure sbNewVioLetterClick(Sender: TObject);
    procedure sbCopyVioLetterClick(Sender: TObject);
    procedure sbAppAcctSortClick(Sender: TObject);

    procedure sortApplicationTable(sortString: string);
    procedure sbAppLetterSortClick(Sender: TObject);
    procedure sbAppIssueDateSortClick(Sender: TObject);
    procedure sbAppPermitSortClick(Sender: TObject);
    procedure sbAppMonthSortClick(Sender: TObject);
    procedure sbAppYearSortClick(Sender: TObject);
    procedure sbAppApplyDateSortClick(Sender: TObject);
    procedure sbAppApproveDateSortClick(Sender: TObject);
    procedure sbLegalSortClick(Sender: TObject);
    procedure sbOffsiteSortClick(Sender: TObject);
    procedure sbNextVioLetterClick(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure btnClearLettersClick(Sender: TObject);
    procedure btnOpenSqlClick(Sender: TObject);
    procedure btnSaveLetterClick(Sender: TObject);
    procedure btnFontClick(Sender: TObject);
    procedure DBGrid8CellClick(Column: TColumn);
    procedure sbVioNumClick(Sender: TObject);
    procedure sbAcctNumClick(Sender: TObject);
    procedure sbLetterNumClick(Sender: TObject);
    procedure sbDateClick(Sender: TObject);
    procedure sbOrigLetterDateClick(Sender: TObject);
    procedure sbLetterTypeClick(Sender: TObject);
    procedure dbGridBrowseGenVioLettersCellClick(Column: TColumn);
    procedure DBGrid3CellClick(Column: TColumn);
    procedure adoTblMemoToLegalNewRecord(DataSet: TDataSet);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure BuildSortAppTableQuery(IndexFieldName: string;
      PushButton: TSpeedButton);
    procedure GenerateLetterButtonClear;
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure SpeedButton12Click(Sender: TObject);
    procedure SpeedButton13Click(Sender: TObject);
    procedure sbWelcomeCloseClick(Sender: TObject);
    procedure sbWelcomeIdClick(Sender: TObject);
    procedure sbWelcomeOwnerClick(Sender: TObject);
    procedure btnVioStatusClick(Sender: TObject);
    procedure SpeedButton16Click(Sender: TObject);
    procedure AdoTableCurrentOwnersAfterScroll(DataSet: TDataSet);
    procedure AdoDataSetViolationsAfterScroll(DataSet: TDataSet);
    procedure AdoDataSetVioStatusAfterInsert(DataSet: TDataSet);
    procedure AdoDataSetViolationsAfterInsert(DataSet: TDataSet);
    procedure AdoDataSetVioLettersAfterInsert(DataSet: TDataSet);
    procedure AdoDataAllAppLettersAfterInsert(DataSet: TDataSet);
    procedure AdoDataSetViolationsAfterPost(DataSet: TDataSet);
    procedure AdoDataSetVioStatusAfterPost(DataSet: TDataSet);
    procedure btnRunLettersClick(Sender: TObject);
    procedure SpeedButton14Click(Sender: TObject);
    procedure menuLettersClick(Sender: TObject);
    procedure menuLettersUncheck;
    procedure menuLettersSetDefault;
    procedure menuLettersSetCheck(Sender: TObject);
    procedure statBarUpdate;
    procedure eCurrentRouteSearchChange(Sender: TObject);
    procedure sbRouteSortClick(Sender: TObject);
    procedure AppTableButtonClear;
    procedure FormResize(Sender: TObject);
    procedure pnlCurrentOwnersResize(Sender: TObject);
    procedure Connect1Click(Sender: TObject);
    procedure pnlAcAppEnterResize(Sender: TObject);
    procedure pnlGenVioLetterEnterResize(Sender: TObject);
    procedure pnlViolationsResize(Sender: TObject);
    procedure pnlAllGenVioLettersResize(Sender: TObject);
    procedure pnlVioStatusResize(Sender: TObject);
    procedure pnlHousesEnterResize(Sender: TObject);
    procedure pnlOwnersEnterResize(Sender: TObject);
    procedure pnlWelcomeEnterResize(Sender: TObject);
    procedure pnlGenLetters2Resize(Sender: TObject);
    procedure pnlGenLetters1Resize(Sender: TObject);
    procedure TabSheet7Resize(Sender: TObject);
    procedure pnlAcApps2Resize(Sender: TObject);
    procedure pnlMemoToLegalResize(Sender: TObject);
    procedure AdoDataSetOwnersAfterInsert(DataSet: TDataSet);
    procedure AdoDataSetOwnersAfterPost(DataSet: TDataSet);
    procedure PageControl1Change(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure tsAllOwnersResize(Sender: TObject);
    procedure tsAllOwnersEnter(Sender: TObject);
    procedure tsAllOwnersHide(Sender: TObject);
    procedure menuCompactDbClick(Sender: TObject);
    procedure eCurrentPhoneSearchChange(Sender: TObject);
    procedure DBGrid2CellClick(Column: TColumn);
    procedure Button1Click(Sender: TObject);
    procedure PageControl1DrawTab(Control: TCustomTabControl; TabIndex: Integer;
      const Rect: TRect; Active: Boolean);
    procedure adoTblWelcomeLettersAfterInsert(DataSet: TDataSet);
    procedure ADOQuery1AfterScroll(DataSet: TDataSet);
    procedure adoTblBrowseGenVioLettersAfterScroll(DataSet: TDataSet);
    procedure adoTableOwnersAfterScroll(DataSet: TDataSet);
    procedure adoTblAllLettersAfterScroll(DataSet: TDataSet);
    procedure adoTblMemoToLegalAfterScroll(DataSet: TDataSet);
  private
    procedure SortColumn(DataTable: TADOTable; IndexFieldName: string;
      PushButton: TSpeedButton); overload;
    procedure SortColumn(DataSet: TADODataSet; IndexFieldName: string;
      PushButton: TSpeedButton); overload;
    procedure RemoveGlyphImages(Sender: TObject);
    procedure TEditChange(filterText: string);
    procedure CompactDatabase(const MdbFileName: string;
      ADOConnection: TADOConnection = nil;
      const ReopenCOnnection: Boolean = True);
    procedure JroRefreshCache(ADOConnection: TADOConnection);
    procedure JroCompactDatabase(const Source, Destination: string);
    procedure UpdateCurrentHouseAcct(newHouseAcct: string);


  end;

  procedure DrawTab( oControl: TCustomTabControl; TabIndex: Integer;
      const Rect: TRect; Active: Boolean);

var
  MainForm: TMainForm;
  LetterDate: Tdate;

implementation

uses
  about, Variants, StrUtils, Types, ComObj, System.UITypes;

{$R *.dfm}

procedure TMainForm.FormCreate(Sender: TObject);
begin
  Application.OnHint := ShowHint;
  ADOConnection1.Connected := True;
  sbCurrentOwners.Panels[2].Text := ADOConnection1.ConnectionString;
  menuLettersSetDefault;
  statBarUpdate;
end;

procedure TMainForm.ShowHint(Sender: TObject);
begin
  // StatusLine.SimpleText := Application.Hint;
end;

procedure TMainForm.FileNew(Sender: TObject);
begin
  { Add code to create a new file }
end;

procedure TMainForm.FileOpen1(Sender: TObject);
const
  { Connection string }
  ConnString = 'Provider=Microsoft.ACE.OLEDB.12.0;' + 'User ID=Admin;' +
    'Data Source=%s;' + 'Mode=Share Deny None;' +
    'Jet OLEDB:System database="";' + 'Jet OLEDB:Registry Path="";' +
    'Jet OLEDB:Database Password="";' + 'Jet OLEDB:Engine Type=6;' +
    'Jet OLEDB:Database Locking Mode=1;' +
    'Jet OLEDB:Global Partial Bulk Ops=2;' +
    'Jet OLEDB:Global Bulk Transactions=1;' +
    'Jet OLEDB:New Database Password="";' +
    'Jet OLEDB:Create System Database=False;' +
    'Jet OLEDB:Encrypt Database=False;' +
    'Jet OLEDB:Don''t Copy Locale on Compact=False;' +
    'Jet OLEDB:Compact Without Replica Repair=False;' + 'Jet OLEDB:SFP=False;' +
    'Jet OLEDB:Support Complex Data=False;' +
    'Jet OLEDB:Bypass UserInfo Validation=False;';
var
  WhichConnection: string;
begin
  { execute an open file dialog }
  if OpenDialog.Execute then
    { first check if file exists }
    if FileExists(OpenDialog.FileName) then
    begin
      { if it exists, connect to the database }
      { First we start with a known condition }
      ADOConnection1.Connected := False;
      { Setup the connection string }
      ADOConnection1.ConnectionString :=
        Format(ConnString, [OpenDialog.FileName]);
      { Disable login prompt }
      ADOConnection1.LoginPrompt := False;
      { connect -- or die trying }
      try
        WhichConnection := 'ADOConnection1';
        ADOConnection1.Connected := True;
      except
        on e: EADOError do
        begin
          MessageDlg('Error while connecting', mtError, [mbOK], 0);
          ShowMessage('Error with ' + WhichConnection);
          Exit;
        end;
      end; // try-except
    end
    else
    begin
      { otherwise raise an exception }
      raise Exception.Create('File does not exist.');
      Exit
    end;
  { Success - now open the connection and connect the tables }
  // AdoTableViolations.Active       := True;
  AdoTableGenVioLetters.Active := True;
  AdoDataSetOwners.Active := True;
  adoTblHouses.Active := True;
  adoTableOwners.Active := True;
  // adoTblApprovalLetters.Active    := True;
  adoTblAllApprovalLetters.Active := True;
  adoTblOffsiteOwners.Active := True;
  adoTblWelcomeLetters.Active := True;
  // AdoTableViolationStatus.Active := True;
  AdoTableCurrentOwners.Active := True;
  adoTblBrowseGenVioLetters.Active := True;
  adoTblMemoToLegal.Active := True;
  adoTblPropInLegal.Active := True;
  adoTblLegalStatus.Active := True;
  // ShowMessage(ADOConnection1.ConnectionString);
end;

procedure TMainForm.FileSave(Sender: TObject);
begin
  { Add code to save current file under current name }
end;

procedure TMainForm.FileSaveAs(Sender: TObject);
begin
  if SaveDialog.Execute then
  begin
    { Add code to save current file under SaveDialog.FileName }
  end;
end;

procedure TMainForm.FilePrint(Sender: TObject);
begin
  if PrintDialog.Execute then
  begin
    { Add code to print current file }
  end;
end;

procedure TMainForm.FilePrintSetup(Sender: TObject);
begin
  PrintSetupDialog.Execute;
end;

procedure TMainForm.FileExit(Sender: TObject);
begin
  Close;
end;

procedure TMainForm.EditUndo(Sender: TObject);
begin
  { Add code to perform Edit Undo }
end;

procedure TMainForm.EditCut(Sender: TObject);
begin
  { Add code to perform Edit Cut }
end;

procedure TMainForm.EditCopy(Sender: TObject);
begin
  { Add code to perform Edit Copy }
end;

procedure TMainForm.EditPaste(Sender: TObject);
begin
  { Add code to perform Edit Paste }
end;

procedure TMainForm.HelpContents(Sender: TObject);
begin
  Application.HelpCommand(HELP_CONTENTS, 0);
end;

procedure TMainForm.HelpSearch(Sender: TObject);
const
  EmptyString: PChar = '';
begin
  Application.HelpCommand(HELP_PARTIALKEY, Longint(EmptyString));
end;

procedure TMainForm.HelpHowToUse(Sender: TObject);
begin
  Application.HelpCommand(HELP_HELPONHELP, 0);
end;

procedure TMainForm.HelpAbout(Sender: TObject);
begin
  AboutBox.ShowModal; { Add code to show program's About Box }
end;

procedure TMainForm.EditColorChangeOnEnter(Sender: TObject);
begin
  TEdit(Sender).Color := clMoneyGreen;
  eCurrentAcctSearch.Text := '';
  eCurrentStreetSearch.Text := '';
  eCurrentAddrSearch.Text := '';
  eCurrentRouteSearch.Text := '';
  eCurrentOwnerSearch.Text := '';
  eCurrentPhoneSearch.Text := '';
end;

procedure TMainForm.EditColorChangeOnExit(Sender: TObject);
begin
  TEdit(Sender).Color := clWindow;
end;

procedure TMainForm.sbNewViolationClick(Sender: TObject);
begin
  AdoTableViolations.Append;
end;

procedure TMainForm.sbNewGenVioLetterClick(Sender: TObject);
begin
  AdoTableGenVioLetters.Append;
end;

procedure TMainForm.frmCurrentOwners21DbGridCurrentOwnersDblClick
  (Sender: TObject);
//var
//  filterHouseAcct: integer;
begin
  { filterHouseAcct := AdoTableCurrentOwners.FieldValues['houseAcct'];
    AdoTableCurrentOwners.Filtered := false;
    AdoTableCurrentOwners.Filter := 'houseAcct = ' + IntToStr(filterHouseAcct) ;
    AdoTableCurrentOwners.Filtered := true; }
end;

procedure TMainForm.sbwww(Sender: TObject);
begin
  // adoTblApprovalLetters.IndexFieldNames := 'letterDate';
end;

{ ----------------------------------------------------------------------+
  TeditChange:
  This procedure is called when the user changes the search TEdit
  boxes. It changes the filters to the CurrentOwners dbGrid. It
  also updates the status bars on both the CurrentOwners and the
  Owners panels.
  +----------------------------------------------------------------------- }
procedure TMainForm.TEditChange(filterText: string);
var
  numOfRecords: integer;
begin
  if (length(filterText) = 0) then
  begin
    AdoTableCurrentOwners.Filter := '';
    AdoTableCurrentOwners.Filtered := True;
    adoTableOwners.Filter := '';
    adoTableOwners.Filtered := True;
  end
  else
  begin
    AdoTableCurrentOwners.Filtered := False;
    AdoTableCurrentOwners.Filter := filterText;
    AdoTableCurrentOwners.Filtered := True;
  end;
  sbCurrentOwners.Panels[0].Text := filterText;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);

  StatusBar1.Panels[0].Text := filterText;
  numOfRecords := adoTableOwners.RecordCount;
  StatusBar1.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

{ ----------------------------------------------------------------------+
  eCurrentAcctSearchChange:
  This procedure is called when the user changes the CurrentAcct
  search box.
  +----------------------------------------------------------------------- }
procedure TMainForm.eCurrentAcctSearchChange(Sender: TObject);
begin
  UpdateCurrentHouseAcct(eCurrentAcctSearch.Text);
end;

{ ----------------------------------------------------------------------+
  eCurrentStreetSearchChange:
  This procedure is called when the user changes the CurrentStreet
  search box.
  +----------------------------------------------------------------------- }
procedure TMainForm.eCurrentStreetSearchChange(Sender: TObject);
var
  filterText: string;
begin
  filterText := '';
  if (length(eCurrentStreetSearch.Text) > 0) then
    filterText := 'streetName like ' + '''' + eCurrentStreetSearch.Text +
      '%' + '''';
  TEditChange(filterText);
end;

{ ----------------------------------------------------------------------+
  eCurrentAddrSearchChange:
  This procedure is called when the user changes the CurrentAddr
  search box.
  +----------------------------------------------------------------------- }
procedure TMainForm.eCurrentAddrSearchChange(Sender: TObject);
var
  filterText: string;
begin
  filterText := '';
  if (length(eCurrentAddrSearch.Text) > 0) then
    filterText := 'address like ' + '''' + eCurrentAddrSearch.Text + '%' + '''';
  TEditChange(filterText);
end;

procedure TMainForm.sbCopyGenVioLetterClick(Sender: TObject);
var
  vinNumber: integer;
begin
  vinNumber := dsGenVioLetters.DataSet.FieldValues['violationId'];
  ShowMessage('VinNumber = ' + IntToStr(vinNumber));
end;

procedure TMainForm.sbAcctSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  adoTableOwners.Filtered := False;
  sbAcctSort.Glyph.Assign(nil);
  if (sbAcctSort.Tag = 1) then
  begin
    AdoTableCurrentOwners.IndexFieldNames := 'houseAcct ASC';
    ImageList1.GetBitmap(1, sbAcctSort.Glyph);
    sbAcctSort.Tag := 0;

    adoTableOwners.IndexFieldNames := 'houseAcct ASC';

  end
  else
  begin
    AdoTableCurrentOwners.IndexFieldNames := 'houseAcct DESC';
    ImageList1.GetBitmap(0, sbAcctSort.Glyph);
    sbAcctSort.Tag := 1;

    adoTableOwners.IndexFieldNames := 'houseAcct DESC';
  end;
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);

  adoTableOwners.Filtered := True;
  numOfRecords := adoTableOwners.RecordCount;
  StatusBar1.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.sbOwnerSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  AdoTableCurrentOwners.IndexFieldNames := 'Owner';
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);

  adoTableOwners.Filtered := False;
  adoTableOwners.IndexFieldNames := 'Owner';
  adoTableOwners.Filtered := True;
  numOfRecords := adoTableOwners.RecordCount;
  StatusBar1.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.sbStrNumSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  if (sbStrNumSort.Tag = 1) then
  begin
    AdoTableCurrentOwners.IndexFieldNames := 'streetNumber ASC';
    sbStrNumSort.Glyph.Assign(nil);
    ImageList1.GetBitmap(1, sbStrNumSort.Glyph);
    sbStrNumSort.Tag := 0;
  end
  else
  begin
    AdoTableCurrentOwners.IndexFieldNames := 'streetNumber DESC';
    sbStrNumSort.Glyph.Assign(nil);
    ImageList1.GetBitmap(0, sbStrNumSort.Glyph);
    sbStrNumSort.Tag := 1;
  end;
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.sbStrNameSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  AdoTableCurrentOwners.IndexFieldNames := 'streetName';
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.eOwnersOwnerOnChange(Sender: TObject);
var
  numOfRecords: integer;
begin
  // First we ignore the wildcard as the first character
  if (eCurrentOwnerSearch.Text = '%') then
    Exit;
  // Now we process the string
  if (length(eCurrentOwnerSearch.Text) = 0) then
  begin
    AdoTableCurrentOwners.Filter := '';
    AdoTableCurrentOwners.Filtered := True;

    adoTableOwners.Filter := '';
    adoTableOwners.Filtered := True;
  end
  else
  begin
    AdoTableCurrentOwners.Filtered := False;
    AdoTableCurrentOwners.Filter := 'Owner like ' + '''' +
      eCurrentOwnerSearch.Text + '%' + '''';
    AdoTableCurrentOwners.Filtered := True;

    adoTableOwners.Filtered := False;
    adoTableOwners.Filter := 'Owner like ' + '''' + eCurrentOwnerSearch.Text +
      '%' + '''';
    adoTableOwners.Filtered := True;
  end;
  sbCurrentOwners.Panels[0].Text := 'Current ' + AdoTableCurrentOwners.Filter;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);

  StatusBar1.Panels[0].Text := 'Any ' + AdoTableCurrentOwners.Filter;
  numOfRecords := adoTableOwners.RecordCount;
  StatusBar1.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

{ ----------------------------------------------------------------------+
  sbApproveApplicationClick:
  This is a procedure to append a new record to the ApprovalLetters
  table. Several fields are populated:

  approvalDate      : Today's date
  letterDate        : Today's date
  permitDateIssued  : Today's date
  permitMonth       : Today's month
  permitYear        : Today's year
  +----------------------------------------------------------------------- }
procedure TMainForm.sbApproveApplicationClick(Sender: TObject);
var
  myHouseAcct: integer;
begin
  myHouseAcct := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  with AdoDataAllAppLetters do
  begin
    if not(AdoDataAllAppLetters.State in [dsInsert, dsEdit]) then
    begin
      dsAllApprovalLetters.Edit;
    end;
    FieldValues['houseAcct'] := myHouseAcct;
    FieldValues['permitNumber'] := 12345;
    FieldValues['approvalDate'] := Date;
    FieldValues['letterDate'] := Date;
    FieldValues['permitDateIssued'] := Date;
    FieldValues['permitMonth'] := MonthOf(Date);
    FieldValues['permitYear'] := YearOf(Date);
  end;
end;

{ ----------------------------------------------------------------------+
  sbNewVioLetterClick:
  This is a procedure to append a new record to the GenVioLetters
  table. Several fields are populated:

  letterDate    : Today's date
  letterType    : GEN (for generic)
  letterNumber  : FOL (First Office Letter)
  senderInitials: BOZO
  Memo Fields   : names of the fields
  +----------------------------------------------------------------------- }
procedure TMainForm.sbNewVioLetterClick(Sender: TObject);
begin
  with AdoDataSetVioLetters do
  begin
    Append;
    FieldByName('letterDate').Value := Date; // Today's Date
    FieldValues['letterType'] := 'GEN';
    FieldByName('letterNumber').Value := 'FOL';
    FieldValues['specificViolationTitle'] := 'Specific Violation Title';
    FieldValues['textViolationWords'] := 'Text Violation Words';
    FieldValues['remedyWords'] := 'Remedy Words';
    FieldValues['senderInitials'] := 'Bozo';
  end;
end;

{ ----------------------------------------------------------------------+
  sbNextVioLetterClick:
  This is a procedure to append a new record the GenVioLetters
  table which is the next letter in the series to the letter in the
  current record.
  +----------------------------------------------------------------------- }
procedure TMainForm.sbNextVioLetterClick(Sender: TObject);
var
  oldLetterNumber, newLetterNumber, oldInitials: string;
  oldLetterType, oldLetterDate, oldOrigLetterDate, oldSignLine: OleVariant;
  oldSpecVioText, oldTextVio, oldRemedyWords: OleVariant;
begin
  // First, copy the date from the current record
  with AdoDataSetVioLetters do
  begin
    oldLetterNumber := FieldValues['letterNumber'];
    oldLetterType := FieldValues['letterType'];
    oldLetterDate := FieldValues['letterDate'];
    oldInitials := FieldValues['senderInitials'];
    oldOrigLetterDate := FieldValues['originalLetterDate'];
    oldSignLine := FieldValues['signatureLine'];
    oldSpecVioText := FieldValues['specificViolationTitle'];
    oldTextVio := FieldValues['textViolationWords'];
    oldRemedyWords := FieldValues['remedyWords'];
  end;

  // Determine the next letter to be sent
  if (oldLetterNumber = 'FOL') then
    newLetterNumber := '2OL'
  else if (oldLetterNumber = '2OL') then
    newLetterNumber := '3OL'
  else if (oldLetterNumber = '3OL') then
    newLetterNumber := '4OL'
  else if (oldLetterNumber = '4OL') then
    newLetterNumber := '209'
  else if (oldLetterNumber = '209') then
    newLetterNumber := '209'
  else
    newLetterNumber := '';

  // 3OL needs the signatureLine field to be populated, so set default value
  if (oldLetterNumber = '2OL') then
    oldSignLine := 'Bozo D. Clown';

  // Append a new record and populate the fields
  with AdoDataSetVioLetters do
  begin
    Append;
    FieldValues['letterNumber'] := newLetterNumber;
    FieldValues['letterType'] := oldLetterType;
    FieldValues['letterDate'] := Date; // Today's Date
    FieldValues['senderInitials'] := oldInitials;
    FieldValues['originalLetterDate'] := oldLetterDate;
    FieldValues['signatureLine'] := oldSignLine;
    FieldValues['specificViolationTitle'] := oldSpecVioText;
    FieldValues['textViolationWords'] := oldTextVio;
    FieldValues['remedyWords'] := oldRemedyWords;
  end;
end;

{ ----------------------------------------------------------------------+
  sbCopyVioLetterClick:
  This is a procedure to append a new record the GenVioLetters
  table which is a copy of the letter in the current record.
  +----------------------------------------------------------------------- }
procedure TMainForm.sbCopyVioLetterClick(Sender: TObject);
var
  oldLetterNumber, oldInitials: string;
  oldSignLine, oldLetterType, oldLetterDate, oldOrigLetterDate: OleVariant;
  oldSpecVioText, oldTextVio, oldRemedyWords: OleVariant;
begin
  with AdoDataSetVioLetters do
  begin
    oldLetterNumber := FieldValues['letterNumber'];
    oldLetterType := FieldValues['letterType'];
    oldLetterDate := FieldValues['letterDate'];
    oldInitials := FieldValues['senderInitials'];
    oldOrigLetterDate := FieldValues['originalLetterDate'];
    oldSignLine := FieldValues['signatureLine'];
    oldSpecVioText := FieldValues['specificViolationTitle'];
    oldTextVio := FieldValues['textViolationWords'];
    oldRemedyWords := FieldValues['remedyWords'];
  end;
  with AdoDataSetVioLetters do
  begin
    Append;
    FieldValues['letterNumber'] := oldLetterNumber;
    FieldValues['letterType'] := oldLetterType;
    FieldValues['letterDate'] := Date; // Today's Date
    FieldValues['senderInitials'] := oldInitials;
    FieldValues['originalLetterDate'] := oldOrigLetterDate;
    FieldValues['signatureLine'] := oldSignLine;
    FieldValues['specificViolationTitle'] := oldSpecVioText;
    FieldValues['textViolationWords'] := oldTextVio;
    FieldValues['remedyWords'] := oldRemedyWords;
  end;
end;

procedure TMainForm.sbAppAcctSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('houseAcct', sbAppAcctSort);
end;

procedure TMainForm.sortApplicationTable(sortString: string);
begin
  WITH ADOQuery1 do
  begin
    Active := False;
    // ShowMessage(sortString);
    SQL.Strings[0] := sortString;
    Active := True;
  end;
end;

{ ----------------------------------------------------------------------+
  AppTableButtonClear:
  Remove the glyphs from all the speedbuttons on the
  "All AC App's" tab.
  +----------------------------------------------------------------------- }
procedure TMainForm.AppTableButtonClear;
begin
  sbAppAcctSort.Glyph.Assign(nil);
  sbAppPermitSort.Glyph.Assign(nil);
  sbAppLetterSort.Glyph.Assign(nil);
  sbAppIssueDateSort.Glyph.Assign(nil);
  sbAppMonthSort.Glyph.Assign(nil);
  sbAppYearSort.Glyph.Assign(nil);
  sbAppApplyDateSort.Glyph.Assign(nil);
  sbAppApproveDateSort.Glyph.Assign(nil);
end;

{ ----------------------------------------------------------------------+
  BuildSortAppTableQuery(string; TSpeedButton):
  This is a procedure to generate a SQL query for the all
  AC Approval Letters. It takes the following arguments:

  string: The index field name be sorted
  TSpeedButton: The TSpeedButton whose glyph is to be modified

  Notice that the sort is toggled between ASCENDING and DESCENDING
  by use of the TAG property of the TSpeedButton.
  +----------------------------------------------------------------------- }
procedure TMainForm.BuildSortAppTableQuery(IndexFieldName: string;
  PushButton: TSpeedButton);
var
  queryStr: string;
begin
  // Clear all the glyphs
  AppTableButtonClear;
  // Load the new glyph
  ImageList1.GetBitmap(PushButton.Tag, PushButton.Glyph);
  queryStr := 'SELECT * FROM approvalLetters ORDER BY ' + IndexFieldName;
  // ASCENDING or DESCENDING
  if (PushButton.Tag = 1) then
    queryStr := queryStr + ' ASC'
  else
    queryStr := queryStr + ' DESC';
  // Change the TAG
  If (PushButton.Tag = 0) then
    PushButton.Tag := 1
  else
    PushButton.Tag := 0;
  sortApplicationTable(queryStr);
end;

procedure TMainForm.sbAppLetterSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('letterDate', sbAppLetterSort);
end;

procedure TMainForm.sbAppIssueDateSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('permitDateIssued', sbAppIssueDateSort);
end;

procedure TMainForm.sbAppPermitSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('permitNumber', sbAppPermitSort);
end;

procedure TMainForm.sbAppMonthSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('permitMonth', sbAppMonthSort);
end;

procedure TMainForm.sbAppYearSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('permitYear', sbAppYearSort);
end;

procedure TMainForm.sbAppApplyDateSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('applicationDate', sbAppApplyDateSort);
end;

procedure TMainForm.sbAppApproveDateSortClick(Sender: TObject);
begin
  BuildSortAppTableQuery('approvalDate', sbAppApproveDateSort);
end;

procedure TMainForm.sbLegalSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  AdoTableCurrentOwners.IndexFieldNames := 'legal';
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.sbOffsiteSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  AdoTableCurrentOwners.IndexFieldNames := 'offsite';
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);

  adoTableOwners.Filtered := False;
  adoTableOwners.IndexFieldNames := 'offsite';
  adoTableOwners.Filtered := True;
  numOfRecords := adoTableOwners.RecordCount;
  StatusBar1.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.Close1Click(Sender: TObject);
begin
  ADOConnection1.Connected := False;
end;

{ ----------------------------------------------------------------------+
  btnClearLettersClick():
  This procedure empties the AllLetters table.
  +----------------------------------------------------------------------- }
procedure TMainForm.btnClearLettersClick(Sender: TObject);
begin
  With adoQryClearLetters do
  begin
    SQL.Clear;
    SQL.Add('DELETE FROM AllLettersTable;');
    ExecSQL;
  end;
  adoTblAllLetters.Active := False;
  adoTblAllLetters.Active := True;
  statBarUpdate;
end;

{ ----------------------------------------------------------------------+
  btnOpenSqlClick():
  This procedure allows the user to select a SQL file for
  execution. The SQL file must be 100% self-contained.
  +----------------------------------------------------------------------- }
procedure TMainForm.btnOpenSqlClick(Sender: TObject);
var
  i: integer;
begin
  if (openDlgSql.Execute) then
  begin
    with adoQryClearLetters do
    begin
      SQL.Clear;
      SQL.LoadFromFile(openDlgSql.FileName);
      for i := (SQL.Count - 1) downto 0 do
      begin
        if (AnsiLeftStr(SQL[i], 2) = '/*') then
          SQL.Delete(i);
      end;
      SQL.SaveToFile('hereIsTheSql.sql');
      ExecSQL;
      adoTblAllLetters.Active := False;
      adoTblAllLetters.Active := True;
      RichEdit1.Lines.Clear;
      RichEdit1.Lines.AddStrings(DBMemo14.Lines);
    end;
  end;
end;

{ ----------------------------------------------------------------------+
  btnSaveLetterClick():
  This procedure saves the current letter displayed in the
  RichEdit1 component to an RTF file. The default filename has the
  following format:
  "street_address YYYY-MM-DD" with leading zeros on the MM & DD

  The leading zeros are necessary to display the files in
  chronlogical order when the user looks at the file names.
  +----------------------------------------------------------------------- }
procedure TMainForm.btnSaveLetterClick(Sender: TObject);
var
  letterFileName, letterAddress: string;
  theLetterDate: TDateTime;
  theYear, theMonth, theDay: word;
  theLDate: string;
begin
  letterAddress := adoTblAllLetters.FieldValues['address'];
  theLetterDate := adoTblAllLetters.FieldValues['letterDate'];
  DecodeDate(theLetterDate, theYear, theMonth, theDay);
  theLDate := IntToStr(theYear) + Format('-%0.2d-%0.2d', [theMonth, theDay]);
  letterFileName := letterAddress + ' ' + theLDate;
  SaveDialog.FileName := letterFileName;
  if (SaveDialog.Execute) then
  begin
    RichEdit1.Lines.SaveToFile(SaveDialog.FileName);
  end;
  // Blank out the file name
  SaveDialog.FileName := '';
end;

procedure TMainForm.btnFontClick(Sender: TObject);
begin
  if (FontDialog1.Execute) then
    RichEdit1.selAttributes.Assign(FontDialog1.font);
end;

procedure TMainForm.DBGrid8CellClick(Column: TColumn);
var
  presX, presY: integer;
  r: TRect;
begin
  presX := GetDeviceCaps(printer.handle, LOGPIXELSX);
  presY := GetDeviceCaps(printer.handle, LOGPIXELSY);
  // ShowMessage('presX =' + IntToStr(presX));
  // ShowMessage('presY =' + IntToStr(presY));
  with r do
  begin
    Left := 5 * presX; // 1.0" ft Margin
    Top := 25 * presY div 2; // 2.5" Top Margin
    Right := printer.PageWidth - (3 * presX div 4); // 0.75" Right Margin
    Bottom := printer.PageHeight - (3 * presY div 2); // 1.50" Bottom Margin
    // ShowMessage('left:' + IntToStr(left) +'     Top:' + IntToStr(Top) +'      Right:' + IntToStr(Right) +'     Bottom:' + IntToStr(Bottom));
  end;
  RichEdit1.PageRect := r;
  with RichEdit1.Lines do
  begin
    Clear;
    AddStrings(DBMemo14.Lines);
    VertScrollBar.Position := 0;
  end;
  statBarUpdate;
end;

procedure TMainForm.sbVioNumClick(Sender: TObject);
begin
  SortColumn(adoTblBrowseGenVioLetters, 'violationId', sbVioNum);
end;

procedure TMainForm.sbAcctNumClick(Sender: TObject);
begin
  SortColumn(adoTblBrowseGenVioLetters, 'houseAcct', sbAcctNum);
end;

procedure TMainForm.sbLetterNumClick(Sender: TObject);
begin
  SortColumn(adoTblBrowseGenVioLetters, 'letterNumber', sbLetterNum);
end;

procedure TMainForm.sbDateClick(Sender: TObject);
begin
  SortColumn(adoTblBrowseGenVioLetters, 'letterDate', sbDate);
end;

procedure TMainForm.sbOrigLetterDateClick(Sender: TObject);
begin
  SortColumn(adoTblBrowseGenVioLetters, 'originalLetterDate', sbOrigLetterDate);
end;

procedure TMainForm.sbLetterTypeClick(Sender: TObject);
begin
  SortColumn(adoTblBrowseGenVioLetters, 'letterType', sbLetterType);
end;

procedure TMainForm.dbGridBrowseGenVioLettersCellClick(Column: TColumn);
begin
  eCurrentAcctSearch.Text := adoTblBrowseGenVioLetters.FieldValues['houseAcct'];
  eCurrentAcctSearchChange(Column);
end;

{ ----------------------------------------------------------------------+
  DBGrid3CellClick():
  This procedure finds the houseAcct value of the current record
  in the All_AC_Apps grid and copies it to the "Acct Search" TEdit
  box. It then calls the eCurrentAcctSearchChange event.
  +----------------------------------------------------------------------- }
procedure TMainForm.DBGrid3CellClick(Column: TColumn);
begin
  UpdateCurrentHouseAcct(ADOQuery1.FieldValues['houseAcct']);
end;

procedure TMainForm.adoTableOwnersAfterScroll(DataSet: TDataSet);
begin
  if (adoTableOwners.Active and AdoTableCurrentOwners.Active) then
    UpdateCurrentHouseAcct(adoTableOwners.FieldValues['houseAcct']);
end;

procedure TMainForm.adoTblAllLettersAfterScroll(DataSet: TDataSet);
begin
  if (adoTblAllLetters.Active and AdoTableCurrentOwners.Active) then begin
    UpdateCurrentHouseAcct(adoTblAllLetters.FieldValues['houseAcct']);
    with RichEdit1.Lines do
      begin
        Clear;
        AddStrings(DBMemo14.Lines);
        VertScrollBar.Position := 0;
      end;
    statBarUpdate;
  end;
end;

procedure TMainForm.adoTblBrowseGenVioLettersAfterScroll(DataSet: TDataSet);
begin
  UpdateCurrentHouseAcct(adoTblBrowseGenVioLetters.FieldValues['houseAcct']);
end;

procedure TMainForm.adoTblMemoToLegalAfterScroll(DataSet: TDataSet);
begin
  if (adoTblMemoToLegal.Active and AdoTableCurrentOwners.Active) then
    UpdateCurrentHouseAcct(adoTblMemoToLegal.FieldValues['houseAcct']);
end;

procedure TMainForm.adoTblMemoToLegalNewRecord(DataSet: TDataSet);
var
  houseAccount: string;
begin
  houseAccount := AdoTableCurrentOwners.FieldValues['houseAcct'];
  adoTblMemoToLegal.FieldValues['houseAcct'] := houseAccount;
  adoTblMemoToLegal.FieldValues['memoDate'] := Date;
end;

procedure TMainForm.adoTblWelcomeLettersAfterInsert(DataSet: TDataSet);
begin
  with adoTblWelcomeLetters do
  begin
    FieldValues['letterDate'] := Date;
  end;
end;

procedure TMainForm.RemoveGlyphImages(Sender: TObject);
begin
  SpeedButton1.Glyph.Assign(nil);
  SpeedButton2.Glyph.Assign(nil);
  SpeedButton3.Glyph.Assign(nil);
end;

procedure TMainForm.SpeedButton1Click(Sender: TObject);
begin
  RemoveGlyphImages(Sender);
  SortColumn(adoTblMemoToLegal, 'houseAcct', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton2Click(Sender: TObject);
begin
  RemoveGlyphImages(Sender);
  SortColumn(adoTblMemoToLegal, 'memoDate', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton3Click(Sender: TObject);
begin
  RemoveGlyphImages(Sender);
  SortColumn(adoTblMemoToLegal, 'votingDate', Sender as TSpeedButton);
end;

procedure TMainForm.DBGrid1CellClick(Column: TColumn);
begin
  eCurrentAcctSearch.Text := adoTblMemoToLegal.FieldValues['houseAcct'];
  eCurrentAcctSearchChange(Column);
end;

{ ----------------------------------------------------------------------+
  SortColumn(TAdoTable; string; TSpeedButton):
  This is a generic procedure to apply a sort criterium to a
  TAdoTable. It takes the following arguments:

  TAdoTable: The TAdoTable object to be sorted
  string: The column name to be sorted
  TSpeedButton: The TSpeedButton whose glyph is to be modified

  Notice that the sort is toggled between ASCENDING and DESCENDING
  by use of the TAG property of the TSpeedButton.
  +----------------------------------------------------------------------- }
procedure TMainForm.SortColumn(DataTable: TADOTable; IndexFieldName: string;
  PushButton: TSpeedButton);
begin
  DataTable.Filtered := False;
  // TODO -- Remove all the speedbutton glyphs here
  PushButton.Glyph.Assign(nil);
  if (PushButton.Tag = 1) then
  begin
    DataTable.IndexFieldNames := IndexFieldName + ' ASC';
    ImageList1.GetBitmap(1, PushButton.Glyph);
    PushButton.Tag := 0;
  end
  else
  begin
    DataTable.IndexFieldNames := IndexFieldName + ' DESC';
    ImageList1.GetBitmap(0, PushButton.Glyph);
    PushButton.Tag := 1;
  end;
  DataTable.Filtered := True;
end;

{ ----------------------------------------------------------------------+
  SortColumn(TADODataSet; string; TSpeedButton):
  This is a generic procedure to apply a sort criterium to a
  TAdoTable. It takes the following arguments:

  TADODataSet: The TADODataSet object to be sorted
  string: The column name to be sorted
  TSpeedButton: The TSpeedButton whose glyph is to be modified

  Notice that the sort is toggled between ASCENDING and DESCENDING
  by use of the TAG property of the TSpeedButton.
  +----------------------------------------------------------------------- }
procedure TMainForm.SortColumn(DataSet: TADODataSet; IndexFieldName: string;
  PushButton: TSpeedButton);
begin
  DataSet.Filtered := False;
  // TODO -- Remove all the speedbutton glyphs here
  PushButton.Glyph.Assign(nil);
  if (PushButton.Tag = 1) then
  begin
    DataSet.IndexFieldNames := IndexFieldName + ' ASC';
    ImageList1.GetBitmap(1, PushButton.Glyph);
    PushButton.Tag := 0;
  end
  else
  begin
    DataSet.IndexFieldNames := IndexFieldName + ' DESC';
    ImageList1.GetBitmap(0, PushButton.Glyph);
    PushButton.Tag := 1;
  end;
  DataSet.Filtered := True;
end;

{ ----------------------------------------------------------------------+
  GenerateLetterButtonClear:
  Remove the glyphs from all the speedbuttons on the
  GenerateLetters tab.
  +----------------------------------------------------------------------- }
procedure TMainForm.GenerateLetterButtonClear;
begin
  SpeedButton8.Glyph.Assign(nil);
  SpeedButton9.Glyph.Assign(nil);
  SpeedButton10.Glyph.Assign(nil);
  SpeedButton11.Glyph.Assign(nil);
  SpeedButton12.Glyph.Assign(nil);
  SpeedButton13.Glyph.Assign(nil);
  SpeedButton14.Glyph.Assign(nil);
  SpeedButton16.Glyph.Assign(nil);
end;

procedure TMainForm.SpeedButton8Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'houseAcct', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton9Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'violationId', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton10Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'letterDate', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton11Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'letterType', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton16Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'permitNumber', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton12Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'owner', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton13Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'address', Sender as TSpeedButton);
end;

procedure TMainForm.SpeedButton14Click(Sender: TObject);
begin
  GenerateLetterButtonClear;
  SortColumn(adoTblAllLetters, 'letterNumber', Sender as TSpeedButton);
end;

procedure TMainForm.sbWelcomeCloseClick(Sender: TObject);
begin
  SortColumn(AdoDataSetOwners, 'closeDate', sbWelcomeClose);
end;

procedure TMainForm.sbWelcomeIdClick(Sender: TObject);
begin
  SortColumn(AdoDataSetOwners, 'OwnerId', sbWelcomeId);
end;

procedure TMainForm.sbWelcomeOwnerClick(Sender: TObject);
begin
  SortColumn(AdoDataSetOwners, 'owner', sbWelcomeOwner);
end;

procedure TMainForm.btnVioStatusClick(Sender: TObject);
var
  myHouse: integer;
  queryString: String;
begin
  AdoDataSetViolations.Active := False;
  AdoDataSetViolations.CommandText := '';
  queryString := 'SELECT * FROM Violations WHERE ';
  myHouse := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  if (btnVioStatus.Tag = 0) then
  begin
    btnVioStatus.Tag := 1;
    btnVioStatus.Caption := 'Open';
    queryString := queryString + '(closeDate IS NULL) AND ';
  end
  else if (btnVioStatus.Tag = 1) then
  begin
    btnVioStatus.Tag := 2;
    btnVioStatus.Caption := 'Closed';
    queryString := queryString + '(closeDate IS NOT NULL) AND ';
  end
  else
  begin
    btnVioStatus.Tag := 0;
    btnVioStatus.Caption := 'All';
  end;
  queryString := queryString + '(houseAcct = ' + IntToStr(myHouse) +
    ' ) ORDER BY houseAcct, openDate;';
  AdoDataSetViolations.CommandText := queryString;
  AdoDataSetViolations.Active := True;
  sbViolations.Panels[0].Text := 'Total Records: ' +
    IntToStr(AdoDataSetViolations.RecordCount);
end;

{ ----------------------------------------------------------------------+
  AdoTableCurrentOwnersAfterScroll(TDataSet):
  When the user scrolls to a new CurrentOwner record the Violation
  and the Violation Status dbGrids have be to updated to the
  new AcctNumber.

  Notice that we have to test for the case when there are no
  existing AcctNumber because of an incomplete field entry.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoTableCurrentOwnersAfterScroll(DataSet: TDataSet);
var
  myHouse: integer;
  queryString: String;
  Val: OleVariant;
begin
  // Cast the field as OleVariant to test for NULL value
  Val := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  if (VarIsNull(Val)) then
    myHouse := -999 // Set to non-existant value (dummy value)
  else
    myHouse := integer(Val);

  // Deactivate the Violations dataset and change its query according
  // to the All/Open/Closed selection, and then reactivate the query
  AdoDataSetViolations.Active := False;
  AdoDataSetViolations.CommandText := '';
  queryString := 'SELECT * FROM Violations WHERE ';
  if (btnVioStatus.Tag = 1) then
  begin
    btnVioStatus.Caption := 'Open';
    queryString := queryString + '(closeDate IS NULL) AND ';
  end
  else if (btnVioStatus.Tag = 2) then
  begin
    btnVioStatus.Caption := 'Closed';
    queryString := queryString + '(closeDate IS NOT NULL) AND ';
  end
  else
  begin
    btnVioStatus.Caption := 'All';
  end;
  queryString := queryString + '(houseAcct = ' + IntToStr(myHouse) +
    ' ) ORDER BY houseAcct, openDate;';
  AdoDataSetViolations.CommandText := queryString;
  AdoDataSetViolations.Active := True;

  // If there are no Violation records then we have to make a dummy SQL
  // query which is quarenteed to return an empty dataset.
  if (AdoDataSetViolations.RecordCount < 1) then
  begin
    AdoDataSetVioStatus.Active := False;
    AdoDataSetVioStatus.CommandText := '';
    queryString := 'SELECT * FROM ViolationStatus WHERE ';
    queryString := queryString + '(violationNumber = -999' +
      ' ) ORDER BY statusDate;';
    AdoDataSetVioStatus.CommandText := queryString;
    AdoDataSetVioStatus.Active := True;
    AdoDataSetVioLetters.Active := False;
  end;

  // Change the query statement for the Owners grid on the
  // Welcome tabsheet
  AdoDataSetOwners.Active := False;
  AdoDataSetOwners.CommandText := '';
  AdoDataSetOwners.CommandText := 'SELECT * FROM Owners WHERE houseAcct = ' +
    IntToStr(myHouse) + ' ORDER BY closeDate DESC;';
  AdoDataSetOwners.Active := True;
  // Change the query statement for the AC Application grid
  with AdoDataAllAppLetters do
  begin
    Active := False;
    CommandText := 'SELECT * FROM  ApprovalLetters WHERE houseAcct = ' +
      IntToStr(myHouse) + ' ORDER BY houseAcct DESC, applicationDate DESC;';
    Active := True;
  end;

  // Set the Status Bars text as appropriate
  sbViolations.Panels[0].Text := 'Total Records: ' +
    IntToStr(AdoDataSetViolations.RecordCount);
  sbOwners.Panels[0].Text := 'Total Records: ' +
    IntToStr(AdoDataSetOwners.RecordCount);
end;

{ ----------------------------------------------------------------------+
  AdoDataSetViolationsAfterScroll(TDataSet):
  When the user scrolls to a new Violation record the Violation
  Status and the ViolationLetters dbGrids have be to updated to the
  new ViolationNumber.

  Notice that we have to test for the case when there are no
  existing violations on the houseAcct.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetViolationsAfterScroll(DataSet: TDataSet);
var
  myViolation, statusRecordNumber: integer;
  queryString: String;
  Val: OleVariant;
begin
  // Cast the field as OleVariant to test for NULL value
  Val := dbGridViolations.DataSource.DataSet.FieldValues['violationID'];
  if (VarIsNull(Val)) then
    myViolation := -999
  else
    myViolation := integer(Val);
  // Build the SQL query with the appropriate value for the Vio Number
  queryString := 'SELECT * FROM ViolationStatus WHERE ';
  queryString := queryString + '(violationNumber = ' + IntToStr(myViolation) +
    ' ) ORDER BY statusDate;';
  with AdoDataSetVioStatus do
  begin
    Active := False;
    CommandText := '';
    CommandText := queryString;
    Active := True;
    statusRecordNumber := AdoDataSetVioStatus.RecordCount;
    StatusBar2.Panels[0].Text := 'Record Count: ' +
      IntToStr(statusRecordNumber);
  end;

  // Now update the Violation Letter dbGrid
  queryString := 'SELECT * FROM GenericViolationLetters WHERE ';
  queryString := queryString + '(violationID = ' + IntToStr(myViolation) +
    ' ) ORDER BY letterDate DESC;';
  AdoDataSetVioLetters.Active := False;
  AdoDataSetVioLetters.CommandText := '';
  AdoDataSetVioLetters.CommandText := queryString;
  AdoDataSetVioLetters.Active := True;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetVioStatusAfterInsert(TDataSet):
  When the user inserts a new Violation Status record.
  It copies the current ViolationNumber and pastes this into the
  new record, along with today's date.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetVioStatusAfterInsert(DataSet: TDataSet);
var
  myVioNumber: integer;
begin
  myVioNumber := dbGridViolations.DataSource.DataSet.FieldValues['violationID'];
  with AdoDataSetVioStatus do
  begin
    FieldValues['violationNumber'] := myVioNumber;
    FieldValues['statusDate'] := Date;
    FieldValues['statusBy'] := 'BOZO';
  end;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetViolationsAfterInsert(TDataSet):
  When the user inserts a new Violation record.
  It copies the current HouseAcct and pastes this into the
  new record, along with today's date.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetViolationsAfterInsert(DataSet: TDataSet);
var
  myHouseAcct: integer;
begin
  myHouseAcct := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  with AdoDataSetViolations do
  begin
    FieldValues['houseAcct'] := myHouseAcct;
    FieldValues['violationDate'] := Date;
    FieldValues['startDate'] := Date;
    FieldValues['openDate'] := Date;
    FieldValues['reportedBy'] := 'BOZO';
    FieldValues['Reason'] := 'Reason';
  end;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetVioLettersAfterInsert(TDataSet):
  When the user inserts a new record into the Violation Letters
  AdoDataSet the houseAcct and ViolationNumber fields are copied
  from the CurrentOwners and Violations datasets respectively and
  copied into the appropriate fields in the Violation Letters table.
  the letterDate is set to the current date and the sender initials
  are set to the default value.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetVioLettersAfterInsert(DataSet: TDataSet);
var
  myHouseAcct: integer;
  myVioNumber: integer;
begin
  // ShowMessage('After Insert');
  myHouseAcct := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  myVioNumber := dbGridViolations.DataSource.DataSet.FieldValues['violationID'];
  with AdoDataSetVioLetters do
  begin
    FieldValues['houseAcct'] := myHouseAcct;
    FieldValues['violationId'] := myVioNumber;
    FieldValues['letterDate'] := Date;
    FieldValues['senderInitials'] := 'BOZO';
  end;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetViolationsAfterPost(TDataSet):
  When the user posts a new ViolationStatus record the AdoDataSet
  has to be toggled so that the command line, that is the SQL Query,
  is executed and the record set is updated.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetViolationsAfterPost(DataSet: TDataSet);
begin
  AdoDataSetViolations.Active := False;
  AdoDataSetViolations.Active := True;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetVioStatusAfterPost(TDataSet):
  When the user posts a new ViolationStatus record the AdoDataSet
  has to be toggled so that the command line, that is the SQL Query,
  is executed and the record set is updated.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetVioStatusAfterPost(DataSet: TDataSet);
begin
  AdoDataSetVioStatus.Active := False;
  AdoDataSetVioStatus.Active := True;
end;

procedure TMainForm.ADOQuery1AfterScroll(DataSet: TDataSet);
begin
  UpdateCurrentHouseAcct(ADOQuery1.FieldValues['houseAcct']);
end;

{ ----------------------------------------------------------------------+
  btnRunLettersClick(TDataSet):
  When the user posts a new ViolationStatus record the AdoDataSet
  has to be toggled so that the command line, that is the SQL Query,
  is executed and the record set is updated.
  +----------------------------------------------------------------------- }
procedure TMainForm.btnRunLettersClick(Sender: TObject);
var
  thisLetterNum, thisLetterType: string;
  OnOffSite: string;
  sqlText: TStringList;
  sqlDirectory, rejectType: string;
  i, j, k: integer;
begin
  { ---------------------------------------------------------------+
    | Retrieve all the L_Type & L_Number                            |
    +--------------------------------------------------------------- }
  with adoQryRunLetters do
  begin
    SQL.Clear;
    SQL.Append
      ('SELECT DISTINCT letterNumber, letterType FROM GenericViolationLetters ');
    SQL.Append('WHERE letterDate >= #' + DateToStr(LetterDate) + '#');
    SQL.Append('ORDER BY letterType, letterNumber;');
    ExecSQL;
    Active := True;
  end;
  RichEdit1.Lines.Clear;

  sqlDirectory :=
    'C:\Users\Mandy.SSCA-FRONTOFFIC\Documents\SSCA Database\SQL Queries\delphi queries\Auto Format\';
  sqlText := TStringList.Create;

  { ---------------------------------------------------------------+
    |  Write the letters to be generated to the RichEdit            |
    +--------------------------------------------------------------- }
  {
    adoQryRunLetters.First;
    for i := 1 to adoQryRunLetters.RecordCount do
    begin
    thisLetterNum  := adoQryRunLetters.FieldValues['letterNumber'];
    thisLetterType := adoQryRunLetters.FieldValues['letterType'];
    RichEdit1.Lines.Append(IntToStr(i) + ': ' + thisLetterType + '_' + thisLetterNum);
    adoQryRunLetters.Next;
    end;
  }
  { ---------------------------------------------------------------+
    |  Run thru the letters: 1st for OnSite, then for OffSite       |
    +--------------------------------------------------------------- }
  OnOffSite := 'OnSite';
  for k := 1 to 2 do
  begin
    // Start at the beginning. . . . .
    adoQryRunLetters.First;
    { Now we run thru the letters }
    for i := 1 to adoQryRunLetters.RecordCount do
    begin
      thisLetterNum := adoQryRunLetters.FieldValues['letterNumber'];
      thisLetterType := adoQryRunLetters.FieldValues['letterType'];
      with adoQryClearLetters do
      begin
        SQL.Clear;
        if (thisLetterNum = '209') then
        begin
          sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'CertHeader.SQL');
          SQL.AddStrings(sqlText);
          sqlText.LoadFromFile(sqlDirectory + 'GEN_209.SQL');
          SQL.AddStrings(sqlText);
          sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'From209.SQL');
          SQL.AddStrings(sqlText);
        end
        else if (thisLetterNum = '4OL') then
        begin
          sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'CertHeader.SQL');
          SQL.AddStrings(sqlText);
          sqlText.LoadFromFile(sqlDirectory + thisLetterType + '_' +
            thisLetterNum + '.SQL');
          SQL.AddStrings(sqlText);
          sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'From.SQL');
          SQL.AddStrings(sqlText);
        end
        else
        begin
          sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'Header.SQL');
          SQL.AddStrings(sqlText);
          if (TRIM(thisLetterType) = 'STAT') then
            sqlText.LoadFromFile(sqlDirectory + thisLetterType + '_STAT.SQL')
          else
            sqlText.LoadFromFile(sqlDirectory + thisLetterType + '_' +
              thisLetterNum + '.SQL');
          SQL.AddStrings(sqlText);
          sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'From.SQL');
          SQL.AddStrings(sqlText);
        end;
        { Delete the comment lines from the SQL lines }
        for j := (SQL.Count - 1) downto 0 do
          if (AnsiLeftStr(SQL[j], 2) = '/*') then
            SQL.Delete(j);
        { Save the SQL to a local file for troubleshooting purposes }
        // ShowMessage(thisLetterType + '_' + thisLetterNum + '   : ' + IntToStr(adoQryClearLetters.Parameters.Count) );
        SQL.SaveToFile(sqlDirectory + 'ZZ_' + thisLetterType + '_' +
          thisLetterNum + '.SQL');
        // Prepared := True;
        if (thisLetterNum <> '209') then
        begin
          Parameters.ParamByName('L_TYPE').Value := thisLetterType;
          Parameters.ParamByName('L_NUMBER').Value := thisLetterNum;
        end;
        Prepared := True;
        ExecSQL;
        { Do the next letter }
        adoQryRunLetters.Next;
      end; // with adoQryClearLetters
    end; // For i
    adoTblAllLetters.Active := False;
    adoTblAllLetters.Active := True;

    RichEdit1.Lines.Clear;
    // Now we run the OFFSITE letters
    OnOffSite := 'OffSite';
  end; // for k
  adoQryRunLetters.Active := False;
  adoQryRunLetters.SQL.Clear;

  { Do the AC Application Approval Letters }
  OnOffSite := 'OnSite';
  for k := 1 to 2 do
  begin
    with adoQryClearLetters do
    begin
      SQL.Clear;
      sqlText.LoadFromFile(sqlDirectory + 'ApprovalHeader.SQL');
      SQL.AddStrings(sqlText);
      sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'Approval.SQL');
      SQL.AddStrings(sqlText);
      sqlText.LoadFromFile(sqlDirectory + OnOffSite + 'ApprovalFrom.SQL');
      SQL.AddStrings(sqlText);
      { Delete the comment lines from the SQL lines }
      for j := (SQL.Count - 1) downto 0 do
        if (AnsiLeftStr(SQL[j], 2) = '/*') then
          SQL.Delete(j);
      { Save the SQL to a local file for troubleshooting purposes }
      SQL.SaveToFile(sqlDirectory + 'ZZ_' + OnOffSite + 'Approval.SQL');
      Prepared := True;
      ExecSQL;
    end;
    // Now we run the OFFSITE letters
    OnOffSite := 'OffSite';
  end; // for k

  { Do the AC Application Rejection Letters }
  OnOffSite := 'OnSite';
  for k := 1 to 2 do
  begin
    rejectType := '22';
    for i := 1 to 5 do
    begin
      with adoQryClearLetters do
      begin
        SQL.Clear;
        sqlText.LoadFromFile(sqlDirectory + 'ApprovalHeader.SQL');
        SQL.AddStrings(sqlText);
        sqlText.LoadFromFile(sqlDirectory + 'Reject' + rejectType + OnOffSite
          + '.SQL');
        SQL.AddStrings(sqlText);
        sqlText.LoadFromFile(sqlDirectory + 'Reject' + rejectType + OnOffSite +
          'From.SQL');
        SQL.AddStrings(sqlText);
        { Delete the comment lines from the SQL lines }
        for j := (SQL.Count - 1) downto 0 do
          if (AnsiLeftStr(SQL[j], 2) = '/*') then
            SQL.Delete(j);
        { Save the SQL to a local file for troubleshooting purposes }
        SQL.SaveToFile(sqlDirectory + 'ZZ_' + 'Reject' + rejectType + OnOffSite
          + '.SQL');
        Prepared := True;
        ExecSQL;
      end; // with adoQryClearLetters
      if (i = 1) then
        rejectType := '24'
      else if (i = 2) then
        rejectType := '33'
      else if (i = 3) then
        rejectType := '55'
      else
        rejectType := '10';
    end; // for i
    // Now we run the OFFSITE letters
    OnOffSite := 'OffSite';
  end; // for k

  { Do the Welcome Letters }
  OnOffSite := 'OnSite';
  for k := 1 to 3 do
  begin
    with adoQryClearLetters do
    begin
      SQL.Clear;
      sqlText.LoadFromFile(sqlDirectory + 'WelcomeHeader.SQL');
      SQL.AddStrings(sqlText);
      sqlText.LoadFromFile(sqlDirectory + 'Welcome' + OnOffSite + '.SQL');
      SQL.AddStrings(sqlText);
      sqlText.LoadFromFile(sqlDirectory + 'Welcome' + OnOffSite + 'From.SQL');
      SQL.AddStrings(sqlText);
      { Delete the comment lines from the SQL lines }
      for j := (SQL.Count - 1) downto 0 do
        if (AnsiLeftStr(SQL[j], 2) = '/*') then
          SQL.Delete(j);
      { Save the SQL to a local file for troubleshooting purposes }
      SQL.SaveToFile(sqlDirectory + 'ZZ_' + 'Welcome' + OnOffSite + '.SQL');
      Prepared := True;
      ExecSQL;
    end; // with adoQryClearLetters
    // Now we run the OFFSITE letters
    if (k = 1) then
      OnOffSite := 'OffSite'
    else
      OnOffSite := 'Investor'
  end; // for k
  // do something here
  sqlText.Free;
  { Cull the letters according to the LetterDate }
  if (LetterDate > 1) then
  begin
    adoQryRunLetters.SQL.Clear;
    adoQryRunLetters.SQL.Append
      ('DELETE * from AllLettersTable WHERE letterDate < #' +
      DateToStr(LetterDate) + '# ;');
    adoQryRunLetters.ExecSQL;
  end;

  adoTblAllLetters.Active := False;
  adoTblAllLetters.Active := True;
  statBarUpdate;
  RichEdit1.Lines.AddStrings(DBMemo14.Lines);

end;

{ ----------------------------------------------------------------------+
  menuLettersClick(Sender):
  This procedure is called when the user changes the letter span.

  The day span is contained in the TAG property of the MenuItem.
  If the user wants ALL letters generated then the LetterDate is
  set to the value of 'zero' otherwise it is set to the beginning
  date of the timeperiod requested.

  Finally the MenuItem.Checked property is set to TRUE.
  +----------------------------------------------------------------------- }
procedure TMainForm.menuLettersClick(Sender: TObject);
begin
  LetterDate := Now - (Sender as TMenuItem).Tag + 1;
  if ((Sender as TMenuItem).Tag = 0) then
    LetterDate := 1;
  menuLettersSetCheck(Sender);
end;

{ ----------------------------------------------------------------------+
  menuLettersUncheck;
  This procedure is unchecks all Letter MenuItems.
  +----------------------------------------------------------------------- }
procedure TMainForm.menuLettersUncheck;
begin
  menuLettersToday.Checked := False;
  menuLetters2Days.Checked := False;
  menuLetters3Days.Checked := False;
  menuLetters7Days.Checked := False;
  menuLetters2Weeks.Checked := False;
  menuLetters30Days.Checked := False;
  menuLetters60Days.Checked := False;
  menuLetters90Days.Checked := False;
  menuLettersAll.Checked := False;
end;

{ ----------------------------------------------------------------------+
  menuLettersUncheck;
  This procedure sets the default Letter MenuItem.
  +----------------------------------------------------------------------- }
procedure TMainForm.menuLettersSetDefault;
begin
  menuLettersUncheck;
  menuLettersToday.Checked := True;
  LetterDate := Now;
  menuLettersSetCheck(menuLettersToday);
end;

{ ----------------------------------------------------------------------+
  menuLettersClick(Sender):
  This procedure is called by menuLettersClick(Sender)

  First all the MenuItems are UNCHECKED.
  Then the chosen MenuItem is CHECKED.
  +----------------------------------------------------------------------- }
procedure TMainForm.menuLettersSetCheck(Sender: TObject);
var
  dayString: string;
begin
  if ((Sender as TMenuItem).Tag = 1) then
    dayString := 'Today'
  else
    dayString := IntToStr((Sender as TMenuItem).Tag) + ' days';

  if ((Sender as TMenuItem).Tag > 0) then
    statBarGenLetters.Panels[0].Text := 'Time: ' + dayString
  else
    statBarGenLetters.Panels[0].Text := 'Time: All Letters';

  menuLettersUncheck;
  (Sender as TMenuItem).Checked := True;
  statBarUpdate;
end;

procedure TMainForm.statBarUpdate;
var
  recordCount, recordNumber: Integer;
begin
  if (adoTblAllLetters.Active) then begin
    recordCount := adoTblAllLetters.RecordCount;
    recordNumber := DataSource4.DataSet.RecNo
  end else begin
    recordCount := 0;
    recordNumber := 0
  end;
  statBarGenLetters.Panels[1].Text := 'Total Letters: ' + IntToStr(recordCount);
  statBarGenLetters.Panels[2].Text := 'Record #: ' + IntToStr(recordNumber);
  statBarGenLetters.Panels[3].Text := 'Letter Date >= ' + DateToStr(LetterDate);
end;

procedure TMainForm.eCurrentRouteSearchChange(Sender: TObject);
var
  numOfRecords: integer;
begin
  if (length(eCurrentRouteSearch.Text) = 0) then
  begin
    AdoTableCurrentOwners.Filter := '';
    AdoTableCurrentOwners.Filtered := True;
  end
  else
  begin
    AdoTableCurrentOwners.Filtered := False;
    AdoTableCurrentOwners.Filter := 'driveRoute like ' + '''' +
      eCurrentRouteSearch.Text + '%' + '''';
    AdoTableCurrentOwners.Filtered := True;
  end;
  sbCurrentOwners.Panels[0].Text := AdoTableCurrentOwners.Filter;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.sbRouteSortClick(Sender: TObject);
var
  numOfRecords: integer;
begin
  AdoTableCurrentOwners.Filtered := False;
  sbRouteSort.Glyph.Assign(nil);
  if (sbStrNumSort.Tag = 1) then
  begin
    AdoTableCurrentOwners.IndexFieldNames := 'driveRoute ASC';
    ImageList1.GetBitmap(1, sbRouteSort.Glyph);
    sbStrNumSort.Tag := 0;
  end
  else
  begin
    AdoTableCurrentOwners.IndexFieldNames := 'driveRoute DESC';
    ImageList1.GetBitmap(0, sbRouteSort.Glyph);
    sbStrNumSort.Tag := 1;
  end;
  AdoTableCurrentOwners.Filtered := True;
  numOfRecords := AdoTableCurrentOwners.RecordCount;
  sbCurrentOwners.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.FormResize(Sender: TObject);
var
  formHeight: double;
begin
  formHeight := pnlCurrentOwners.Height * 0.36144;
  pnlCurrentOwners.Height := floor(formHeight);

  // := doubleMainForm.Height *
end;

procedure TMainForm.pnlCurrentOwnersResize(Sender: TObject);
var
  margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  bdGridCurrentOwners.Left := margin;
  bdGridCurrentOwners.Width := pnlCurrentOwners.Width - 2 * margin;
  bdGridCurrentOwners.Height := pnlCurrentOwners.Height -
    bdGridCurrentOwners.Top - sbCurrentOwners.Height - margin;
end;

procedure TMainForm.Connect1Click(Sender: TObject);
var
  sConn, sSearch: WideString;
  sourceStart, sourceEnd: integer;
begin
  sConn := PromptDataSource(MainForm.handle, 'aaa');
  // sConn := ADOConnection1.ConnectionString;
  ShowMessage(sConn);
  sSearch := 'Data Source=';
  sourceStart := Pos(sSearch, sConn, 0);
  sourceEnd := PosEx(';', sConn, sourceStart + 1);
  ShowMessage('start = ' + IntToStr(sourceStart) + '    end = ' +
    IntToStr(sourceEnd));

end;

{ ----------------------------------------------------------------------+
  Panel1Resize(Sender):
  This procedure is called when Panel1 needs to be resized.
  +----------------------------------------------------------------------- }
procedure TMainForm.pnlAcAppEnterResize(Sender: TObject);
var
  memoSpacing, bottomSpacing, gutter, margin: integer;
begin
  memoSpacing := 24;
  bottomSpacing := 10;
  gutter := 30;
  margin := 5;
  dbGridApprovalLetters.Height := pnlAcAppEnter.Height -
    dbGridApprovalLetters.Top - margin;
  dbGridApprovalLetters.Width :=
    floor((pnlAcAppEnter.Width - dbGridApprovalLetters.Left - bottomSpacing)
    * 0.50987);
  lblAcApplications.Left :=
    (dbGridApprovalLetters.Width - dbGridApprovalLetters.Left -
    lblAcApplications.Width) div 2;
  DBNavigator5.Left := (dbGridApprovalLetters.Width - dbGridApprovalLetters.Left
    - DBNavigator5.Width) div 2;
  // Set the Left edge for the Memo fields and their Labels
  lblAcAppSpecAppWords.Left := dbGridApprovalLetters.Left +
    dbGridApprovalLetters.Width + gutter;
  lblAcAppSpecRejectWords.Left := lblAcAppSpecAppWords.Left;
  lblAcAppSpecRemedyWords.Left := lblAcAppSpecAppWords.Left;
  dbMemoAcAppApprovalWords.Left := lblAcAppSpecAppWords.Left;
  dbMemoAcAppRejectWords.Left := lblAcAppSpecAppWords.Left;
  dbMemoAcAppRemedyWords.Left := lblAcAppSpecAppWords.Left;
  // Set the width of the Memo fields
  dbMemoAcAppApprovalWords.Width := pnlAcAppEnter.Width - 2 * margin -
    dbGridApprovalLetters.Width - gutter;
  dbMemoAcAppRejectWords.Width := dbMemoAcAppApprovalWords.Width;
  dbMemoAcAppRemedyWords.Width := dbMemoAcAppApprovalWords.Width;
  // Set the Height of the Memo fields
  dbMemoAcAppApprovalWords.Height :=
    (pnlAcAppEnter.Height - 3 * lblAcAppSpecAppWords.Height - 3 * memoSpacing -
    margin) div 3;
  dbMemoAcAppRejectWords.Height := dbMemoAcAppApprovalWords.Height;
  dbMemoAcAppRemedyWords.Height := dbMemoAcAppApprovalWords.Height;
  // Set the Top of the Memo fields and the associated Labels
  dbMemoAcAppApprovalWords.Top := lblAcAppSpecAppWords.Top +
    lblAcAppSpecAppWords.Height;
  lblAcAppSpecRejectWords.Top := dbMemoAcAppApprovalWords.Top +
    dbMemoAcAppApprovalWords.Height + memoSpacing;
  dbMemoAcAppRejectWords.Top := lblAcAppSpecRejectWords.Top +
    lblAcAppSpecRejectWords.Height;
  lblAcAppSpecRemedyWords.Top := dbMemoAcAppRejectWords.Top +
    dbMemoAcAppRejectWords.Height + memoSpacing;
  dbMemoAcAppRemedyWords.Top := lblAcAppSpecRemedyWords.Top +
    lblAcAppSpecRemedyWords.Height;
end;

procedure TMainForm.pnlGenVioLetterEnterResize(Sender: TObject);
var
  memoSpacing, margin: integer;
begin
  memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  dbGridGenVioLetters.Height := floor(pnlGenVioLetterEnter.Height * 0.3148148);
  dbGridGenVioLetters.Width := pnlGenVioLetterEnter.Width - 2 * margin;
  // Set the width of the Memo fields
  dbMemoGenVioLettersSpecText.Width := dbGridGenVioLetters.Width;
  dbMemoGenVioLettersVioText.Width := dbGridGenVioLetters.Width;
  dbMemoGenVioLettersRemedyText.Width := dbGridGenVioLetters.Width;
  // Set the Height of the Memo fields
  dbMemoGenVioLettersSpecText.Height :=
    (pnlGenVioLetterEnter.Height - dbGridGenVioLetters.Height -
    dbGridGenVioLetters.Top - 3 * lblGenVioLettersSpecText.Height - 3 *
    memoSpacing - margin) div 3;
  dbMemoGenVioLettersVioText.Height := dbMemoGenVioLettersSpecText.Height;
  dbMemoGenVioLettersRemedyText.Height := dbMemoGenVioLettersSpecText.Height;
  // Set the Top of the Memo fields and the associated Labels
  lblGenVioLettersSpecText.Top := dbGridGenVioLetters.Top +
    dbGridGenVioLetters.Height + memoSpacing;
  dbMemoGenVioLettersSpecText.Top := lblGenVioLettersSpecText.Top +
    lblGenVioLettersSpecText.Height;
  lblGenVioLettersVioText.Top := dbMemoGenVioLettersSpecText.Top +
    dbMemoGenVioLettersSpecText.Height + memoSpacing;
  dbMemoGenVioLettersVioText.Top := lblGenVioLettersVioText.Top +
    lblGenVioLettersVioText.Height;
  lblGenVioLettersRemedyText.Top := dbMemoGenVioLettersVioText.Top +
    dbMemoGenVioLettersVioText.Height + memoSpacing;
  dbMemoGenVioLettersRemedyText.Top := lblGenVioLettersRemedyText.Top +
    lblGenVioLettersRemedyText.Height;
end;

procedure TMainForm.pnlViolationsResize(Sender: TObject);
var
  gutter, margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  gutter := 30;
  margin := 5;
  dbGridViolations.Left := margin;
  dbGridViolations.Height := pnlViolations.Height - dbGridViolations.Top -
    sbViolations.Height - margin;
  dbGridViolations.Width := floor((pnlViolations.Width - 2 * margin - gutter) *
    0.721190);
  // Set Left position of Label and Memo field
  Label7.Left := dbGridViolations.Left + dbGridViolations.Width + gutter;
  dbMemoVioReason.Left := Label7.Left;
  // Set the width of the Memo fields
  // dbMemoVioReason.Top := dbGridViolations.Top;
  dbMemoVioReason.Width := pnlViolations.Width - Label7.Left - margin;
  // Set the Height of the Memo fields
  dbMemoVioReason.Height := dbGridViolations.Height - Label7.Height - margin;
end;

procedure TMainForm.pnlAllGenVioLettersResize(Sender: TObject);
var
  memoSpacing, gutter, margin: integer;
begin
  memoSpacing := 24;
  // bottomSpacing := 10;
  gutter := 30;
  margin := 5;
  // Center the panel Label
  Label17.Left := (pnlAllGenVioLetters.Width - Label7.Width) div 2;
  dbGridBrowseGenVioLetters.Height := pnlAllGenVioLetters.Height -
    dbGridBrowseGenVioLetters.Top - 5;
  dbGridBrowseGenVioLetters.Width :=
    floor(pnlAllGenVioLetters.Width * 0.477870);
  // Set the Label and dbMemo field Left position
  Label14.Left := dbGridBrowseGenVioLetters.Width + margin + gutter;
  Label15.Left := Label14.Left;
  Label16.Left := Label14.Left;
  DBMemo8.Left := Label14.Left;
  DBMemo9.Left := Label14.Left;
  DBMemo10.Left := Label14.Left;
  // Set the dbMemo field Width
  DBMemo8.Width := pnlAllGenVioLetters.Width - DBMemo8.Left - margin;
  DBMemo9.Width := DBMemo8.Width;
  DBMemo10.Width := DBMemo8.Width;
  // Set the dbMemo field Height
  DBMemo8.Height := (dbGridBrowseGenVioLetters.Height -
    (3 * Label14.Height + 2 * memoSpacing)) div 3;
  DBMemo9.Height := DBMemo8.Height;
  DBMemo10.Height := DBMemo8.Height;
  // Set the Top position for the dbMemos and the Labels
  Label16.Top := DBMemo8.Top + DBMemo8.Height + memoSpacing;
  DBMemo9.Top := Label16.Top + Label16.Height;
  Label15.Top := DBMemo9.Top + DBMemo9.Height + memoSpacing;
  DBMemo10.Top := Label15.Top + Label15.Height;
end;

procedure TMainForm.pnlVioStatusResize(Sender: TObject);
var
  memoSpacing, gutter, margin: integer;
begin
  memoSpacing := 24;
  // bottomSpacing := 10;
  gutter := 30;
  margin := 5;
  lblVioStatus.Left := (pnlVioStatus.Width - lblVioStatus.Width) div 3;
  DBNavigator4.Left := dbMemoVioStatStatus.Left +
    (dbMemoVioStatStatus.Width div 2);
  dbGridVioStatus.Left := margin;
  dbGridVioStatus.Height := pnlVioStatus.Height - dbGridVioStatus.Top -
    StatusBar2.Height - margin;
  lblVioStatStatus.Left := dbGridVioStatus.Left + dbGridVioStatus.Width
    + gutter;
  lblVioStatAction.Left := lblVioStatStatus.Left;
  dbMemoVioStatAction.Left := lblVioStatStatus.Left;
  dbMemoVioStatStatus.Left := lblVioStatStatus.Left;
  dbMemoVioStatAction.Width := pnlVioStatus.Width - 2 * margin - gutter -
    dbGridVioStatus.Width;
  dbMemoVioStatStatus.Width := dbMemoVioStatAction.Width;
  dbMemoVioStatStatus.Height :=
    (dbGridVioStatus.Height - 2 * lblVioStatStatus.Height - memoSpacing) div 2;
  dbMemoVioStatAction.Height := dbMemoVioStatStatus.Height;
  dbMemoVioStatStatus.Top := lblVioStatStatus.Top + lblVioStatStatus.Height;
  lblVioStatAction.Top := dbMemoVioStatStatus.Top + dbMemoVioStatStatus.Height +
    memoSpacing;
  dbMemoVioStatAction.Top := lblVioStatAction.Top + lblVioStatAction.Height;

end;

procedure TMainForm.pnlHousesEnterResize(Sender: TObject);
var
  margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  Label18.Left := (pnlHousesEnter.Width - Label18.Height) div 2;
  dbGridHouses.Left := margin;
  dbGridHouses.Height := pnlHousesEnter.Height - dbGridHouses.Top - margin;
  dbGridHouses.Width := pnlHousesEnter.Width - 2 * margin;
end;

procedure TMainForm.pnlOwnersEnterResize(Sender: TObject);
var
  margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  Label19.Left := (pnlOwnersEnter.Width - Label19.Height) div 2;
  DBGrid6.Left := margin;
  DBGrid6.Height := pnlOwnersEnter.Height - DBGrid6.Top -
    sbOwners.Height - margin;
  DBGrid6.Width := pnlOwnersEnter.Width - 2 * margin;
end;

procedure TMainForm.pnlWelcomeEnterResize(Sender: TObject);
var
  margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  Label20.Left := (pnlWelcomeEnter.Width - Label20.Height) div 2;
  DBGrid7.Left := margin;
  DBGrid7.Height := pnlWelcomeEnter.Height - DBGrid7.Top - margin;
  DBGrid7.Width := pnlWelcomeEnter.Width - 2 * margin;
  dbNavWelcome.Left := (pnlWelcomeEnter.Width - dbNavWelcome.Width) div 2;
end;

procedure TMainForm.pnlGenLetters2Resize(Sender: TObject);
var
  margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  RichEdit1.Height := pnlGenLetters2.Height - 2 * margin;
  RichEdit1.Width := pnlGenLetters2.Width - margin - RichEdit1.Left;

end;

procedure TMainForm.pnlGenLetters1Resize(Sender: TObject);
var
  margin: integer;
begin
  // memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  DBGrid8.Width := pnlGenLetters1.Width - 2 * margin;
  DBGrid8.Height := pnlGenLetters1.Height - DBGrid8.Top -
    statBarGenLetters.Height - margin;

end;

procedure TMainForm.TabSheet7Resize(Sender: TObject);
begin
  // do something here
  pnlGenLetters1Resize(Sender);
  pnlGenLetters2Resize(Sender);
end;

procedure TMainForm.pnlAcApps2Resize(Sender: TObject);
var
  memoSpacing, margin: integer;
begin
  memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  // Set Memo field Height
  DBMemo17.Height := (pnlAcApps2.Height - Label25.Top - 3 * Label25.Height - 2 *
    memoSpacing - margin) div 3;
  DBMemo18.Height := DBMemo17.Height;
  DBMemo19.Height := DBMemo17.Height;
  DBMemo17.Width := pnlAcApps2.Width - DBMemo17.Left - margin;
  DBMemo18.Width := DBMemo17.Width;
  DBMemo19.Width := DBMemo18.Width;
  Label26.Top := DBMemo17.Top + DBMemo17.Height + memoSpacing;
  DBMemo18.Top := Label26.Top + Label26.Height;
  Label27.Top := DBMemo18.Top + DBMemo18.Height + memoSpacing;
  DBMemo19.Top := Label27.Top + Label27.Height;
  DBGrid3.Left := margin;
  DBGrid3.Width := pnlAcApps1.Width - DBGrid3.Left - memoSpacing;
  DBNavigator1.Top := pnlAcApps1.Height - DBNavigator1.Height - margin;
  DBNavigator1.Left := (pnlAcApps1.Width - DBNavigator1.Width) div 2;
  DBGrid3.Height := DBNavigator1.Top - memoSpacing - DBGrid3.Top;

end;

{ ----------------------------------------------------------------------+
  AdoDataAllAppLettersAfterInsert:
  This is a procedure to append a new record to the ApprovalLetters
  table. Several fields are populated:

  applicationDate        : Today's date
  specificApprovalWords  : 'Specific Approval Words'
  specificRejectionWords : 'Specific Rejection Words'
  RemedyWords            : 'Remedy Words'
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataAllAppLettersAfterInsert(DataSet: TDataSet);
var
  myHouseAcct: integer;
begin
  myHouseAcct := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  with AdoDataAllAppLetters do
  begin
    FieldValues['houseAcct'] := myHouseAcct;
    FieldValues['applicationDate'] := Date;
    FieldValues['specificApprovalWords'] := 'Specific Approval Words';
    FieldValues['specificRejectionWords'] := 'Specific Rejection Words';
    FieldValues['RemedyWords'] := 'Remedy Words';
  end;
end;

procedure TMainForm.pnlMemoToLegalResize(Sender: TObject);
var
  margin: integer;
begin
//  memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  DBGrid1.Width := TabSheet8.Width - 2 * margin;
  DBGrid1.Height := TabSheet8.Height - DBGrid1.Top - margin;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetOwnersAfterInsert(TDataSet):
  When the user inserts a new Owner record.
  It copies the current houseAcct number and pastes this into the
  new record.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetOwnersAfterInsert(DataSet: TDataSet);
var
  myHouseAcct: integer;
begin
  myHouseAcct := dsCurrentOwners.DataSet.FieldValues['houseAcct'];
  with AdoDataSetOwners do
  begin
    FieldValues['houseAcct'] := myHouseAcct;
  end;
end;

{ ----------------------------------------------------------------------+
  AdoDataSetOwnersAfterPost(TDataSet):
  This procedure toggles the ADODataSets Owners & CurrentOwners
  after the user changes a record in the Owners table. This forces
  an cursor update which then refreshes both the dataset views.When the user inserts a new Owner record.
  It copies the current houseAcct number and pastes this into the
  new record.
  +----------------------------------------------------------------------- }
procedure TMainForm.AdoDataSetOwnersAfterPost(DataSet: TDataSet);
var
  recNumber: integer;
begin
//  acctNumber := AdoDataSetOwners.FieldValues['houseacct'];
  recNumber := AdoTableCurrentOwners.RecNo;
  AdoDataSetOwners.Active := False;
  AdoTableCurrentOwners.Active := False;
  AdoTableCurrentOwners.Active := True;
  AdoDataSetOwners.Active := True;
  AdoTableCurrentOwners.RecNo := recNumber;
  // eCurrentAcctSearch.Text := IntToStr(acctNumber);
  // eCurrentAcctSearch.OnChange(eCurrentAcctSearch);
end;

procedure TMainForm.PageControl1Change(Sender: TObject);
begin
  tsAllOwners.Brush.Color := RGB(200, 140, 230);
end;

procedure TMainForm.PageControl1DrawTab(Control: TCustomTabControl;
  TabIndex: Integer; const Rect: TRect; Active: Boolean);
begin
  DrawTab(Control, TabIndex, Rect, Active);
end;

procedure DrawTab( oControl: TCustomTabControl; TabIndex: Integer;
const Rect: TRect; Active: Boolean);
var
  oRect: TRect;
  sCaption: string;
  iTop: Integer;
  iLeft: Integer;
  cTabColor: TColor;
  iTabTag: Integer;
begin
  oRect := Rect;
  sCaption := TPageControl(oControl).Pages[TabIndex].Caption;
  // Default color for the tab
  cTabColor := clBtnFace;

  // Color the tab the same as the Major panel
  iTabTag := TPageControl(oControl).Pages[TabIndex].Tag;
  case iTabTag of
    1:  cTabColor := $0061ae0c;
    2:  cTabColor := $00f1a65a;
    3:  cTabColor := $0054c622;
    4:  cTabColor := $004998e0;
    5:  cTabColor := clActiveCaption;
    6:  cTabColor := clBtnFace ;
    7:  cTabColor := $00eec29b;
    8:  cTabColor := $00fb5942;
    9:  cTabColor := clBtnFace;
    10: cTabColor := $00e68cc8;
  end;

  iTop  := Rect.Top  + ((Rect.Bottom - Rect.Top - oControl.Canvas.TextHeight(sCaption)) div 2) + 1;
  iLeft := Rect.Left + ((Rect.Right  - Rect.Left - oControl.Canvas.TextWidth(sCaption)) div 2) + 1;

  if (Active) then begin
    oControl.Canvas.Brush.Color := cTabColor;
    oControl.Canvas.FillRect(Rect);
    Frame3D(oControl.Canvas, oRect, clBtnHighlight, clBtnShadow, 2);
  end else begin
    oControl.Canvas.Brush.Color := clBtnFace;
    oControl.Canvas.FillRect(Rect);
  end;
  oControl.Canvas.TextOut(iLeft,iTop,sCaption);
end;

procedure TMainForm.SpeedButton5Click(Sender: TObject);
var
  numOfRecords: integer;
begin
  adoTableOwners.Filtered := False;
  adoTableOwners.IndexFieldNames := 'Owner';
  adoTableOwners.Filtered := True;
  numOfRecords := adoTableOwners.RecordCount;
  StatusBar1.Panels[1].Text := 'Record Count: ' + IntToStr(numOfRecords);
end;

procedure TMainForm.tsAllOwnersResize(Sender: TObject);
var
  margin: integer;
begin
//  memoSpacing := 24;
  // bottomSpacing := 10;
  // gutter := 30;
  margin := 5;
  with DBGrid2 do
  begin
    Top := 9;
    Left := margin;
    Width := pnlAllOwners.Width - 2 * margin;
    Height := pnlAllOwners.Height - StatusBar1.Height - Top - margin;
  end;
end;

procedure TMainForm.UpdateCurrentHouseAcct(newHouseAcct: string);
var
  filterText: string;
begin
  filterText := '';
  if (length(newHouseAcct) > 0) then
    filterText := 'houseAcct = ' + newHouseAcct;
  TEditChange(filterText);
end;

procedure TMainForm.tsAllOwnersEnter(Sender: TObject);
begin
  TabSheet1.Highlighted := True;
  TabSheet1.Font.Style := TabSheet1.Font.Style + [TFontStyle.fsBold];
end;

procedure TMainForm.tsAllOwnersHide(Sender: TObject);
begin
  (Sender as TTabSheet).Highlighted := False;
  (Sender as TTabSheet).Font.Style := (Sender as TTabSheet).Font.Style - [TFontStyle.fsBold];
end;

procedure TMainForm.menuCompactDbClick(Sender: TObject);
begin
  // ShowMessage(OpenDialog.FileName);
  // Close all the database components
  // Tables
  ADOConnection1.Connected := False;
  AdoTableGenVioLetters.Active := False;
  adoTblHouses.Active := False;
  adoTblAllApprovalLetters.Active := False;
  adoTblOffsiteOwners.Active := False;
  adoTblWelcomeLetters.Active := False;
  AdoTableCurrentOwners.Active := False;
  AdoTableViolations.Active := False;
  adoTableAllGenVioLetters.Active := False;
  ADOTable1.Active := False;
  adoTblPropInLegal.Active := False;
  adoTblLegalStatus.Active := False;
  adoTableViolationStatus.Active := False;
  adoTblAllLetters.Active := False;
  adoTblBrowseGenVioLetters.Active := False;
  adoTblMemoToLegal.Active := False;
  adoTableOwners.Active := False;
  // Queries
  ADOQuery1.Active := False;
  adoQryClearLetters.Active := False;
  adoVioQry.Active := False;
  adoQryRunLetters.Active := False;
  adoQryCullLetters.Active := False;
  // Datasets
  AdoDataSetViolations.Active := False;
  AdoDataSetVioStatus.Active := False;
  AdoDataSetVioLetters.Active := False;
  AdoDataSetOwners.Active := False;
  AdoDataAllAppLetters.Active := False;
  CompactDatabase(OpenDialog.FileName, ADOConnection1, True);
  ADOConnection1.Connected := True;

end;

procedure TMainForm.CompactDatabase(const MdbFileName: string;
  ADOConnection: TADOConnection = nil; const ReopenCOnnection: Boolean = True);
var
  LdbFileName, TempFileName: string;
  FailCount: integer;
  FileHandle: integer;
begin
  TempFileName := ChangeFileExt(MdbFileName, '.temp.accdb');
  // ShowMessage('Temp File Name: ' + TempFileName);
  if Assigned(ADOConnection) then
  begin
    // Force the database engine to write data to disk, releasing locks on memory
    JroRefreshCache(ADOConnection);
    // Close the connection - This will also close all associated datasets
    ADOConnection.Close;
  end;
  // ADO Connection.Close SHOULD delete the ldb
  // Force delete of ldb lock file just in case if we don't have an active ADOConnection
  LdbFileName := ChangeFileExt(MdbFileName, '.ldb');
  if FileExists(LdbFileName) then
    DeleteFile(PChar(LdbFileName));
  // Could fail becasue data is still locked - we ignore this
  // Delete temp file if any
  if FileExists(TempFileName) then
    if not DeleteFile(PChar(TempFileName)) then
      RaiseLastOSError;
  // Try to open for exclusive use
  FailCount := 0;
  repeat
    FileHandle := FileOpen(MdbFileName, fmShareExclusive);
    try
      if FileHandle = -1 then // ERROR
      begin
        Inc(FailCount);
        Sleep(100);
        // Give the database engine time to close completely and unlock
      end
      else
      begin
        FailCount := 0;
        Break; // Success
      end;
    finally
      FileClose(FileHandle);
    end;
  until FailCount = 10; // Maximum 1 second of attempts
  if FailCount <> 0 then // File is probably locked by another user/process
    raise Exception.Create(Format('Error opening %s for exclusive use.',
      [MdbFileName]));
  // Compact the db
  JroCompactDatabase(MdbFileName, TempFileName);
  // Copy temp file to original mdb and delete temp file on success
  if Windows.CopyFile(PChar(TempFileName), PChar(MdbFileName), False) then
    DeleteFile(PChar(TempFileName))
  else
    RaiseLastOSError;
  // Reopen ADOConnection
  if Assigned(ADOConnection) and ReopenCOnnection then
    ADOConnection.Open;
end;

procedure TMainForm.JroCompactDatabase(const Source, Destination: string);
var
  JetEngine: OleVariant;
  sString, dString: string;
begin
  sString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + Source +
    '; Jet OLEDB: Engine Type=6;';
  dString := 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + Destination +
    '; Jet OLEDB: Engine Type=6;';
  JetEngine := CreateOleObject('jro.JetEngine');
  // ShowMessage('#1: ' + sString);
  // ShowMessage('#2: ' + dString);
  JetEngine.CompactDatabase(sString, dString);

  // CompactRepair( LogFile:=True, SourceFile:=strSource, DestinationFile:=strDestination)

end;

procedure TMainForm.JroRefreshCache(ADOConnection: TADOConnection);
var
  JetEngine: OleVariant;
begin
  if not ADOConnection.Connected then
    Exit;
  JetEngine := CreateOleObject('jro.JetEngine');
  JetEngine.Refresh(ADOConnection.ConnectionObject);
end;

procedure TMainForm.eCurrentPhoneSearchChange(Sender: TObject);
var
  filterText: string;
begin
  // First we ignore the wildcard as the first character
  if (eCurrentPhoneSearch.Text = '%') then
    Exit;
  filterText := '';
  if (length(eCurrentPhoneSearch.Text) > 0) then
  begin
    filterText := '(phone like ' + '''%' + eCurrentPhoneSearch.Text + '%' +
      ''') OR (altPhone like ' + '''%' + eCurrentPhoneSearch.Text + '%' + ''')';
  end;
  TEditChange(filterText);
end;

{ ----------------------------------------------------------------------+
  DBGrid2CellClick():
  This procedure finds the houseAcct value of the current record
  in the All_Owners grid and copies it to the "Acct Search" TEdit
  box. It then calls the eCurrentAcctSearchChange event.
  +----------------------------------------------------------------------- }
procedure TMainForm.DBGrid2CellClick(Column: TColumn);
begin
  eCurrentAcctSearch.Text := adoTableOwners.FieldValues['houseAcct'];
  eCurrentAcctSearchChange(Column);
end;

procedure TMainForm.Button1Click(Sender: TObject);
begin
  ADOTable2.Active := True;
end;

end.
