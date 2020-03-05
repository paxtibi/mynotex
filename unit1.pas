// ***********************************************************************
// ***********************************************************************
// MyNotex 1.4
// Author and copyright: Massimo Nardello, Modena (Italy) 2010-2017.
// Free software released under GPL licence version 3 or later.

// In this software is used DBZVDateTimePicker component
// (http://wiki.freepascal.org/ZVDateTimeControls_Package#TDBZVDateTimePicker),
// a modified version of RichMemo component
// (http://wiki.lazarus.freepascal.org/RichMemo) and Dcpcrypt
// component (http://wiki.lazarus.freepascal.org/DCPcrypt).
// The source code of the modified version of RichMemo is available in
// the website of MyNotex.

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************

unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Sqlite3DS, DB, FileUtil, LResources, Forms, Controls,
  Graphics, Dialogs, Menus, ComCtrls, ExtCtrls, StdCtrls, Grids, DBGrids,
  DBCtrls, DBZVDateTimePicker, ZVDateTimePicker, RichMemo, WSRichMemo, IpHtml,
  Ipfilebroker, PrintersDlgs, DCPrijndael, DCPsha1, Inifiles, Variants, Zipper,
  LazUTF8, Process, LCLType, Clipbrd, Buttons, FileCtrl, ExtDlgs, types,
  LCLIntf, EditBtn, Printers, Dateutils;

type
  { TfmMain }
  TfmMain = class(TForm)
    apAppProp: TApplicationProperties;
    bnFind2: TBitBtn;
    bnSubDown: TBitBtn;
    bnNotesUp: TBitBtn;
    bnFind: TBitBtn;
    bnFindFirst: TButton;
    bnFindNext: TButton;
    bnNotesDown: TBitBtn;
    bnSubUp: TBitBtn;
    bvBevelFind: TBevel;
    cbFindKind: TComboBox;
    dbDate: TDBZVDateTimePicker;
    dbNavigator: TDBNavigator;
    dbSubComm: TDBMemo;
    dbTags: TDBEdit;
    dbText: TRichMemo;
    dbTitle: TDBEdit;
    dcAES: TDCP_rijndael;
    dsFind: TDatasource;
    dtCalAct: TZVDateTimePicker;
    edLocateText: TEdit;
    edPassword: TEdit;
    fdColorFormatting: TColorDialog;
    fdBackColorFormatting: TColorDialog;
    fdFontGenDialog: TFontDialog;
    grActGrid: TStringGrid;
    grFilter: TDBGrid;
    edFindText: TEdit;
    fdColorDialog: TColorDialog;
    grSubjects: TDBGrid;
    dsNotes: TDatasource;
    dsSubjects: TDatasource;
    fdFontSelDialog: TFontDialog;
    grTitles: TStringGrid;
    ilImageBarActGrid: TImageList;
    ilImageBtActGrid: TImageList;
    ilImageListToolbar: TImageList;
    ilNotes: TImageList;
    imBackground: TImage;
    IpFileDataProvider1: TIpFileDataProvider;
    ipPrint: TIpHtmlPanel;
    lbActDates: TLabel;
    lbAttNames: TListBox;
    lbCalAct: TLabel;
    lbDate: TLabel;
    lbListAttach: TLabel;
    lbListTags: TLabel;
    lbFound: TLabel;
    lbPwdBottom: TLabel;
    lbPwdTop: TLabel;
    lbTags: TLabel;
    lbTagsNames: TListBox;
    lbTitle: TLabel;
    meActNotes: TMemo;
    miFileZim: TMenuItem;
    meSearchCond: TMemo;
    miToolsAlarm: TMenuItem;
    miLines21: TMenuItem;
    miNotesUndo: TMenuItem;
    miLine8a: TMenuItem;
    miLine7a: TMenuItem;
    miLine8b: TMenuItem;
    miLines17b: TMenuItem;
    miNotesShowCal: TMenuItem;
    miFilePrinterSetup: TMenuItem;
    miLine5c: TMenuItem;
    miOrderCustom: TMenuItem;
    miSubjectOrderTitle: TMenuItem;
    miSubjectOrderCustom: TMenuItem;
    miSubjectOrder: TMenuItem;
    miLine7b: TMenuItem;
    pnMinaData: TPanel;
    pnAttList: TPanel;
    pnDetDates: TPanel;
    pnDates: TPanel;
    pnBtnNotes: TPanel;
    pnBtnSubjects: TPanel;
    pnSubNotesGrid: TPanel;
    pmNotesLook: TMenuItem;
    pmNotesLine2: TMenuItem;
    pmSubLook: TMenuItem;
    miNotesLook: TMenuItem;
    miSubjectLook: TMenuItem;
    miToolsEncryptGPG: TMenuItem;
    miLines20: TMenuItem;
    miToolsDecryptGPG: TMenuItem;
    miLine16: TMenuItem;
    miNotesPrint: TMenuItem;
    miLine17: TMenuItem;
    miNotesShowActivities: TMenuItem;
    pmBullLine: TMenuItem;
    pmNoBull: TMenuItem;
    pmBull2: TMenuItem;
    pmBull3: TMenuItem;
    pmBull4: TMenuItem;
    pmBull1: TMenuItem;
    MenuItem1b: TMenuItem;
    pmTextSendAsEmail: TMenuItem;
    pmTextCopyLatex: TMenuItem;
    miLine11b: TMenuItem;
    miNotesSendToBrowser: TMenuItem;
    miLineHelp: TMenuItem;
    miHelp: TMenuItem;
    miNotesimages: TMenuItem;
    opOpenPictureDlg: TOpenPictureDialog;
    pmTextCut: TMenuItem;
    pmTextCopy: TMenuItem;
    pmTextPaste: TMenuItem;
    pmTextLine2: TMenuItem;
    pmOpenBrowser: TMenuItem;
    miTrayExit: TMenuItem;
    miConvertGNote: TMenuItem;
    miConvertTomboy: TMenuItem;
    miLine5b: TMenuItem;
    miFileConvert: TMenuItem;
    miTagsRename: TMenuItem;
    miTagsRemove: TMenuItem;
    miNotesTags: TMenuItem;
    pmChangeFontBackColor: TMenuItem;
    pmBackColorFormatting: TPopupMenu;
    pmRestoreFont: TMenuItem;
    pmHeadingLine: TMenuItem;
    pmHeading3: TMenuItem;
    pmHeading2: TMenuItem;
    pmHeading1: TMenuItem;
    miLines22: TMenuItem;
    miToolsOptions: TMenuItem;
    miLine10: TMenuItem;
    miNotesEncDecrypt: TMenuItem;
    miNotesSendToOO: TMenuItem;
    miNotesSendToLO: TMenuItem;
    miNotesSendToWp: TMenuItem;
    pnActivities: TPanel;
    pnPassword: TPanel;
    pnText: TPanel;
    pmOpenLibreOffice: TMenuItem;
    pmOpenOpenOffice: TMenuItem;
    pmChangeFontKind: TMenuItem;
    miToolsLanguage: TMenuItem;
    pnGridSubjects: TPanel;
    pmTextCopyHtml: TMenuItem;
    pmTextLine1: TMenuItem;
    pmTextSelectAll: TMenuItem;
    pmChangeFontColor: TMenuItem;
    pmSubComments: TMenuItem;
    pmSubLine2: TMenuItem;
    pmAttOpen: TMenuItem;
    pmAttSaveAs: TMenuItem;
    pmAttDelete: TMenuItem;
    miLineAtt1: TMenuItem;
    pmAttNew: TMenuItem;
    miLine14: TMenuItem;
    miAttachSaveAs: TMenuItem;
    miLine13: TMenuItem;
    miAttachDelete: TMenuItem;
    miAttachOpen: TMenuItem;
    miAttachNew: TMenuItem;
    miNotesAttach: TMenuItem;
    miLine12: TMenuItem;
    miNotesInsert: TMenuItem;
    miSubjectComments: TMenuItem;
    miLine7: TMenuItem;
    miLine11: TMenuItem;
    miNotesMove: TMenuItem;
    miFileHTML: TMenuItem;
    miLine19: TMenuItem;
    pmNotesDelete: TMenuItem;
    pmNotes: TPopupMenu;
    pmSubLine1: TMenuItem;
    pmSubDelete: TMenuItem;
    pmNotesLine1: TMenuItem;
    pmSubNew: TMenuItem;
    miLine15: TMenuItem;
    miFileExport: TMenuItem;
    miFileImport: TMenuItem;
    miLine4: TMenuItem;
    miFileClose: TMenuItem;
    miLine5: TMenuItem;
    miLine2: TMenuItem;
    miFileCopyAs: TMenuItem;
    miToolsSyncDo: TMenuItem;
    miLine18: TMenuItem;
    miFileUndo: TMenuItem;
    miNotesShowOnlyText: TMenuItem;
    miOrderByTitle: TMenuItem;
    miOrderByDate: TMenuItem;
    miNotesOrderBy: TMenuItem;
    miLine9: TMenuItem;
    miFileUpdate: TMenuItem;
    miLicence: TMenuItem;
    miFileOpenLast2: TMenuItem;
    miFileOpenLast3: TMenuItem;
    miFileOpenLast4: TMenuItem;
    miFileOpenLast1: TMenuItem;
    miLine6: TMenuItem;
    miFileSave: TMenuItem;
    miLine1: TMenuItem;
    miCopyright: TMenuItem;
    miTools: TMenuItem;
    miToolsCompact: TMenuItem;
    miLine8: TMenuItem;
    miNotesFind: TMenuItem;
    miNotesDelete: TMenuItem;
    miNotesNew: TMenuItem;
    miNotes: TMenuItem;
    miSubjectDelete: TMenuItem;
    miSubjectNew: TMenuItem;
    miSubject: TMenuItem;
    miFileExit: TMenuItem;
    miLine3: TMenuItem;
    miFileOpen: TMenuItem;
    miFileNew: TMenuItem;
    miFile: TMenuItem;
    mmMainMenu: TMainMenu;
    odOpenDialog: TOpenDialog;
    pmNotesNew: TMenuItem;
    pnFindText: TPanel;
    pnAttachments: TPanel;
    pnTextTags: TPanel;
    pnFindLeft: TPanel;
    pnNotesTop: TPanel;
    pnNotes: TPanel;
    pnFind: TPanel;
    pmSubjects: TPopupMenu;
    pmAttachments: TPopupMenu;
    pmColorFormatting: TPopupMenu;
    pmText: TPopupMenu;
    pmFontKind: TPopupMenu;
    pmWordProcessor: TPopupMenu;
    pmHeadings: TPopupMenu;
    pmTray: TPopupMenu;
    pmBulltes: TPopupMenu;
    pdPrintDialog: TPrinterSetupDialog;
    sbStatusBar: TStatusBar;
    sdSaveDialog: TSaveDialog;
    sdSelDirDialog: TSelectDirectoryDialog;
    spActivities: TSplitter;
    spActNotes: TSplitter;
    spSplitterAttach: TSplitter;
    spSplitterSubComm: TSplitter;
    spSplitterFind: TSplitter;
    spSplitterNotes: TSplitter;
    spSplitterAtt: TSplitter;
    spSplitterSubjects: TSplitter;
    sqFind: TSqlite3Dataset;
    sqDelRec: TSqlite3Dataset;
    sqToolsTables: TSqlite3Dataset;
    sqNotes: TSqlite3Dataset;
    sqSubjects: TSqlite3Dataset;
    tbActBar: TToolBar;
    tbActDeleteRow: TToolButton;
    tbActIndLeft: TToolButton;
    tbActIndRight: TToolButton;
    tbActInsertRow: TToolButton;
    tbActMoveDown: TToolButton;
    tbActMoveUp: TToolButton;
    tbActShowWbs: TToolButton;
    tbToolBar: TToolBar;
    tbFileNew: TToolButton;
    tbFileOpen: TToolButton;
    tbFileSave: TToolButton;
    tmAlarm: TTimer;
    tmTimerSave: TTimer;
    ToolButton1: TToolButton;
    ToolButton10: TToolButton;
    tbFontBold: TToolButton;
    tbFontUnderline: TToolButton;
    tbColorFontChange: TToolButton;
    tbAlignRight: TToolButton;
    tbAlignIndent: TToolButton;
    tbCopyHtml: TToolButton;
    tbPaste: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton13: TToolButton;
    tbActCopyAll: TToolButton;
    tbActCopyGroup: TToolButton;
    tbActMoveFromText: TToolButton;
    tbActMoveBack: TToolButton;
    ToolButton14: TToolButton;
    ToolButton15: TToolButton;
    tbActClearAll: TToolButton;
    tbActPasteGroup: TToolButton;
    tbActMoveFor: TToolButton;
    ToolButton2: TToolButton;
    tbFontRestore: TToolButton;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    tbSubjectNew: TToolButton;
    tbNotesNew: TToolButton;
    tbKindFontChange: TToolButton;
    tbOpenNote: TToolButton;
    tbAlignLeft: TToolButton;
    tbBackColorFontChange: TToolButton;
    tbFind: TToolButton;
    ToolButton5: TToolButton;
    ToolButton6: TToolButton;
    ToolButton7: TToolButton;
    tbToolsSyncDo: TToolButton;
    tbFontItalic: TToolButton;
    tbFontStrike: TToolButton;
    tiTrayIcon: TTrayIcon;
    tbAlignCenter: TToolButton;
    tbAlignFill: TToolButton;
    tbCut: TToolButton;
    tbCopy: TToolButton;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    procedure apAppPropException(Sender: TObject; E: Exception);
    procedure bnFind2Click(Sender: TObject);
    procedure bnFindClick(Sender: TObject);
    procedure bnNotesDownClick(Sender: TObject);
    procedure bnNotesUpClick(Sender: TObject);
    procedure bnSubDownClick(Sender: TObject);
    procedure bnSubUpClick(Sender: TObject);
    procedure dbDateCloseUp(Sender: TObject);
    procedure dbDateKeyPress(Sender: TObject; var Key: char);
    procedure dbTagsKeyPress(Sender: TObject; var Key: char);
    procedure dbTextClick(Sender: TObject);
    procedure dbTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure dbTextKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure dbTextMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
    procedure dbTextMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
    procedure dbTitleKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure dbTitleKeyPress(Sender: TObject; var Key: char);
    procedure dsNotesDataChange(Sender: TObject; Field: TField);
    procedure dsNotesStateChange(Sender: TObject);
    procedure dsSubjectsDataChange(Sender: TObject; Field: TField);
    procedure dsSubjectsStateChange(Sender: TObject);
    procedure edFindTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure edLocateTextKeyPress(Sender: TObject; var Key: char);
    procedure edPasswordKeyPress(Sender: TObject; var Key: char);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDragOver(Sender, Source: TObject; X, Y: integer; State: TDragState; var Accept: boolean);
    procedure FormDropFiles(Sender: TObject; const FileNames: array of string);
    procedure FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure FormResize(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormWindowStateChange(Sender: TObject);
    procedure grActGridKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure grActGridPrepareCanvas(Sender: TObject; aCol, aRow: integer; aState: TGridDrawState);
    procedure grSubjectsExit(Sender: TObject);
    procedure grSubjectsKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure grSubjectsPrepareCanvas(Sender: TObject; DataCol: integer; Column: TColumn; AState: TGridDrawState);
    procedure grTitlesDblClick(Sender: TObject);
    procedure grTitlesDrawCell(Sender: TObject; aCol, aRow: integer; aRect: TRect; aState: TGridDrawState);
    procedure grTitlesKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure grTitlesSelectCell(Sender: TObject; aCol, aRow: integer; var CanSelect: boolean);
    procedure lbAttNamesDblClick(Sender: TObject);
    procedure lbAttNamesExit(Sender: TObject);
    procedure lbPwdTopResize(Sender: TObject);
    procedure lbTagsNamesDblClick(Sender: TObject);
    procedure lbTagsNamesExit(Sender: TObject);
    procedure meActNotesChange(Sender: TObject);
    procedure miFileZimClick(Sender: TObject);
    procedure miToolsAlarmClick(Sender: TObject);
    procedure miFilePrinterSetupClick(Sender: TObject);
    procedure miNotesLookClick(Sender: TObject);
    procedure miNotesShowCalClick(Sender: TObject);
    procedure miNotesUndoClick(Sender: TObject);
    procedure miSubjectLookClick(Sender: TObject);
    procedure miSubjectOrderTitleClick(Sender: TObject);
    procedure miToolsEncryptGPGClick(Sender: TObject);
    procedure miToolsDecryptGPGClick(Sender: TObject);
    procedure miAttachDeleteClick(Sender: TObject);
    procedure miAttachNewClick(Sender: TObject);
    procedure miAttachOpenClick(Sender: TObject);
    procedure miConvertTomboyClick(Sender: TObject);
    procedure miFileCloseClick(Sender: TObject);
    procedure miFileCopyAsClick(Sender: TObject);
    procedure miFileExitClick(Sender: TObject);
    procedure miFileImportClick(Sender: TObject);
    procedure miFileNewClick(Sender: TObject);
    procedure miFileOpenClick(Sender: TObject);
    procedure miFileOpenLast1Click(Sender: TObject);
    procedure miFileOpenLast2Click(Sender: TObject);
    procedure miFileOpenLast3Click(Sender: TObject);
    procedure miFileOpenLast4Click(Sender: TObject);
    procedure miFileSaveClick(Sender: TObject);
    procedure miFileUndoClick(Sender: TObject);
    procedure miFileUpdateClick(Sender: TObject);
    procedure miHelpClick(Sender: TObject);
    procedure miLicenceClick(Sender: TObject);
    procedure miNotesEncDecryptClick(Sender: TObject);
    procedure miNotesimagesClick(Sender: TObject);
    procedure miNotesInsertClick(Sender: TObject);
    procedure miNotesMoveClick(Sender: TObject);
    procedure miNotesPrintClick(Sender: TObject);
    procedure miNotesSendToBrowserClick(Sender: TObject);
    procedure miNotesSendToLOClick(Sender: TObject);
    procedure miNotesSendToOOClick(Sender: TObject);
    procedure miNotesDeleteClick(Sender: TObject);
    procedure miNotesFindClick(Sender: TObject);
    procedure miNotesNewClick(Sender: TObject);
    procedure miNotesShowActivitiesClick(Sender: TObject);
    procedure miOrderByDateClick(Sender: TObject);
    procedure miNotesShowOnlyTextClick(Sender: TObject);
    procedure miSubjectCommentsClick(Sender: TObject);
    procedure miSubjectDeleteClick(Sender: TObject);
    procedure miSubjectNewClick(Sender: TObject);
    procedure miTagsRenameClick(Sender: TObject);
    procedure miToolsCompactClick(Sender: TObject);
    procedure miToolsLanguageClick(Sender: TObject);
    procedure miToolsOptionsClick(Sender: TObject);
    procedure miToolsSyncDoClick(Sender: TObject);
    procedure pmBullClick(Sender: TObject);
    procedure pmChangeFontBackColorClick(Sender: TObject);
    procedure pmChangeFontColorClick(Sender: TObject);
    procedure pmChangeFontKindClick(Sender: TObject);
    procedure pmOpenOpenOfficeClick(Sender: TObject);
    procedure pmTextCopyLatexClick(Sender: TObject);
    procedure pmTextSelectAllClick(Sender: TObject);
    procedure pmTextSendAsEmailClick(Sender: TObject);
    procedure sbStatusBarDrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel; const Rect: TRect);
    procedure sqFindAfterScroll(DataSet: TDataSet);
    procedure sqNotesAfterDelete(DataSet: TDataSet);
    procedure sqNotesAfterInsert(DataSet: TDataSet);
    procedure sqNotesAfterPost(DataSet: TDataSet);
    procedure sqNotesAfterScroll(DataSet: TDataSet);
    procedure sqNotesBeforeDelete(DataSet: TDataSet);
    procedure sqNotesBeforeInsert(DataSet: TDataSet);
    procedure sqNotesBeforePost(DataSet: TDataSet);
    procedure sqNotesBeforeScroll(DataSet: TDataSet);
    procedure sqSubjectsAfterDelete(DataSet: TDataSet);
    procedure sqSubjectsAfterInsert(DataSet: TDataSet);
    procedure sqSubjectsAfterPost(DataSet: TDataSet);
    procedure sqSubjectsAfterScroll(DataSet: TDataSet);
    procedure sqSubjectsBeforeDelete(DataSet: TDataSet);
    procedure sqSubjectsBeforeInsert(DataSet: TDataSet);
    procedure sqSubjectsBeforePost(DataSet: TDataSet);
    procedure FindFirstNext(Sender: TObject);
    function IsDirectoryEmpty(const myDir: string): boolean;
    procedure sqSubjectsBeforeScroll(DataSet: TDataSet);
    procedure tbActClearAllClick(Sender: TObject);
    procedure tbActCopyAllClick(Sender: TObject);
    procedure tbActCopyGroupClick(Sender: TObject);
    procedure tbActMoveBackClick(Sender: TObject);
    procedure tbActMoveForClick(Sender: TObject);
    procedure tbActMoveFromTextClick(Sender: TObject);
    procedure tbActPasteGroupClick(Sender: TObject);
    procedure tbColorFontChangeClick(Sender: TObject);
    procedure SetFontFormat(Sender: TObject);
    procedure tbCopyClick(Sender: TObject);
    procedure tbCopyHtmlClick(Sender: TObject);
    procedure tbCutClick(Sender: TObject);
    procedure tbKindFontChangeClick(Sender: TObject);
    procedure tbOpenNoteClick(Sender: TObject);
    procedure tbPasteClick(Sender: TObject);
    procedure tiTrayIconClick(Sender: TObject);
    procedure tmAlarmTimer(Sender: TObject);
    procedure tmTimerSaveTimer(Sender: TObject);
    procedure SetHeadings(Sender: TObject);
    procedure AddImage(FileName: string);
    procedure grActGridDrawCell(Sender: TObject; aCol, aRow: integer; aRect: TRect; aState: TGridDrawState);
    procedure grActGridEditingDone(Sender: TObject);
    procedure grActGridGetEditMask(Sender: TObject; ACol, ARow: integer; var Value: string);
    procedure grActGridKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
    procedure grActGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
    procedure grActGridSelectCell(Sender: TObject; aCol, aRow: integer; var CanSelect: boolean);
    procedure grActGridSelection(Sender: TObject; aCol, aRow: integer);
    procedure grActGridSetEditText(Sender: TObject; ACol, ARow: integer; const Value: string);
    procedure tbActDeleteRowClick(Sender: TObject);
    procedure tbActIndLeftClick(Sender: TObject);
    procedure tbActIndRightClick(Sender: TObject);
    procedure tbActInsertRowClick(Sender: TObject);
    procedure tbActMoveDownClick(Sender: TObject);
    procedure tbActMoveUpClick(Sender: TObject);
    procedure tbActShowWbsClick(Sender: TObject);
    function CheckNumbers(Source: string): boolean;
    procedure CreateTagsList;
    function ConvertDateFormat(Data, FileFormat, SftFormat: string): string;
    procedure SetCharCount(ToHide: boolean);
  private
    function AdjustPosForImages(Pos: integer): integer;
    function AreFontParamsEqual(fp1, fp2: TFontParams): boolean;
    procedure CalcActDates;
    function CheckLink: integer;
    procedure CheckPasswordInput;
    function CleanTagsField(myField: string): string;
    procedure ActivateFormatIcons;
    procedure ClearUndoData;
    procedure ConvertFromTomboyGnote(DataPath: string);
    procedure CopyActivityGroup;
    procedure CopyAllRows;
    procedure CreateAttachment(FileNames: array of string);
    procedure CreateDataTables(DataFileName: string);
    procedure DisableShowTextOnly;
    procedure EditNotesDataset;
    function ExpTextToZim(NoteText: string): string;
    procedure FindData(SearchWithin: boolean);
    procedure LoadActivitiesData;
    function GetBullet: integer;
    function GetTomboyGnoteCreateDate(stText: string): TDate;
    function GetTomboyGnoteLastChangeDate(stText: string): TDateTime;
    function GetTomboyGnoteSubject(stText: string): string;
    function GetTomboyGnoteText(stText: string): string;
    function GetTomboyGnoteTitle(stText: string): string;
    procedure HidePasswordInput;
    procedure LoadTitleDateGrid;
    procedure MakeLink(SelStart: integer);
    procedure LoadRichMemo;
    procedure LocateFromGrid;
    procedure MakeTableUpgrade(DataFileName: string);
    procedure OpenDataTables(DataFileName: string);
    procedure PasteActivityGroup;
    procedure SaveAllData;
    procedure CloseDataTables;
    function SaveRichMemo(FromIdx, ToIdx: integer; SavePictures: boolean): WideString;
    procedure SaveActivitiesData;
    procedure SetLastRowGridActivities;
    procedure GetUndoData;
    procedure ShiftBackwardActDates;
    procedure ShiftForwardActDates;
    procedure ShowPasswordInput;
    function SpaceToLine(stName: string): string;
    procedure StoreUndoData;
    procedure SyncDelRec(ReadFile, WriteFile: string);
    function SyncDelSubjectsNotes(ReadFile, WriteFile: string): integer;
    procedure SyncTitleDateGrid;
    function SyncUpdateSubjectsNotes(ReadFile, WriteFile: string): integer;
    procedure Translation;
    procedure CompactLines(NumRow: integer; Cod: integer);
    procedure DeleteRow(aRow: integer);
    procedure ExpandLines(NumRow: integer; Cod: integer);
    function GetCode(CellVal: string): integer;
    function GetSign(CellVal: string): string;
    procedure IndentToRight(ARow: integer);
    procedure IndentToLeft(aRow: integer);
    procedure InsertRow(aRow: integer);
    procedure MoveDown(aRow: integer);
    procedure MoveUp(aRow: integer);
    procedure UpdateMainAct(IDRow: integer);
    procedure WriteWbs;
    procedure ZoomDecrease;
    procedure ZoomIncrease;
    { private declarations }
  public
    { public declarations }
    // Messages dialogs
    msg001, msg002, msg003, msg004, msg005, msg006, msg007, msg008, msg009, msg010, msg011, msg012, msg013, msg014, msg015, msg016, msg017, msg018, msg019, msg020, msg021, msg022, msg023, msg024, msg025, msg026, msg027, msg028, msg029, msg030, msg031, msg032, msg033, msg034, msg035, msg036, msg037, msg038, msg039, msg040, msg041, msg042, msg043, msg044, msg045, msg046, msg047, msg048, msg049, msg050, msg051, msg052, msg053, msg054, msg055, msg056, msg057, msg058, msg059, msg060, msg061, msg062, msg063, msg064, msg065, msg066, msg067, msg068, msg069, msg070, msg071, msg072, msg073, msg074, msg075, msg076, msg077, msg078, msg079, msg080, msg081, msg082, msg083, msg084,
    // Status bar messages
    sbr001, sbr002, sbr003, sbr004, sbr005, sbr006, sbr007, sbr008, sbr009, sbr010, sbr011,
    // Various captions and labels modified in the code
    cpt001, cpt002, cpt003, cpt004, cpt005, cpt006, cpt007, cpt008, cpt009, cpt010, cpt011, cpt012, cpt013, cpt014, cpt015, cpt016, cpt017, cpt018, cpt019, cpt020, cpt021, cpt022, cpt023, cpt024, cpt025, cpt026, cpt027, cpt028, cpt029, cpt030, cpt031, cpt032, cpt033, cpt034, cpt035, cpt036, cpt037, cpt038, cpt039, cpt040, cpt041, cpt042, cpt043, cpt044, cpt045, cpt046, cpt047, cpt048, cpt049, cpt050, cpt051, cpt052, cpt053, cpt054, cpt055, cpt056, cpt057, cpt058, cpt059, cpt060, cpt061, cpt062, cpt063, cpt064, cpt065, cpt066, cpt067, cpt068, cpt069, cpt070, cpt071, cpt072, cpt073, cpt074, cpt075, cpt076, cpt077, cpt078, cpt079, cpt080, cpt081, cpt082, cpt083, cpt084, cpt085, cpt086, cpt087, cpt088, cpt089, cpt090, cpt091, cpt092, cpt093, cpt094, cpt095, cpt096, cpt097, cpt098, cpt099, cmn001, cmn002, cmn003, cmn004, cmn005, cmn006, cmn007, cmn008, cmn009, cmn010, cmn011, cmn012, cdn001, cdn002, cdn003, cdn004, cdn005, cdn006, cdn007: ShortString;
    // String to check if the password is correct
    PwdCheckString: string;
    // Flag autosync
    flAutosync: boolean;
    // Flag tray icon
    flTrayIcon: boolean;
    // Flag no sync message
    flNoSyncMsg: boolean;
    // Flag no char count
    flNoCharCount: boolean;
    // Flag no autosave
    flNoAutosave: boolean;
    // Transparency
    TranspForm: integer;
    // Folder with other db to sync
    SyncFolder: string;
    // Default font size
    DefFontSize: integer;
    // Default font name
    DefFontName: string;
    // Space between paragraph
    ParagraphSpace: integer;
    // Line within paragraph
    LineSpace: integer;
    // Open last file
    flOpenLastFile: boolean;
    // Activity group
    ActivityGroup: TStringList;
    // Default font and background color
    stColFont1: string;
    stColBack1: string;
    stColFont2: string;
    stColBack2: string;
    stColFont3: string;
    stColback3: string;
    // Date order
    DateOrder: ShortString;
    // 24 hour format
    fl24Hour: boolean;
    // Labels of the activity grid
    lbID: string;
    lbWbs: string;
    lbState: string;
    lbActivity: string;
    lbStartDate: string;
    lbEndDate: string;
    lbDuration: string;
    lbResources: string;
    lbPriority: string;
    lbCompletion: string;
    lbCost: string;
    lbNotes: string;
  end;

var
  fmMain: TfmMain;
  // Version of the software for translation file
  VersMyNt: string = '1.4.1';
  // Last Database used
  LastDatabase1, LastDatabase2, LastDatabase3, LastDatabase4: string;
  // Delay (in days) before purging a record from the deleted record table
  DelayDays: TDate = 180;
  // Filter of sdSaveDialog (useful to save in MyNotex and HTML)
  OpenSaveDlgFilter: string;
  // RichMemo is modified by user
  IsMemoModified: boolean = False;
  // The file in use is modified
  flFileChanged: boolean = False;
  // Tags list must be recrated;
  RecreateTagsList: boolean = False;
  // Saved value of tags (sqNotes does not support OldValue)
  OldTagsValue: string;
  // Text must be loaded
  IsTextToLoad: boolean = True;
  // Confirm for creating a new subject
  ConfirmNewSubject: boolean = True;
  // Windows of the main form maximized or not
  FmMaximized: boolean;
  // Default width of the Search grid visible columns
  WidthSubCol: integer = 300;
  WidthNoteCol: integer = 300;
  WidthDateCol: integer = 150;
  //List of Bookmarks
  BookmarkList: TStringList;
  // Indentation: 50 because the first tab
  // is set to 50 in the RichMemo component
  WidthInden: integer = 50;
  // Left margin of the text set to 15 in the RichMemo component
  DefaultIndent: integer = 15;
  // Zoom font size
  ZoomFontSize: shortint = 0;
  // Installation directory
  // (useful only to let the language file to be installed automatically)
  InstallDir: string = '/opt/mynotex/';
  // Show error messages for debug
  ShowErrorMsg: boolean = False;
  // Activities grid
  LastRowAct: integer = 1; // Last used row
  ColID: integer = 0;
  ColCode: integer = 1;
  ColWbs: integer = 2;
  ColState: integer = 3;
  ColName: integer = 4;
  ColStartDate: integer = 5;
  ColEndDate: integer = 6;
  ColDuration: integer = 7;
  ColResources: integer = 8;
  ColPriority: integer = 9;
  ColCompletion: integer = 10;
  ColCost: integer = 11;
  ColNotes: integer = 12;
  IsGridEditing: boolean = False;
  FDate: TFormatSettings;
  DateMask: ShortString;
  // Note text cursor position, for undo
  iNoteTextPos: integer = 0;
  // Mail for pgp encryption
  stRecipient: string = '';
  // Attachments label
  stAttachments: string = 'Attachments';

const
  // Activities grid costants
  NumRow: integer = 500;
  IndentLev: integer = 18;

implementation

uses Unit2, Unit3, Unit4, Unit5, Unit6, Unit7, Unit8, Unit9, Unit10,
  Unit11, UnitCopyright;

{ TfmMain }

procedure TfmMain.FormCreate(Sender: TObject);
var
  MyIni: TIniFile;
  myHomeDir: string;
  i: integer;
begin
  // Set string to check if password is correct
  PwdCheckString := 'jFd7W0kSmè!=F73Nml<aTegYtDjFXbnGq28' + 'hTdPIj42£àKLèìEjNbfDQlpEWXcv70DnmEkGt8xòùWE23Ghd2HJlmoAQxDjgT67Dq';
  // Set flag autosync
  flAutosync := False;
  // Flag tray icon
  flTrayIcon := False;
  // Flag no sync message
  flNoSyncMsg := False;
  // Flag no char count
  flNoCharCount := False;
  // Flag no autosave
  flNoAutosave := False;
  // Transparency
  TranspForm := 255;
  // Folder with other db to sync
  SyncFolder := '';
  // Default font size
  DefFontSize := 11;
  // Default font name
  DefFontName := 'Sans';
  // Paragraph space
  ParagraphSpace := 0;
  // Line space
  LineSpace := 0;
  // Open last file
  flOpenLastFile := False;
  // Labels of the activity grid
  lbID := 'ID';
  lbWbs := 'WBS';
  lbState := 'State';
  lbActivity := 'Activity';
  lbStartDate := 'Start date';
  lbEndDate := 'End date';
  lbDuration := 'Duration';
  lbResources := 'Resources';
  lbPriority := 'Priority';
  lbCompletion := 'Completion';
  lbCost := 'Cost';
  lbNotes := 'Notes';
  //Initialize clipboard for activities
  ActivityGroup := TStringList.Create;
  // Initialize bookmarks list
  BookmarkList := TStringList.Create;
  for i := 0 to 9 do
    BookmarkList.Add('');
  // Subjects grid
  grSubjects.FocusRectVisible := False;
  grSubjects.FixedColor := clBtnFace;
  // Titles grid
  with grTitles do
  begin
    Clear;
    FocusRectVisible := False;
    for i := 0 to 6 do
      Columns.Add;
    ColWidths[0] := 0;
    Columns.Items[0].Visible := False;
    ColWidths[1] := 250;
    ColWidths[2] := 150;
    ColWidths[3] := 0;
    Columns.Items[3].Visible := False;
    ColWidths[4] := 0;
    Columns.Items[4].Visible := False;
    ColWidths[5] := 0;
    Columns.Items[5].Visible := False;
    ColWidths[6] := 0;
    Columns.Items[6].Visible := False;
    Cells[1, 0] := 'Note title';
    Cells[2, 0] := 'Note date';
  end;
  // Filter grid
  grFilter.FocusRectVisible := False;
  grFilter.FixedColor := clBtnFace;
  // Set activities grid
  with grActGrid do
  begin
    Clear;
    for i := 0 to 11 do
      Columns.Add;
    ColWidths[ColID] := 55;
    ColWidths[ColWbs] := 80;
    Columns.Items[ColWbs - 1].ReadOnly := True;
    ColWidths[ColState] := 100;
    Columns.Items[ColState - 1].ButtonStyle := cbsPickList;
    Columns.Items[ColState - 1].PickList.Add('To do');
    Columns.Items[ColState - 1].PickList.Add('Started');
    Columns.Items[ColState - 1].PickList.Add('Done');
    ColWidths[ColName] := 320;
    ColWidths[ColStartDate] := 180;
    ColWidths[ColEndDate] := 180;
    ColWidths[ColDuration] := 100;
    ColWidths[ColResources] := 250;
    ColWidths[ColPriority] := 100;
    ColWidths[ColCompletion] := 120;
    ColWidths[ColCost] := 100;
    RowCount := NumRow;
    for i := 1 to RowCount - 1 do
      Cells[ColID, i] := IntToStr(i);
    Columns.Items[ColCode - 1].Visible := False;
    Columns.Items[ColWbs - 1].Visible := False;
    Columns.Items[ColNotes - 1].Visible := False;
  end;
  stColFont1 := 'clMaroon';
  stColBack1 := 'clMoneyGreen';
  stColFont2 := 'clMaroon';
  stColBack2 := 'clSkyBlue';
  stColFont3 := 'clNavy';
  stColback3 := 'clSkyBlue';
  cbFindKind.ItemIndex := 0;
  dtCalAct.Date := Date;
  dbSubComm.Color := clForm;
  lbTagsNames.Color := clForm;
  meSearchCond.Color := clForm;
  fl24Hour := False;
  // Set home directory and data directories
  myHomeDir := GetEnvironmentVariable('HOME') + '/.config';
  if DirectoryExists(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator) = False then
  begin
    CreateDirUTF8(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator);
  end;
  // Load main form dimensions from ini file
  if FileExists(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'mynotex') then
  begin
    MyIni := TIniFile.Create(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'mynotex');
    try
      if MyIni.ReadString('mynotex', 'maximize', '') = 'true' then
      begin
        fmMain.WindowState := wsMaximized;
        FmMaximized := True;
      end
      else
      begin
        FmMaximized := False;
        fmMain.Top := MyIni.ReadInteger('mynotex', 'top', 0);
        fmMain.Left := MyIni.ReadInteger('mynotex', 'left', 0);
        if MyIni.ReadInteger('mynotex', 'width', 0) > 100 then
          fmMain.Width := MyIni.ReadInteger('mynotex', 'width', 0)
        else
          fmMain.Width := 800;
        if MyIni.ReadInteger('mynotex', 'heigth', 0) > 100 then
          fmMain.Height := MyIni.ReadInteger('mynotex', 'heigth', 0)
        else
          fmMain.Height := 600;
      end;
      DefFontName := MyIni.ReadString('mynotex', 'fontname', 'Sans');
      DefFontSize := MyIni.ReadInteger('mynotex', 'fontsize', 11);
      fmMain.Color := MyIni.ReadInteger('mynotex', 'formcolor', 0);
      if fmMain.Color <> clDefault then
      begin
        grSubjects.FixedColor := fmMain.Color;
        grTitles.FixedColor := fmMain.Color;
        grActGrid.FixedColor := fmMain.Color;
        grFilter.FixedColor := fmMain.Color;
        dbSubComm.Color := fmMain.Color;
        lbTagsNames.Color := fmMain.Color;
        meSearchCond.Color := fmMain.Color;
      end;
      cbFindKind.ItemIndex := MyIni.ReadInteger('mynotex', 'findkind', 0);
      // The color of the other forms is set on their create event
      pnGridSubjects.Width := MyIni.ReadInteger('mynotex', 'grsubwidth', 0);
      pnSubNotesGrid.Width := MyIni.ReadInteger('mynotex', 'grtitwidth', 0);
      grTitles.ColWidths[1] := MyIni.ReadInteger('mynotex', 'grtitcol0width', 250);
      grTitles.ColWidths[2] := MyIni.ReadInteger('mynotex', 'grtitcol1width', 150);
      pnAttachments.Width := MyIni.ReadInteger('mynotex', 'grattwidth', 200);
      pnAttList.Width := MyIni.ReadInteger('mynotex', 'attlist', 200);
      pnActivities.Height := MyIni.ReadInteger('mynotex', 'actheigth', 0);
      pnFind.Height := MyIni.ReadInteger('mynotex', 'find', 120);
      // Set values of font kind and color buttons
      fdFontSelDialog.Font.Name := MyIni.ReadString('mynotex', 'selfontname', '');
      fdFontSelDialog.Font.Size := MyIni.ReadInteger('mynotex', 'selfontsize', 11);
      if MyIni.ReadInteger('mynotex', 'selfontbold', 0) = 1 then
        fdFontSelDialog.Font.Style := fdFontSelDialog.Font.Style + [fsBold];
      if MyIni.ReadInteger('mynotex', 'selfontitalics', 0) = 1 then
        fdFontSelDialog.Font.Style := fdFontSelDialog.Font.Style + [fsItalic];
      if MyIni.ReadInteger('mynotex', 'selfontunderline', 0) = 1 then
        fdFontSelDialog.Font.Style := fdFontSelDialog.Font.Style + [fsUnderline];
      if MyIni.ReadInteger('mynotex', 'selfontstrike', 0) = 1 then
        fdFontSelDialog.Font.Style := fdFontSelDialog.Font.Style + [fsStrikeOut];
      fdColorFormatting.Color := MyIni.ReadInteger('mynotex', 'selfontcolor', 0);
      fdBackColorFormatting.Color :=
        MyIni.ReadInteger('mynotex', 'selbackfontcolor', 65535); //clYellow
      tbOpenNote.Tag := MyIni.ReadInteger('mynotex', 'wordprocessor', 0);
      WidthSubCol := MyIni.ReadInteger('mynotex', 'grfindcol1width', 300);
      WidthNoteCol := MyIni.ReadInteger('mynotex', 'grfindcol3width', 300);
      WidthDateCol := MyIni.ReadInteger('mynotex', 'grfindcol4width', 150);
      with grActGrid do
      begin
        if MyIni.ReadInteger('mynotex', 'gract1width', 0) > 10 then
          ColWidths[ColWbs] := MyIni.ReadInteger('mynotex', 'gract1width', 80);
        if MyIni.ReadInteger('mynotex', 'gract2width', 0) > 10 then
          ColWidths[ColState] := MyIni.ReadInteger('mynotex', 'gract2width', 100);
        if MyIni.ReadInteger('mynotex', 'gract3width', 0) > 10 then
          ColWidths[ColName] := MyIni.ReadInteger('mynotex', 'gract3width', 320);
        if MyIni.ReadInteger('mynotex', 'gract4width', 0) > 10 then
          ColWidths[ColStartDate] := MyIni.ReadInteger('mynotex', 'gract4width', 180);
        if MyIni.ReadInteger('mynotex', 'gract5width', 0) > 10 then
          ColWidths[ColEndDate] := MyIni.ReadInteger('mynotex', 'gract5width', 180);
        if MyIni.ReadInteger('mynotex', 'gract6width', 0) > 10 then
          ColWidths[ColDuration] := MyIni.ReadInteger('mynotex', 'gract6width', 100);
        if MyIni.ReadInteger('mynotex', 'gract7width', 0) > 10 then
          ColWidths[ColResources] := MyIni.ReadInteger('mynotex', 'gract7width', 250);
        if MyIni.ReadInteger('mynotex', 'gract8width', 0) > 10 then
          ColWidths[ColPriority] := MyIni.ReadInteger('mynotex', 'gract8width', 100);
        if MyIni.ReadInteger('mynotex', 'gract9width', 0) > 10 then
          ColWidths[ColCompletion] := MyIni.ReadInteger('mynotex', 'gract9width', 120);
        if MyIni.ReadInteger('mynotex', 'gract10width', 0) > 10 then
          ColWidths[ColCost] := MyIni.ReadInteger('mynotex', 'gract10width', 100);
      end;
      // Email for pgp encryption
      if MyIni.ReadString('mynotex', 'email', '') <> '' then
        stRecipient := MyIni.ReadString('mynotex', 'email', '');
      // Set titles order
      if MyIni.ReadString('mynotex', 'suborderby', '') = 't' then
        miSubjectOrderTitle.Checked := True
      else if MyIni.ReadString('mynotex', 'suborderby', '') = 'c' then
        miSubjectOrderCustom.Checked := True;
      bnSubUp.Enabled := miSubjectOrderCustom.Checked;
      bnSubDown.Enabled := miSubjectOrderCustom.Checked;
      // Set notes order
      if MyIni.ReadString('mynotex', 'orderby', '') = 't' then
        miOrderByTitle.Checked := True
      else if MyIni.ReadString('mynotex', 'orderby', '') = 'd' then
        miOrderByDate.Checked := True
      else if MyIni.ReadString('mynotex', 'orderby', '') = 'c' then
        miOrderCustom.Checked := True;
      bnNotesUp.Enabled := miOrderCustom.Checked;
      bnNotesDown.Enabled := miOrderCustom.Checked;
      // Tray icon
      if MyIni.ReadString('mynotex', 'trayicon', '') = 'y' then
        flTrayIcon := True;
      if MyIni.ReadString('mynotex', 'nosyncmsg', '') = 'y' then
        flNoSyncMsg := True;
      if MyIni.ReadString('mynotex', 'nocharcount', '') = 'y' then
        flNoCharCount := True;
      if MyIni.ReadString('mynotex', 'noautosave', '') = 'y' then
        flNoAutosave := True;
      // Autosync
      if MyIni.ReadString('mynotex', 'autosync', '') = 'y' then
        flAutosync := True;
      // Show grid activities;
      if MyIni.ReadString('mynotex', 'showact', '') = 'y' then
        miNotesShowActivitiesClick(nil);
      // Menu of opened database
      if MyIni.ReadString('mynotex', 'lastfile1', '') <> '' then
      begin
        LastDatabase1 := MyIni.ReadString('mynotex', 'lastfile1', '');
        miFileOpenLast1.Caption := ExtractFileNameOnly(LastDatabase1);
      end
      else
      begin
        miFileOpenLast1.Visible := False;
      end;
      if MyIni.ReadString('mynotex', 'lastfile2', '') <> '' then
      begin
        LastDatabase2 := MyIni.ReadString('mynotex', 'lastfile2', '');
        miFileOpenLast2.Caption := ExtractFileNameOnly(LastDatabase2);
      end
      else
      begin
        miFileOpenLast2.Visible := False;
      end;
      if MyIni.ReadString('mynotex', 'lastfile3', '') <> '' then
      begin
        LastDatabase3 := MyIni.ReadString('mynotex', 'lastfile3', '');
        miFileOpenLast3.Caption := ExtractFileNameOnly(LastDatabase3);
      end
      else
      begin
        miFileOpenLast3.Visible := False;
      end;
      if MyIni.ReadString('mynotex', 'lastfile4', '') <> '' then
      begin
        LastDatabase4 := MyIni.ReadString('mynotex', 'lastfile4', '');
        miFileOpenLast4.Caption := ExtractFileNameOnly(LastDatabase4);
      end
      else
      begin
        miFileOpenLast4.Visible := False;
      end;
      if ((miFileOpenLast1.Visible = False) and (miFileOpenLast2.Visible = False) and (miFileOpenLast3.Visible = False) and (miFileOpenLast4.Visible = False)) then
      begin
        miLine6.Visible := False;
      end;
      if MyIni.ReadString('mynotex', 'openlast', '') = 'y' then
        flOpenLastFile := True
      else
        flOpenLastFile := False;
      if MyIni.ReadString('mynotex', 'syncfolder', '') <> '' then
      begin
        SyncFolder := MyIni.ReadString('mynotex', 'syncfolder', '');
      end;
      // Form transparency
      if MyIni.ReadInteger('mynotex', 'transparency', 255) > 0 then
        TranspForm := MyIni.ReadInteger('mynotex', 'transparency', 255);
      // Zoom level
      if MyIni.ReadInteger('mynotex', 'zoomfontsize', 0) > 0 then
        ZoomFontSize := MyIni.ReadInteger('mynotex', 'zoomfontsize', 0);
      // Paragraph space
      if MyIni.ReadInteger('mynotex', 'paragraphspace', 0) > 0 then
        ParagraphSpace := MyIni.ReadInteger('mynotex', 'paragraphspace', 0);
      // Line space
      if MyIni.ReadInteger('mynotex', 'linespace', 0) > 0 then
        LineSpace := MyIni.ReadInteger('mynotex', 'linespace', 0);
      // Default colors
      if MyIni.ReadString('mynotex', 'stcolfont1', '') <> '' then
        stColFont1 := MyIni.ReadString('mynotex', 'stcolfont1', '');
      if MyIni.ReadString('mynotex', 'stcolfont2', '') <> '' then
        stColFont2 := MyIni.ReadString('mynotex', 'stcolfont2', '');
      if MyIni.ReadString('mynotex', 'stcolfont3', '') <> '' then
        stColFont3 := MyIni.ReadString('mynotex', 'stcolfont3', '');
      if MyIni.ReadString('mynotex', 'stcolback1', '') <> '' then
        stColBack1 := MyIni.ReadString('mynotex', 'stcolback1', '');
      if MyIni.ReadString('mynotex', 'stcolback2', '') <> '' then
        stColBack2 := MyIni.ReadString('mynotex', 'stcolback2', '');
      if MyIni.ReadString('mynotex', 'stcolback3', '') <> '' then
        stColBack3 := MyIni.ReadString('mynotex', 'stcolback3', '');
      if MyIni.ReadInteger('mynotex', 'firstrun', 0) <> 1 then
      begin
        // A greeting message to the new MyNotex users
        if OpenURL('https://sites.google.com/site/mynotex/firstrun') = True then
        begin
          MyIni.WriteInteger('mynotex', 'firstrun', 1);
        end;
      end;
    finally
      MyIni.Free;
    end;
  end
  else
  begin
    miFileOpenLast1.Visible := False;
    miFileOpenLast2.Visible := False;
    miFileOpenLast3.Visible := False;
    miFileOpenLast4.Visible := False;
    miLine6.Visible := False;
  end;
  // To avoid possibile show of component name in the memo
  dbText.Clear;
  // Property non settable in the inspector
  dbText.WantTabs := True;
  // Dates, might be changed in translation
  FDate := DefaultFormatSettings;
  FDate.ShortDateFormat := 'mm-dd-yyyy';
  FDate.LongDateFormat := 'dddd mmmm d yyyy';
  FDate.DateSeparator := '-';
  dbDate.DateDisplayOrder := ddoMDY;
  dtCalAct.DateDisplayOrder := ddoMDY;
  DateOrder := 'MDY';
  DateMask := '  -  -    ';
  // Load and activate translation
  Translation;
  // Set all elements within opened data
  CloseDataTables;
  // Set the position of the password fields
  pnPassword.Align := alClient;
  // Enable or disable menu items to covert data from Tomboy and GNote
  miConvertTomboy.Enabled :=
    DirectoryExistsUTF8(GetEnvironmentVariable('HOME') + DirectorySeparator + '.local/share/tomboy');
  miConvertGNote.Enabled :=
    DirectoryExistsUTF8(GetEnvironmentVariable('HOME') + DirectorySeparator + '.local/share/gnote');
  // Enable or disable user manual menu
  if FileExists(InstallDir + 'manual-mynotex-en.pdf') = False then
  begin
    miHelp.Visible := False;
    miLineHelp.Visible := False;
  end;
end;

procedure TfmMain.FormDragOver(Sender, Source: TObject; X, Y: integer; State: TDragState; var Accept: boolean);
begin
  // Check if a file can be dropped
  // This function actually does not work; maybe in the future...
  Accept := miAttachNew.Enabled = True;
end;

procedure TfmMain.FormDropFiles(Sender: TObject; const FileNames: array of string);
begin
  // Attach dropped files
  if miAttachNew.Enabled = True then
    CreateAttachment(FileNames);
end;

procedure TfmMain.FormShow(Sender: TObject);
begin
  // Open files with double clic or by console (-l to open the last archive)
  if sqSubjects.Active = False then
  begin
    if Application.Params[1] <> '' then
    begin
      if Application.Params[1] = '-l' then
      begin
        if LastDatabase1 <> '' then
          OpenDataTables(LastDatabase1);
      end
      else if Application.Params[1] = '-debug' then
        ShowErrorMsg := True
      else
        OpenDataTables(Application.Params[1]);
    end
    else if flOpenLastFile = True then
    begin
      if LastDatabase1 <> '' then
        OpenDataTables(LastDatabase1);
    end;
    // Set Transparency
    if TranspForm < 255 then
    begin
      fmMain.AlphaBlendValue := TranspForm;
      fmImpExp.AlphaBlendValue := TranspForm;
      fmMoveNote.AlphaBlendValue := TranspForm;
      fmCommentsSubjects.AlphaBlendValue := TranspForm;
      fmEncryption.AlphaBlendValue := TranspForm;
      fmResizeImage.AlphaBlendValue := TranspForm;
      fmLook.AlphaBlendValue := TranspForm;
      fmCalendar.AlphaBlendValue := TranspForm;
      fmSetAlarm.AlphaBlendValue := TranspForm;
      fmOptions.AlphaBlendValue := TranspForm;
    end;
  end;
  // Set tray icon
  if flTrayIcon = True then
  begin
    fmMain.tiTrayIcon.Visible := True;
    fmMain.ShowInTaskBar := stNever;
  end;
end;

procedure TfmMain.FormClose(Sender: TObject; var CloseAction: TCloseAction);
var
  MyIni: TIniFile;
  myHomeDir: string;
begin
  // Save all data
  SaveAllData;
  // Close Tables and save ID Subject and IDNote
  CloseDataTables;
  // Sometimes exiting with data to be saved it raises an error;
  // following line seems to prevents it.
  Application.ProcessMessages;
  // Destroy bookmarks list
  BookmarkList.Free;
  // Set home directory and data directories
  myHomeDir := GetEnvironmentVariable('HOME') + '/.config';
  // Save main form dimensions and other elements to ini file
  try
    MyIni := TIniFile.Create(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'mynotex');
    if fmMain.WindowState = wsMaximized then
    begin
      MyIni.WriteString('mynotex', 'maximize', 'true');
    end
    else
    begin
      MyIni.WriteString('mynotex', 'maximize', 'false');
      MyIni.WriteInteger('mynotex', 'top', fmMain.Top);
      MyIni.WriteInteger('mynotex', 'left', fmMain.Left);
      MyIni.WriteInteger('mynotex', 'width', fmMain.Width);
      MyIni.WriteInteger('mynotex', 'heigth', fmMain.Height);
    end;
    MyIni.WriteString('mynotex', 'fontname', DefFontName);
    MyIni.WriteInteger('mynotex', 'fontsize', DefFontSize);
    MyIni.WriteInteger('mynotex', 'formcolor', fmMain.Color);
    MyIni.WriteInteger('mynotex', 'findkind', cbFindKind.ItemIndex);
    // Set values of font kind and color buttons
    MyIni.WriteString('mynotex', 'selfontname', fdFontSelDialog.Font.Name);
    if fdFontSelDialog.Font.Size > 5 then
      MyIni.WriteInteger('mynotex', 'selfontsize', fdFontSelDialog.Font.Size)
    else
      MyIni.WriteInteger('mynotex', 'selfontsize', 11);
    MyIni.WriteInteger('mynotex', 'selfontcolor', fdColorFormatting.Color);
    MyIni.WriteInteger('mynotex', 'selbackfontcolor', fdBackColorFormatting.Color);
    MyIni.WriteInteger('mynotex', 'wordprocessor', tbOpenNote.Tag);
    // Set various width
    if pnGridSubjects.Width < 10 then
    begin
      MyIni.WriteInteger('mynotex', 'grsubwidth', 50);
    end
    else
    begin
      MyIni.WriteInteger('mynotex', 'grsubwidth', pnGridSubjects.Width);
    end;
    if pnSubNotesGrid.Width < 10 then
    begin
      MyIni.WriteInteger('mynotex', 'grtitwidth', 50);
    end
    else
    begin
      MyIni.WriteInteger('mynotex', 'grtitwidth', pnSubNotesGrid.Width);
    end;
    if grTitles.ColWidths[1] < 10 then
      MyIni.WriteInteger('mynotex', 'grtitcol0width', 50)
    else
      MyIni.WriteInteger('mynotex', 'grtitcol0width', grTitles.ColWidths[1]);
    if grTitles.ColWidths[2] < 10 then
      MyIni.WriteInteger('mynotex', 'grtitcol1width', 50)
    else
      MyIni.WriteInteger('mynotex', 'grtitcol1width', grTitles.ColWidths[2]);
    if pnAttachments.Width < 0 then
    begin
      MyIni.WriteInteger('mynotex', 'grattwidth', 200);
    end
    else
    begin
      MyIni.WriteInteger('mynotex', 'grattwidth', pnAttachments.Width);
    end;
    if pnAttList.Width < 100 then
    begin
      MyIni.WriteInteger('mynotex', 'attlist', 200);
    end
    else
    begin
      MyIni.WriteInteger('mynotex', 'attlist', pnAttList.Width);
    end;
    if pnActivities.Height < 100 then
    begin
      MyIni.WriteInteger('mynotex', 'actheigth', 100);
    end
    else
    begin
      MyIni.WriteInteger('mynotex', 'actheigth', pnActivities.Height);
    end;
    if pnFind.Height < 120 then
    begin
      MyIni.WriteInteger('mynotex', 'find', 120);
    end
    else
    begin
      MyIni.WriteInteger('mynotex', 'find', pnFind.Height);
    end;
    // Save the width of the Search grid
    if WidthSubCol < 10 then
      MyIni.WriteInteger('mynotex', 'grfindcol1width', 10)
    else
      MyIni.WriteInteger('mynotex', 'grfindcol1width', WidthSubCol);
    if WidthNoteCol < 10 then
      MyIni.WriteInteger('mynotex', 'grfindcol3width', 10)
    else
      MyIni.WriteInteger('mynotex', 'grfindcol3width', WidthNoteCol);
    if WidthDateCol < 10 then
      MyIni.WriteInteger('mynotex', 'grfindcol4width', 10)
    else
      MyIni.WriteInteger('mynotex', 'grfindcol4width', WidthDateCol);
    // Save activity grid
    with grActGrid do
    begin
      if ColWidths[ColWbs] < 10 then
        MyIni.WriteInteger('mynotex', 'gract1width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract1width', ColWidths[ColWbs]);
      if ColWidths[ColState] < 10 then
        MyIni.WriteInteger('mynotex', 'gract2width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract2width', ColWidths[ColState]);
      if ColWidths[ColName] < 10 then
        MyIni.WriteInteger('mynotex', 'gract3width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract3width', ColWidths[ColName]);
      if ColWidths[ColStartDate] < 10 then
        MyIni.WriteInteger('mynotex', 'gract4width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract4width', ColWidths[ColStartDate]);
      if ColWidths[ColEndDate] < 10 then
        MyIni.WriteInteger('mynotex', 'gract5width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract5width', ColWidths[ColEndDate]);
      if ColWidths[ColDuration] < 10 then
        MyIni.WriteInteger('mynotex', 'gract6width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract6width', ColWidths[ColDuration]);
      if ColWidths[ColResources] < 10 then
        MyIni.WriteInteger('mynotex', 'gract7width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract7width', ColWidths[ColResources]);
      if ColWidths[ColPriority] < 10 then
        MyIni.WriteInteger('mynotex', 'gract8width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract8width', ColWidths[ColPriority]);
      if ColWidths[ColCompletion] < 10 then
        MyIni.WriteInteger('mynotex', 'gract9width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract9width', ColWidths[ColCompletion]);
      if ColWidths[ColCost] < 10 then
        MyIni.WriteInteger('mynotex', 'gract10width', 10)
      else
        MyIni.WriteInteger('mynotex', 'gract10width', ColWidths[ColCost]);
    end;
    // Email for pgp encryption
    if stRecipient <> '' then
      MyIni.WriteString('mynotex', 'email', stRecipient);
    // Save titles order
    if miSubjectOrderTitle.Checked = True then
      MyIni.WriteString('mynotex', 'suborderby', 't')
    else if miSubjectOrderCustom.Checked = True then
      MyIni.WriteString('mynotex', 'suborderby', 'c');
    // Save notes order
    if miOrderByTitle.Checked = True then
      MyIni.WriteString('mynotex', 'orderby', 't')
    else if miOrderByDate.Checked = True then
      MyIni.WriteString('mynotex', 'orderby', 'd')
    else if miOrderCustom.Checked = True then
      MyIni.WriteString('mynotex', 'orderby', 'c');
    // Tray icon
    if flTrayIcon = True then
      MyIni.WriteString('mynotex', 'trayicon', 'y')
    else
      MyIni.WriteString('mynotex', 'trayicon', 'n');
    // No sync message
    if flNoSyncMsg = True then
      MyIni.WriteString('mynotex', 'nosyncmsg', 'y')
    else
      MyIni.WriteString('mynotex', 'nosyncmsg', 'n');
    // No char count
    if flNoCharCount = True then
      MyIni.WriteString('mynotex', 'nocharcount', 'y')
    else
      MyIni.WriteString('mynotex', 'nocharcount', 'n');
    // No autosave
    if flNoAutosave = True then
      MyIni.WriteString('mynotex', 'noautosave', 'y')
    else
      MyIni.WriteString('mynotex', 'noautosave', 'n');
    // Autosync
    if flAutosync = True then
      MyIni.WriteString('mynotex', 'autosync', 'y')
    else
      MyIni.WriteString('mynotex', 'autosync', 'n');
    // Show grid activities
    if miNotesShowActivities.Checked = True then
      MyIni.WriteString('mynotex', 'showact', 'y')
    else
      MyIni.WriteString('mynotex', 'showact', 'n');
    if LastDatabase1 <> '' then
    begin
      MyIni.WriteString('mynotex', 'lastfile1', LastDatabase1);
    end;
    if LastDatabase2 <> '' then
    begin
      MyIni.WriteString('mynotex', 'lastfile2', LastDatabase2);
    end;
    if LastDatabase3 <> '' then
    begin
      MyIni.WriteString('mynotex', 'lastfile3', LastDatabase3);
    end;
    if LastDatabase4 <> '' then
    begin
      MyIni.WriteString('mynotex', 'lastfile4', LastDatabase4);
    end;
    if flOpenLastFile = True then
      MyIni.WriteString('mynotex', 'openlast', 'y')
    else
      MyIni.WriteString('mynotex', 'openlast', 'n');
    if SyncFolder <> '' then
    begin
      MyIni.WriteString('mynotex', 'syncfolder', SyncFolder);
    end;
    // Form transparency
    MyIni.WriteInteger('mynotex', 'transparency', TranspForm);
    // Zoom level
    MyIni.WriteInteger('mynotex', 'zoomfontsize', ZoomFontSize);
    // Space between paragraph
    MyIni.WriteInteger('mynotex', 'paragraphspace', ParagraphSpace);
    // Line within paragraph
    MyIni.WriteInteger('mynotex', 'linespace', LineSpace);
    // Default colors
    MyIni.WriteString('mynotex', 'stcolfont1', stColFont1);
    MyIni.WriteString('mynotex', 'stcolfont2', stColFont2);
    MyIni.WriteString('mynotex', 'stcolfont3', stColFont3);
    MyIni.WriteString('mynotex', 'stcolback1', stColBack1);
    MyIni.WriteString('mynotex', 'stcolback2', stColBack2);
    MyIni.WriteString('mynotex', 'stcolback3', stColBack1);
    // Delete possibile file MyNotexFile.html in tmp
    if FileExistsUTF8(GetTempDir + DirectorySeparator + 'MyNotexFile.html') then
    begin
      DeleteFileUTF8(GetTempDir + DirectorySeparator + 'MyNotexFile.html');
    end;
  finally
    MyIni.Free;
  end;
end;

procedure TfmMain.FormResize(Sender: TObject);
begin
  // Resize the status bar
  sbStatusBar.Panels[0].Width := fmMain.Width - 450;
  sbStatusBar.Panels[1].Width := 300;
  sbStatusBar.Panels[2].Width := 150;
end;

procedure TfmMain.FormWindowStateChange(Sender: TObject);
begin
  // Set variable for tiTrayIconClick
  if fmMain.WindowState = wsMaximized then
    FmMaximized := True
  else if fmMain.WindowState = wsNormal then
    FmMaximized := False;
end;

procedure TfmMain.FormKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Manage menu shortcuts, non working from every fields
  // Open File
  if ((Key = Ord('O')) and (Shift = [ssCtrl])) then
  begin
    if miFileOpen.Enabled = True then
    begin
      miFileOpenClick(nil);
      Key := 0;
    end;
  end;
  // Exit
  if ((Key = Ord('Q')) and (Shift = [ssCtrl])) then
  begin
    if miFileExit.Enabled = True then
    begin
      miFileExitClick(nil);
      Key := 0;
    end;
  end;
  // New Note
  if ((Key = Ord('N')) and (Shift = [ssCtrl])) then
  begin
    if miNotesNew.Enabled = True then
    begin
      miNotesNewClick(nil);
      Key := 0;
    end;
  end;
  // Save Note
  if ((Key = Ord('S')) and (Shift = [ssCtrl])) then
  begin
    if miFileSave.Enabled = True then
    begin
      miFileSaveClick(nil);
      Key := 0;
    end;
  end;
  // Undo editing for file
  if ((Key = Ord('Z')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if miFileUndo.Enabled = True then
    begin
      miFileUndoClick(nil);
      Key := 0;
    end;
  end;
  // Undo editing for note
  if ((Key = Ord('Z')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if miNotesUndo.Enabled = True then
    begin
      miNotesUndoClick(nil);
      Key := 0;
    end;
  end;
  // Show only text memo
  if ((Key = Ord('H')) and (Shift = [ssCtrl])) then
  begin
    if miNotesShowOnlyText.Enabled = True then
    begin
      miNotesShowOnlyTextClick(nil);
      Key := 0;
    end;
  end;
  // Order Note by date
  if ((Key = Ord('D')) and (Shift = [ssCtrl])) then
  begin
    if miOrderByDate.Enabled = True then
    begin
      miOrderByDate.Checked := True;
      miOrderByDateClick(nil);
      Key := 0;
    end;
  end;
  // Order Note by title
  if ((Key = Ord('T')) and (Shift = [ssCtrl])) then
  begin
    if miOrderByTitle.Enabled = True then
    begin
      miOrderByTitle.Checked := True;
      // The following event is called also by miOrderByTitle
      miOrderByDateClick(nil);
      Key := 0;
    end;
  end;
  // Order Note by custom
  if ((Key = Ord('U')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if miOrderCustom.Enabled = True then
    begin
      miOrderCustom.Checked := True;
      // The following event is called also by miOrderByTitle
      miOrderByDateClick(nil);
      Key := 0;
    end;
  end;
  // Encrypt Note
  if ((Key = Ord('Y')) and (Shift = [ssCtrl])) then
  begin
    if miNotesEncDecrypt.Enabled = True then
    begin
      miNotesEncDecryptClick(nil);
      Key := 0;
    end;
  end;
  // Insert text in a new note
  if ((Key = Ord('I')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if miNotesInsert.Enabled = True then
    begin
      miNotesInsertClick(nil);
      Key := 0;
    end;
  end;
  // Print
  if ((Key = Ord('P')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if miNotesPrint.Enabled = True then
    begin
      miNotesPrintClick(nil);
      Key := 0;
    end;
  end;
  // Find Note
  if ((Key = Ord('F')) and (Shift = [ssCtrl])) then
  begin
    if miNotesFind.Enabled = True then
    begin
      miNotesFindClick(nil);
      Key := 0;
    end;
  end;
  // Sync
  if ((Key = Ord('K')) and (Shift = [ssCtrl])) then
  begin
    if miToolsSyncDo.Enabled = True then
    begin
      miToolsSyncDoClick(nil);
      Key := 0;
    end;
  end;
  // Send to email
  if ((Key = Ord('E')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if pmTextSendAsEmail.Enabled = True then
    begin
      pmTextSendAsEmailClick(nil);
      Key := 0;
    end;
  end;
  // PgDn - Prev Record
  if ((Key = 33) and (Shift = [ssCtrl])) then
  begin
    if sqNotes.Active = True then
    begin
      if ((grSubjects.Focused = False) and (grTitles.Focused = False) and (miNotesShowOnlyText.Checked = False)) then
      begin
        sqNotes.Prior;
        Key := 0;
      end;
    end;
  end;
  // PgDn - Next Record
  if ((Key = 34) and (Shift = [ssCtrl])) then
  begin
    if sqNotes.Active = True then
    begin
      if ((grSubjects.Focused = False) and (grTitles.Focused = False) and (miNotesShowOnlyText.Checked = False)) then
      begin
        sqNotes.Next;
        Key := 0;
      end;
    end;
  end;
  // Set bookmark (0 Key=48; 9 Key=57)
  if ((Key >= Ord('0')) and (Key <= Ord('9')) and (Shift = [ssShift, ssCtrl])) then
  begin
    if ((sqSubjects.RecordCount > 0) and (sqNotes.RecordCount > 0)) then
    begin
      BookmarkList[Key - 48] :=
        sqSubjects.FieldByName('IDSubjects').AsString + '|' + sqNotes.FieldByName('IDNotes').AsString;
      sbStatusBar.Panels[0].Text := ' ' + sbr011 + ' ' + IntToStr(Key - 48);
    end;
  end
  // Go to bookmark
  else if ((Key >= Ord('0')) and (Key <= Ord('9')) and (Shift = [ssCtrl, ssAlt])) then
  begin
    IsTextToLoad := False;
    sqSubjects.Locate('IDSubjects', Copy(BookmarkList[Key - 48], 1, Pos('|', BookmarkList[Key - 48]) - 1), []);
    sqNotes.Locate('IDNotes', Copy(BookmarkList[Key - 48], Pos('|', BookmarkList[Key - 48]) + 1, Length(BookmarkList[Key - 48])), []);
    IsTextToLoad := True;
    sqNotesAfterScroll(nil);
  end
  // Copy note link
  else if ((Key = Ord('N')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if sqNotes.RecordCount > 0 then
    begin
      Clipboard.AsText := 'mnt://' + sqSubjects.FieldByName('SubjectsName').AsString + '/' + sqNotes.FieldByName('NotesTitle').AsString;
      Clipboard.AsText := StringReplace(Clipboard.AsText, ' ', '_', [rfReplaceAll]);
      Key := 0;
    end;
  end
  else if ((Key = Ord('B')) and (Shift = [ssCtrl, ssShift])) then
  begin
    tbToolBar.Visible := not tbToolBar.Visible;
    Key := 0;
  end;
end;

procedure TfmMain.dbTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
var
  fp: TFontParams;
  iChar, idxStart, idxLength: integer;
  NewItemBullet: shortint = 0;
begin
  // Manage the bullet/number in the next paragraph
  if ((Key = 13) and (Shift = [ssCtrl])) then
  begin
    dbText.GetTextAttributes(dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True), fp);
    if fp.Indented > DefaultIndent then
    begin
      // Indentation may remain after delete, so...
      idxStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True);
      idxLength := dbText.GetWordParagraphStartEnd(dbText.SelStart, False, False) - idxStart;
      fp.Changed := [fiIndented];
      fp.Indented := DefaultIndent;
      dbText.SetTextAttributes(idxStart - 1, idxLength, fp);
      dbText.Lines.Delete(dbText.CaretPos.Y);
      EditNotesDataset;
    end;
  end
  else if Key = 13 then
  begin
    dbText.GetTextAttributes(dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True), fp);
    if fp.Indented > DefaultIndent then
    begin
      NewItemBullet := GetBullet;
      if dbText.CaretPos.X > 0 then
        dbText.InsertText(dbText.SelStart, LineEnding + ' ' + #9)
      else
        dbText.InsertText(dbText.SelStart, LineEnding);
      if NewItemBullet = 1 then
        dbText.ListNumber(dbText.SelStart, WidthInden, '*')
      else if NewItemBullet = 2 then
        dbText.ListNumber(dbText.SelStart, WidthInden, '1')
      else if NewItemBullet = 3 then
        dbText.ListNumber(dbText.SelStart, WidthInden, 'A')
      else if NewItemBullet = 4 then
        dbText.ListNumber(dbText.SelStart, WidthInden, 'a');
      Key := 0;
    end;
  end;
  // If sqSubjects = 0 is still possibile to write a letter in the memo,
  // also if it is cancelled when it lost the focus; to prevent this...
  if sqSubjects.RecordCount = 0 then
  begin
    key := 0;
  end;
  // Put in edit the dataset if the key is not a shortcut
  // 16 Shift - 18 Left Alt - 19 Break - 20 Caps lock - 27 Esc - 33 PgUp
  // 34 PgDn - 35 End - 36 Home - 37 Left arrow - 38 Up arrow
  // 39 Right arrow - 40 Down arrow - 45 Ins - 91 Left Win - 92 Right Win
  // 93 pop menu - 112-123 F1-F12 - 144 Lock num - 145 Lock scroll - 235 Alt gr
  // 48-57+Ctrl bookmarks - 17 Ctrl - 13 Return - 8 backspace
  if ((((Shift <> [ssCtrl]) and (Shift <> [ssCtrl, ssShift]) and (Shift <> [ssCtrl, ssAlt]) and (Shift <> [ssCtrl, ssShift, ssAlt])) and not (key in [16, 18, 19, 20, 27, 33, 34, 35, 36, 37, 38, 39, 40, 45, 91, 92, 93, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 144, 145, 235])) or ((Shift = [ssCtrl]) and (Key = Ord('X'))) or ((Shift = [ssCtrl]) and (Key = Ord('V'))) or ((Shift = [ssCtrl, ssShift]) and (Key = Ord('Y'))) or (Key = 8)) then
  begin
    EditNotesDataset;
    // Store undo on space without selection active
    if ((Key = Ord(' ')) and (dbText.SelLength < 1)) then
    begin
      StoreUndoData;
    end;
    // Format possible internet link
    if ((Key = Ord(' ')) or (Key = 13)) then
    begin
      MakeLink(CheckLink);
    end;
  end;
  // Ctrl + Shift + 0-9 are use to set bookmarks
  if ((Shift = [ssCtrl, ssShift]) and (Key >= Ord('0')) and (Key <= Ord('9'))) then
    Key := 0;
  // Ctrl + Alt + 0-9 are use to get bookmarks
  if ((Shift = [ssCtrl, ssAlt]) and (Key >= Ord('0')) and (Key <= Ord('9'))) then
    Key := 0;
  // Shortcut for bold, italic, underline and restore
  if ((Key = Ord('B')) and (Shift = [ssCtrl])) then
  begin
    if tbFontBold.Enabled = True then
    begin
      tbFontBold.Down := not tbFontBold.Down;
      SetFontFormat(tbFontBold);
      Key := 0;
    end;
  end;
  if ((Key = Ord('I')) and (Shift = [ssCtrl])) then
  begin
    if tbFontItalic.Enabled = True then
    begin
      tbFontItalic.Down := not tbFontItalic.Down;
      SetFontFormat(tbFontItalic);
      Key := 0;
    end;
  end;
  if ((Key = Ord('U')) and (Shift = [ssCtrl])) then
  begin
    if tbFontUnderline.Enabled = True then
    begin
      tbFontUnderline.Down := not tbFontUnderline.Down;
      SetFontFormat(tbFontUnderline);
      Key := 0;
    end;
  end;
  if ((Key = Ord('M')) and (Shift = [ssCtrl])) then
  begin
    if tbFontRestore.Enabled = True then
    begin
      SetFontFormat(tbFontRestore);
      Key := 0;
    end;
  end;
  // Shortcut for alignment
  if ((Key = Ord('L')) and (Shift = [ssCtrl])) then
  begin
    if tbAlignLeft.Enabled = True then
    begin
      tbAlignLeft.Down := True;
      tbAlignCenter.Down := False;
      tbAlignRight.Down := False;
      tbAlignFill.Down := False;
      SetFontFormat(tbAlignLeft);
      Key := 0;
    end;
  end;
  if ((Key = Ord('E')) and (Shift = [ssCtrl])) then
  begin
    if tbAlignCenter.Enabled = True then
    begin
      tbAlignCenter.Down := True;
      tbAlignLeft.Down := False;
      tbAlignRight.Down := False;
      tbAlignFill.Down := False;
      SetFontFormat(tbAlignCenter);
      Key := 0;
    end;
  end;
  if ((Key = Ord('R')) and (Shift = [ssCtrl])) then
  begin
    if tbAlignRight.Enabled = True then
    begin
      tbAlignRight.Down := True;
      tbAlignLeft.Down := False;
      tbAlignCenter.Down := False;
      tbAlignFill.Down := False;
      SetFontFormat(tbAlignRight);
      Key := 0;
    end;
  end;
  if ((Key = Ord('J')) and (Shift = [ssCtrl])) then
  begin
    if tbAlignFill.Enabled = True then
    begin
      tbAlignFill.Down := True;
      tbAlignLeft.Down := False;
      tbAlignCenter.Down := False;
      tbAlignRight.Down := False;
      SetFontFormat(tbAlignFill);
      Key := 0;
    end;
  end;
  // Shortcut for color background
  if ((Key = Ord('P')) and (Shift = [ssCtrl])) then
  begin
    if tbBackColorFontChange.Enabled = True then
    begin
      tbColorFontChangeClick(tbBackColorFontChange);
      Key := 0;
    end;
  end;
  // Shortcut for copy as html
  if ((Key = Ord('C')) and (Shift = [ssCtrl, ssShift])) then
  begin
    if tbCopyHtml.Enabled = True then
    begin
      tbCopyHtmlClick(nil);
      Key := 0;
    end;
  end;
  // Shortcut for copy as Latex
  if ((Key = Ord('C')) and (Shift = [ssCtrl, ssShift, ssAlt])) then
  begin
    if pmTextCopyLatex.Enabled = True then
    begin
      pmTextCopyLatexClick(nil);
      Key := 0;
    end;
  end;
  // Shortcut for Heading 1
  if ((Key = Ord('1')) and (Shift = [ssCtrl])) then
  begin
    if pmHeading1.Enabled = True then
    begin
      SetHeadings(pmHeading1);
      Key := 0;
    end;
  end;
  // Shortcut for Heading 2
  if ((Key = Ord('2')) and (Shift = [ssCtrl])) then
  begin
    if pmHeading2.Enabled = True then
    begin
      SetHeadings(pmHeading2);
      Key := 0;
    end;
  end;
  // Shortcut for Heading 3
  if ((Key = Ord('3')) and (Shift = [ssCtrl])) then
  begin
    if pmHeading3.Enabled = True then
    begin
      SetHeadings(pmHeading3);
      Key := 0;
    end;
  end;
  // Shortcut for restore default font on all the paragraph
  if ((Key = Ord('0')) and (Shift = [ssCtrl])) then
  begin
    if pmRestoreFont.Enabled = True then
    begin
      SetHeadings(pmRestoreFont);
      Key := 0;
    end;
  end;
  // Shortcut for inserting current date and time in a new line
  if ((Key = Ord('D')) and (Shift = [ssCtrl, ssShift, ssAlt])) then
  begin
    dbText.InsertText(dbText.SelStart, FormatDateTime(FDate.LongDateFormat, Now));
    EditNotesDataset;
    Key := 0;
  end;
  // Clean paragraph
  if ((Key = Ord('P')) and (Shift = [ssCtrl, ssShift, ssAlt])) then
  begin
    dbText.CleanParagraph(dbText.SelStart);
    EditNotesDataset;
    Key := 0;
  end;
  // Delete line with Ctrl + Shift Y (not Ctrl only to avoid mistyping)
  if ((Key = 89) and (Shift = [ssShift, ssCtrl])) then
  begin
    dbText.Lines.Delete(dbText.CaretPos.Y);
    EditNotesDataset;
    Key := 0;
  end;
  // Move up paragraph
  if ((Key = 38) and (Shift = [ssShift, ssCtrl])) then
  begin
    iChar := dbText.SelStart;
    dbText.MoveParUpDown(dbText.CaretPos.Y, True);
    dbText.SelStart := iChar - UTF8Length(dbText.Lines[dbText.CaretPos.Y - 1]) - 1;
    // Using SelStart, the current line becomes the last one
    // To recenter it...
    Application.ProcessMessages;
    dbText.SetCursorMiddleScreen(dbText.CaretPos.Y);
    dbText.GetTextAttributes(dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True), fp);
    if fp.Indented > DefaultIndent then
    begin
      NewItemBullet := GetBullet;
      if NewItemBullet = 2 then
        dbText.ListNumber(dbText.SelStart, WidthInden, '1')
      else if NewItemBullet = 3 then
        dbText.ListNumber(dbText.SelStart, WidthInden, 'A')
      else if NewItemBullet = 4 then
        dbText.ListNumber(dbText.SelStart, WidthInden, 'a');
    end;
    EditNotesDataset;
    key := 0;
  end;
  // Move down paragraph
  if ((Key = 40) and (Shift = [ssShift, ssCtrl])) then
  begin
    iChar := dbText.SelStart;
    dbText.MoveParUpDown(dbText.CaretPos.Y, False);
    dbText.SelStart := iChar + UTF8Length(dbText.Lines[dbText.CaretPos.Y]) + 1;
    // Using SelStart, the current line becomes the last one
    // To recenter it...
    Application.ProcessMessages;
    dbText.SetCursorMiddleScreen(dbText.CaretPos.Y);
    dbText.GetTextAttributes(dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True), fp);
    if fp.Indented > DefaultIndent then
    begin
      NewItemBullet := GetBullet;
      if NewItemBullet = 2 then
        dbText.ListNumber(dbText.SelStart, WidthInden, '1')
      else if NewItemBullet = 3 then
        dbText.ListNumber(dbText.SelStart, WidthInden, 'A')
      else if NewItemBullet = 4 then
        dbText.ListNumber(dbText.SelStart, WidthInden, 'a');
    end;
    EditNotesDataset;
    key := 0;
  end;
  // Zoom increase with +
  if ((Key = 187) and (Shift = [ssCtrl])) then
    ZoomIncrease;
  // Zoom decrease with -
  if ((Key = 189) and (Shift = [ssCtrl])) then
    ZoomDecrease;
end;

procedure TfmMain.dbTextKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
var
  iBullet: integer;
begin
  // Link after the text has been pasted;
  if ((Shift = [ssCtrl]) and (Key = Ord('V'))) then
  begin
    MakeLink(CheckLink);
  end;
  // Toggle bullets with Ctrl + .
  if ((Key = 190) and (Shift = [ssCtrl, ssShift])) then
  begin
    dbText.ListNumber(dbText.SelStart, 0, 'X');
    EditNotesDataset;
  end
  else if ((Key = 190) and (Shift = [ssCtrl, ssAlt])) then
  begin
    iBullet := GetBullet;
    if iBullet = 2 then
      dbText.ListNumber(dbText.SelStart, WidthInden, '1')
    else if iBullet = 3 then
      dbText.ListNumber(dbText.SelStart, WidthInden, 'A')
    else if iBullet = 4 then
      dbText.ListNumber(dbText.SelStart, WidthInden, 'a');
    EditNotesDataset;
  end
  else if ((Key = 190) and (Shift = [ssCtrl])) then
  begin
    iBullet := GetBullet;
    if iBullet = 0 then
      dbText.ListNumber(dbText.SelStart, WidthInden, '*')
    else if iBullet = 1 then
      dbText.ListNumber(dbText.SelStart, WidthInden, '1')
    else if iBullet = 2 then
      dbText.ListNumber(dbText.SelStart, WidthInden, 'A')
    else if iBullet = 3 then
      dbText.ListNumber(dbText.SelStart, WidthInden, 'a')
    else if iBullet = 4 then
      dbText.ListNumber(dbText.SelStart, 0, 'X');
    EditNotesDataset;
  end;
  // Todo characters
  if ((Key = Ord('R')) and (Shift = [ssCtrl, ssShift])) then
  begin
    dbText.SetToDo(dbText.CaretPos.Y, True);
    EditNotesDataset;
  end;
  if ((Key = Ord('T')) and (Shift = [ssCtrl, ssShift])) then
  begin
    dbText.SetToDo(dbText.CaretPos.Y, False);
    EditNotesDataset;
  end;
  // Activate and deactivate format icons
  ActivateFormatIcons;
  // Count char
  SetCharCount(flNoCharCount);
end;

procedure TfmMain.dbTextMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
var
  StartSel, EndSel: integer;
  ThingToRun: string;
begin
  // Run the link
  if ssCtrl in Shift then
  begin
    StartSel := CheckLink;
    if StartSel > -1 then
    begin
      EndSel := dbText.SelStart;
      while ((UTF8Copy(dbText.Text, EndSel, 1) <> ' ') and (UTF8Copy(dbText.Text, EndSel, 1) <> LineEnding) and (UTF8Copy(dbText.Text, EndSel, 1) <> #9) and (EndSel <= UTF8Length(dbText.Text))) do
        EndSel := EndSel + 1;
      if ((UTF8Copy(dbText.Text, EndSel - 1, 1) = '.') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = ',') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = ';') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = ':') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = '?') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = '!')) then
        EndSel := EndSel - 1;
      if ((UTF8Copy(dbText.Text, EndSel - 1, 1) = ')') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = ']') or (UTF8Copy(dbText.Text, EndSel - 1, 1) = '}')) then
        EndSel := EndSel - 1;
      ThingToRun := UTF8Copy(dbText.Text, startSel + 1, EndSel - StartSel - 1);
      if UTF8Copy(ThingToRun, 1, 1) = '#' then
      begin
        if miNotesFind.Checked = False then
        begin
          miNotesFindClick(nil);
          edFindText.Text := ThingToRun;
          cbFindKind.ItemIndex := 4;
          FindData(False);
        end;
      end
      else if UTF8Copy(ThingToRun, 1, 6) = 'mnt://' then
      begin
        ThingToRun := UTF8Copy(ThingToRun, 7, Length(ThingToRun));
        ThingToRun := StringReplace(ThingToRun, '_', ' ', [rfReplaceAll, rfIgnoreCase]);
        if UTF8Pos('/', ThingToRun) > 0 then
        begin
          IsTextToLoad := False;
          sqSubjects.Locate('SubjectsName',
            UTF8Copy(ThingToRun, 1, UTF8Pos('/', ThingToRun) - 1),
            [loCaseInsensitive]);
          sqNotes.Locate('NotesTitle',
            UTF8Copy(ThingToRun, UTF8Pos('/', ThingToRun) + 1, UTF8Length(ThingToRun)), [loCaseInsensitive]);
          IsTextToLoad := True;
          sqNotesAfterScroll(nil);
        end
        else
          sqSubjects.Locate('SubjectsName', ThingToRun,
            [loCaseInsensitive]);
      end
      else
      begin
        if UTF8Copy(ThingToRun, 1, 7) = 'file://' then
        begin
          ThingToRun := StringReplace(ThingToRun, '_', ' ', [rfReplaceAll, rfIgnoreCase]);
          ThingToRun := UTF8Copy(ThingToRun, 7, Length(ThingToRun));
          OpenDocument(ThingToRun);
        end
        else
        begin
          if UTF8Copy(ThingToRun, 1, 4) = 'www.' then
            ThingToRun := 'http://' + ThingToRun;
          OpenURL(ThingToRun);
        end;
      end;
    end;
  end;
end;

procedure TfmMain.dbTextMouseWheel(Sender: TObject; Shift: TShiftState; WheelDelta: integer; MousePos: TPoint; var Handled: boolean);
begin
  // Change Zoom with mouse weel
  if Shift = [ssCtrl] then
  begin
    if WheelDelta > 0 then
      ZoomIncrease
    else
      ZoomDecrease;
  end;
end;

procedure TfmMain.dbTitleKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Create title in the text of the notes
  if ((Key = 13) and (Shift = [ssCtrl])) then
  begin
    dbText.InsertText(0, dbTitle.Text + LineEnding + LineEnding + LineEnding);
    dbText.SelStart := 0;
    // This is necessary to let TRichMemo to know selection start
    Application.ProcessMessages;
    SetHeadings(pmHeading1);
    dbTags.SetFocus;
    Key := 0;
  end;
end;

procedure TfmMain.dbTitleKeyPress(Sender: TObject; var Key: char);
begin
  // Select tags field from title
  if ((Key = #13) or (Key = #9)) then
  begin
    dbTags.SetFocus;
    Key := #0;
  end;
end;

procedure TfmMain.dbTagsKeyPress(Sender: TObject; var Key: char);
begin
  // Select date field from tag
  if ((Key = #13) or (Key = #9)) then
  begin
    dbText.SetFocus;
    Key := #0;
  end;
end;

procedure TfmMain.dbDateKeyPress(Sender: TObject; var Key: char);
begin
  // Select text field from date
  if ((Key = #13) or (Key = #9)) then
  begin
    dbText.SetFocus;
    Key := #0;
  end;
end;

procedure TfmMain.dbDateCloseUp(Sender: TObject);
begin
  // Save data on close the calendar
  // Necessary to avoid a bug: see below.
  SaveAllData;
end;

procedure TfmMain.edFindTextKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Find data on Return key
  if ((Key = 13) and (Shift = [ssCtrl])) then
  begin
    if bnFind2.Enabled = True then
    begin
      FindData(True);
      Key := 0;
    end;
  end
  else if ((Key = 13) and (Shift = [])) then
  begin
    FindData(False);
    Key := 0;
  end;
end;

procedure TfmMain.grSubjectsKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Move down
  if ((Key = 40) and (Shift = [ssCtrl])) then
  begin
    if bnSubDown.Enabled = True then
      bnSubDownClick(nil);
    Key := 0;
  end
  // Move up
  else if ((Key = 38) and (Shift = [ssCtrl])) then
  begin
    if bnSubUp.Enabled = True then
      bnSubUpClick(nil);
    Key := 0;
  end
  // Save or edit subjects on Enter key
  else if Key = 13 then
  begin
    if dsSubjects.State in [dsBrowse] then
      sqSubjects.Edit
    else if dsSubjects.State in [dsEdit, dsInsert] then
      SaveAllData;
    Key := 0;
    grSubjects.SetFocus;
  end;
end;

procedure TfmMain.grSubjectsPrepareCanvas(Sender: TObject; DataCol: integer; Column: TColumn; AState: TGridDrawState);
begin
  // Subject color
  // Only if the raw is not selected
  if not (gdSelected in AState) then
  begin
    with (Sender as TDBGrid) do
    begin
      if sqSubjects.FieldByName('SubjectsBackColor').AsString <> '' then
      begin
        Canvas.Brush.Color :=
          StringToColor(sqSubjects.FieldByName('SubjectsBackColor').AsString);
      end;
      if sqSubjects.FieldByName('SubjectsFontColor').AsString <> '' then
      begin
        Canvas.Font.Color :=
          StringToColor(sqSubjects.FieldByName('SubjectsFontColor').AsString);
      end;
    end;
  end;
end;

procedure TfmMain.grSubjectsExit(Sender: TObject);
begin
  // Save subjects on exit
  SaveAllData;
end;

procedure TfmMain.bnFindClick(Sender: TObject);
begin
  // Find data on clic
  FindData(False);
end;

procedure TfmMain.bnFind2Click(Sender: TObject);
begin
  // Find data within the results
  FindData(True);
end;

procedure TfmMain.edLocateTextKeyPress(Sender: TObject; var Key: char);
begin
  // Find First in current note text on Return key
  if Key = #13 then
  begin
    FindFirstNext(bnFindFirst);
    Key := #0;
  end;
end;

procedure TfmMain.lbAttNamesDblClick(Sender: TObject);
begin
  // New attachment
  if lbAttNames.Items.Count = 0 then
  begin
    if miAttachNew.Enabled = True then
      miAttachNewClick(nil);
  end
  // Open attachment
  else if miAttachOpen.Enabled = True then
    miAttachOpenClick(nil);
end;

procedure TfmMain.lbAttNamesExit(Sender: TObject);
begin
  // To avoid a permanent selection
  lbAttNames.ItemIndex := -1;
end;

procedure TfmMain.lbTagsNamesDblClick(Sender: TObject);
var
  StTag: string;
begin
  // Save current tag in dbTags or in search field
  if lbTagsNames.Items.Count = 0 then
    Abort;
  StTag := lbTagsNames.Items.ValueFromIndex[lbTagsNames.ItemIndex];
  // There is onlty the message that tags are too numerous to create the list
  if UTF8Pos('[', StTag) = 0 then
    Exit;
  if miNotesFind.Checked = False then
  begin
    if sqNotes.RecordCount > 0 then
    begin
      if dsNotes.State in [dsBrowse] then
        sqNotes.Edit;
      if dbTags.Text = '' then
        dbTags.Text := UTF8Copy(StTag, 1, UTF8Pos('[', StTag) - 2)
      else
        dbTags.Text := dbTags.Text + ', ' + UTF8Copy(StTag, 1, UTF8Pos('[', StTag) - 2);
      dbTags.Text := CleanTagsField(dbTags.Text);
    end;
  end
  else
  begin
    // Search in tags
    cbFindKind.ItemIndex := 5;
    if edFindText.Text = '' then
      edFindText.Text := UTF8Copy(StTag, 1, UTF8Pos('[', StTag) - 2)
    else
      edFindText.Text := edFindText.Text + ', ' + UTF8Copy(StTag, 1, UTF8Pos('[', StTag) - 2);
    edFindText.Text := CleanTagsField(edFindText.Text);
  end;
end;

procedure TfmMain.lbTagsNamesExit(Sender: TObject);
begin
  // To avoid a permanent selection
  lbTagsNames.ItemIndex := -1;
end;

procedure TfmMain.grTitlesDblClick(Sender: TObject);
begin
  // Set the focus on title on double clic
  if dbTitle.Visible = True then
    dbTitle.SetFocus;
end;

procedure TfmMain.grTitlesDrawCell(Sender: TObject; aCol, aRow: integer; aRect: TRect; aState: TGridDrawState);
var
  iIndLock, iIndAct: integer;
begin
  // Set the alternate color from encrypted notes or for activities
  iIndLock := 0;
  iIndAct := 0;
  if aRow = 0 then
  begin
    grTitles.Canvas.Brush.Color := grTitles.FixedColor;
  end
  else
  begin
    if not (gdSelected in aState) then
    begin
      if ((aCol > 0) and (aRow > 0)) then
      begin
        if grTitles.Cells[5, aRow] <> '' then
        begin
          grTitles.Canvas.Brush.Color :=
            StringToColor(grTitles.Cells[5, aRow]);
        end;
        if grTitles.Cells[6, aRow] <> '' then
        begin
          grTitles.Canvas.Font.Color :=
            StringToColor(grTitles.Cells[6, aRow]);
        end;
      end;
    end;
    if ((aCol = 1) and (aRow > 0)) then
    begin
      if grTitles.Cells[3, aRow] = 'T' then
        iIndLock := 20;
      if grTitles.Cells[4, aRow] = 'T' then
        iIndAct := 20;
    end;
  end;
  grTitles.Canvas.FillRect(aRect);
  grTitles.Canvas.TextOut(aRect.Left + 2 + iIndLock + iIndAct,
    aRect.Top + 3, grTitles.Cells[aCol, aRow]);
  if iIndLock > 0 then
    ilNotes.Draw(grTitles.Canvas, aRect.Left + 2, aRect.Top + 4, 0);
  if iIndAct > 0 then
  begin
    if iIndLock > 0 then
      ilNotes.Draw(grTitles.Canvas, aRect.Left + 2 + iIndLock, aRect.Top + 4, 1)
    else
      ilNotes.Draw(grTitles.Canvas, aRect.Left + 2, aRect.Top + 4, 1);
  end;
end;

procedure TfmMain.grTitlesKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Move down
  if ((Key = 40) and (Shift = [ssCtrl])) then
  begin
    if bnNotesDown.Enabled = True then
      bnNotesDownClick(nil);
    Key := 0;
  end
  // Move up
  else if ((Key = 38) and (Shift = [ssCtrl])) then
  begin
    if bnNotesUp.Enabled = True then
      bnNotesUpClick(nil);
    Key := 0;
  end;
end;

procedure TfmMain.lbPwdTopResize(Sender: TObject);
begin
  // Resize the elements of password panel when the form is shrinked
  edPassword.Top := lbPwdTop.Top + lbPwdTop.Height + 10;
  lbPwdBottom.Top := edPassword.Top + edPassword.Height + 10;
  edPassword.Width := lbPwdTop.Width - 40;
  lbPwdBottom.Width := lbPwdTop.Width;
end;

// *****************************************************************************
// ************************* DATASET PROCEDURES ********************************
// *****************************************************************************

// ************************* SUBJECTS PROCEDURES *******************************

procedure TfmMain.dsSubjectsStateChange(Sender: TObject);
begin
  // Active & deactivate subjects' menus
  if ((dsSubjects.State in [dsEdit, dsInsert]) or (dsNotes.State in [dsEdit, dsInsert])) then
  begin
    miFileSave.Enabled := True;
    tbFileSave.Enabled := True;
    miFileUndo.Enabled := True;
  end
  else
  begin
    miFileSave.Enabled := False;
    tbFileSave.Enabled := False;
    miFileUndo.Enabled := False;
  end;
  if dsSubjects.State in [dsEdit, dsInsert] then
  begin
    miSubjectNew.Enabled := False;
    tbSubjectNew.Enabled := False;
    miSubjectDelete.Enabled := False;
    miSubjectComments.Enabled := False;
    miSubjectLook.Enabled := False;
    miSubjectOrder.Enabled := False;
    miSubjectOrderTitle.Enabled := False;
    miSubjectOrderCustom.Enabled := False;
    pmSubNew.Enabled := False;
    pmSubDelete.Enabled := False;
    pmSubComments.Enabled := False;
    pmSubLook.Enabled := False;
  end
  else
  begin
    miSubjectNew.Enabled := True;
    tbSubjectNew.Enabled := True;
    miSubjectDelete.Enabled := True;
    miSubjectComments.Enabled := True;
    miSubjectLook.Enabled := True;
    miSubjectOrder.Enabled := True;
    miSubjectOrderTitle.Enabled := True;
    miSubjectOrderCustom.Enabled := True;
    pmSubNew.Enabled := True;
    pmSubDelete.Enabled := True;
    pmSubComments.Enabled := True;
    pmSubLook.Enabled := True;
  end;
end;

procedure TfmMain.dsSubjectsDataChange(Sender: TObject; Field: TField);
begin
  // If Subjects data changes without a State change
  if sqSubjects.RecordCount = 0 then
  begin
    miSubjectDelete.Enabled := False;
    pmSubDelete.Enabled := False;
    miSubjectComments.Enabled := False;
    miSubjectLook.Enabled := False;
    miSubjectOrder.Enabled := False;
    miSubjectOrderTitle.Enabled := False;
    miSubjectOrderCustom.Enabled := False;
    pmSubComments.Enabled := False;
    pmSubLook.Enabled := False;
  end
  else
  begin
    miSubjectDelete.Enabled := dsSubjects.State in [dsBrowse];
    pmSubDelete.Enabled := dsSubjects.State in [dsBrowse];
    miSubjectComments.Enabled := dsSubjects.State in [dsBrowse];
    miSubjectLook.Enabled := dsSubjects.State in [dsBrowse];
    miSubjectOrder.Enabled := dsSubjects.State in [dsBrowse];
    miSubjectOrderTitle.Enabled := dsSubjects.State in [dsBrowse];
    miSubjectOrderCustom.Enabled := dsSubjects.State in [dsBrowse];
    pmSubComments.Enabled := dsSubjects.State in [dsBrowse];
    pmSubLook.Enabled := dsSubjects.State in [dsBrowse];
  end;
end;

procedure TfmMain.sqSubjectsAfterPost(DataSet: TDataSet);
var
  IDSub: integer;
begin
  // To recreate alphabetic order in subjects after post
  // if the subject is not empty (= new)
  if sqSubjects.FieldByName('SubjectsName').AsString <> '' then
  begin
    ;
    IDSub := sqSubjects.FieldByName('IDSubjects').AsInteger;
    // To avoid that the text is loaded
    IsTextToLoad := False;
    sqSubjects.RefetchData;
    sqSubjects.Locate('IDSubjects', IDSub, []);
    IsTextToLoad := True;
    // Now load the text
    sqNotesAfterScroll(nil);
    // File changed
    flFileChanged := True;
  end;
end;

procedure TfmMain.sqSubjectsAfterScroll(DataSet: TDataSet);
begin
  // Select the Notes corresponding to the subject
  with sqNotes do
  begin
    Close;
    if miOrderByTitle.Checked = True then
    begin
      SQL := 'Select * from Notes where ID_Subjects = ' + sqSubjects.FieldByName('IDSubjects').AsString + ' order by NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      SQL := 'Select * from Notes where Notes.ID_Subjects = ' + sqSubjects.FieldByName('IDSubjects').AsString + ' order by NotesDate, IDNotes';
    end
    else
    begin
      SQL := 'Select * from Notes where Notes.ID_Subjects = ' + sqSubjects.FieldByName('IDSubjects').AsString + ' order by NotesSort';
    end;
    // To show the last note if order by note is active
    if miOrderByDate.Checked = True then
    begin
      // sqNotes.DisableControls does not work, so...
      if IsTextToLoad = True then
      begin
        IsTextToLoad := False;
        Open;
        IsTextToLoad := True;
        // Call to last with recordcount = 0 raises an error
        if RecordCount > 0 then
          Last;
      end
      // IsTextToLoad could be set to False by OpenDataTables
      else
      begin
        Open;
        // Call to last with recordcount = 0 raises an error
        if RecordCount > 0 then
          Last;
      end;
    end
    else
    begin
      Open;
    end;
    if sqNotes.RecordCount = 0 then
    begin
      dbText.Clear;
      lbAttNames.Clear;
      grActGrid.Clean(1, 1, grActGrid.ColCount - 1,
        grActGrid.RowCount - 1, [gzNormal]);
      grActGrid.Row := 1;
      LastRowAct := 1;
      meActNotes.Clear;
      SetCharCount(True);
      ClearUndoData;
    end;
  end;
  // Load title and date in title grid
  LoadTitleDateGrid;
  // To activate menu items and buttons
  dsNotesDataChange(nil, nil);
end;

procedure TfmMain.sqSubjectsBeforeDelete(DataSet: TDataSet);
begin
  // Confirm for deletion and clear all related notes
  if MessageDlg(msg001, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end
  else
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      IsTextToLoad := False;
      while sqNotes.RecordCount > 0 do
      begin
        // Set the flag to recreate tags list
        if sqNotes.FieldByName('NotesTags').AsString <> '' then
          RecreateTagsList := True;
        sqNotes.Delete;
        sqNotes.ApplyUpdates;
      end;
      if RecreateTagsList = True then
      begin
        CreateTagsList;
        RecreateTagsList := False;
      end;
      IsTextToLoad := True;
      // Save the UID of the record to be deleted for sync
      with sqDelRec do
      begin
        Open;
        Append;
        sqDelRec.FieldByName('DelRecUID').AsString :=
          sqSubjects.FieldByName('SubjectsUID').AsString;
        sqDelRec.FieldByName('DelRecDTMod').AsDateTime := Now;
        Post;
        ApplyUpdates;
        Close;
      end;
    finally
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.sqSubjectsBeforeInsert(DataSet: TDataSet);
begin
  // Check before creating subject
  if ConfirmNewSubject = True then
  begin
    if MessageDlg(msg002, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
      Abort;
  end;
end;

procedure TfmMain.sqSubjectsBeforePost(DataSet: TDataSet);
begin
  // Update date-time of last editing
  sqSubjects.FieldByName('SubjectsDTMod').AsDateTime := Now;
end;

procedure TfmMain.sqSubjectsAfterDelete(DataSet: TDataSet);
begin
  // ApplyUpdates is in sqSubjects.AfterDelete, so that it can work also
  // for deletion with Ctrl+Canc from the grid
  sqSubjects.ApplyUpdates;
end;

procedure TfmMain.sqSubjectsAfterInsert(DataSet: TDataSet);
var
  myGUID: TGUID;
begin
  // Create UID and sort
  CreateGUID(myGUID);
  sqSubjects.FieldByName('SubjectsUID').AsString :=
    Copy(GUIDToString(myGUID), 2, Length(GUIDToString(myGUID)) - 2);
  sqSubjects.FieldByName('SubjectsSort').AsInteger :=
    sqSubjects.FieldByName('IDSubjects').AsInteger;
  // Save
  sqSubjects.Post;
  sqSubjects.ApplyUpdates;
  grSubjects.SetFocus;
end;

procedure TfmMain.sqSubjectsBeforeScroll(DataSet: TDataSet);
begin
  // Save data before scroll if richedit is modified
  if IsMemoModified = True then
  begin
    sqNotes.Edit;
    sqNotes.Post;
  end;
end;

// ************************* NOTES PROCEDURES **********************************

procedure TfmMain.dsNotesDataChange(Sender: TObject; Field: TField);
var
  fmTime: string = 'hh:nn am/pm';
begin
  // If Notes data changes without a State change
  if ((sqNotes.RecordCount = 0) and (dsNotes.State in [dsBrowse])) then
  begin
    miNotesDelete.Enabled := False;
    pmNotesDelete.Enabled := False;
    pmNotesLook.Enabled := False;
    miNotesEncDecrypt.Enabled := False;
    miNotesMove.Enabled := False;
    miNotesSendToWp.Enabled := False;
    miNotesSendToOO.Enabled := False;
    miNotesSendToLO.Enabled := False;
    miNotesSendToBrowser.Enabled := False;
    miNotesAttach.Enabled := False;
    miNotesimages.Enabled := False;
    miAttachNew.Enabled := False;
    miAttachOpen.Enabled := False;
    miAttachSaveAs.Enabled := False;
    miAttachDelete.Enabled := False;
    miNotesLook.Enabled := False;
    pmAttNew.Enabled := False;
    pmAttOpen.Enabled := False;
    pmAttSaveAs.Enabled := False;
    pmAttDelete.Enabled := False;
    tbCut.Enabled := False;
    tbCopy.Enabled := False;
    tbCopyHtml.Enabled := False;
    tbPaste.Enabled := False;
    tbKindFontChange.Enabled := False;
    tbColorFontChange.Enabled := False;
    tbBackColorFontChange.Enabled := False;
    tbFontBold.Enabled := False;
    tbFontItalic.Enabled := False;
    tbFontUnderline.Enabled := False;
    tbFontStrike.Enabled := False;
    tbFontRestore.Enabled := False;
    tbAlignLeft.Enabled := False;
    tbAlignCenter.Enabled := False;
    tbAlignRight.Enabled := False;
    tbAlignFill.Enabled := False;
    tbAlignIndent.Enabled := False;
    tbOpenNote.Enabled := False;
    grActGrid.Enabled := False;
    tbActIndLeft.Enabled := False;
    tbActIndRight.Enabled := False;
    tbActMoveUp.Enabled := False;
    tbActMoveDown.Enabled := False;
    tbActInsertRow.Enabled := False;
    tbActDeleteRow.Enabled := False;
    tbActShowWbs.Enabled := False;
    tbActCopyGroup.Enabled := False;
    tbActPasteGroup.Enabled := False;
    tbActCopyAll.Enabled := False;
    tbActClearAll.Enabled := False;
    meActNotes.Enabled := False;
    SetCharCount(True);
  end
  else
  begin
    miNotesMove.Enabled := True;
    miNotesDelete.Enabled := dsNotes.State in [dsBrowse];
    pmNotesDelete.Enabled := dsNotes.State in [dsBrowse];
    pmNotesLook.Enabled := dsNotes.State in [dsBrowse];
    miNotesEncDecrypt.Enabled := dsNotes.State in [dsBrowse];
    if pnPassword.Visible = False then
    begin
      miNotesSendToWp.Enabled := dsNotes.State in [dsBrowse];
      miNotesSendToOO.Enabled := dsNotes.State in [dsBrowse];
      miNotesSendToLO.Enabled := dsNotes.State in [dsBrowse];
      miNotesSendToBrowser.Enabled := dsNotes.State in [dsBrowse];
    end
    else
    begin
      miNotesSendToWp.Enabled := False;
      miNotesSendToOO.Enabled := False;
      miNotesSendToLO.Enabled := False;
      miNotesSendToBrowser.Enabled := False;
    end;
    miNotesImages.Enabled := True;
    miNotesAttach.Enabled := True;
    miAttachNew.Enabled := True;
    miAttachOpen.Enabled := sqNotes.FieldByName('NotesAttName').AsString <> '';
    miAttachSaveAs.Enabled := miAttachOpen.Enabled;
    miAttachDelete.Enabled := miAttachOpen.Enabled;
    miNotesLook.Enabled := dsNotes.State in [dsBrowse];
    pmAttNew.Enabled := miAttachNew.Enabled;
    pmAttOpen.Enabled := miAttachOpen.Enabled;
    pmAttSaveAs.Enabled := miAttachSaveAs.Enabled;
    pmAttDelete.Enabled := miAttachDelete.Enabled;
    tbCut.Enabled := not pnPassword.Visible;
    tbCopy.Enabled := not pnPassword.Visible;
    tbCopyHtml.Enabled := not pnPassword.Visible;
    tbPaste.Enabled := not pnPassword.Visible;
    tbKindFontChange.Enabled := not pnPassword.Visible;
    tbColorFontChange.Enabled := not pnPassword.Visible;
    tbBackColorFontChange.Enabled := not pnPassword.Visible;
    tbFontBold.Enabled := not pnPassword.Visible;
    tbFontItalic.Enabled := not pnPassword.Visible;
    tbFontUnderline.Enabled := not pnPassword.Visible;
    tbFontStrike.Enabled := not pnPassword.Visible;
    tbFontRestore.Enabled := not pnPassword.Visible;
    tbAlignLeft.Enabled := not pnPassword.Visible;
    tbAlignCenter.Enabled := not pnPassword.Visible;
    tbAlignRight.Enabled := not pnPassword.Visible;
    tbAlignFill.Enabled := not pnPassword.Visible;
    tbAlignIndent.Enabled := not pnPassword.Visible;
    tbOpenNote.Enabled := not pnPassword.Visible;
    grActGrid.Enabled := True;
    tbActIndLeft.Enabled := True;
    tbActIndRight.Enabled := True;
    tbActMoveUp.Enabled := True;
    tbActMoveDown.Enabled := True;
    tbActInsertRow.Enabled := True;
    tbActDeleteRow.Enabled := True;
    tbActShowWbs.Enabled := True;
    tbActCopyGroup.Enabled := True;
    tbActPasteGroup.Enabled := True;
    tbActCopyAll.Enabled := True;
    tbActClearAll.Enabled := True;
    meActNotes.Enabled := True;
    SetCharCount(flNoCharCount);
  end;
  // Update notes numbers in status bar;
  // it is modified also in miToolsLanguageClick
  if dsNotes.State in [dsInsert, dsEdit] then
  begin
    sbStatusBar.Panels[0].Text := ' ' + sbr001;
  end
  else if sqNotes.RecordCount = 0 then
  begin
    sbStatusBar.Panels[0].Text := ' ' + sbr002;
    SetCharCount(True);
  end
  else
  begin
    if fl24Hour = True then
    begin
      fmTime := 'hh:nn';
    end;
    sbStatusBar.Panels[0].Text :=
      ' ' + sbr003 + ' ' + IntToStr(sqNotes.RecNo) + ' ' + sbr004 + ' ' + IntToStr(sqNotes.RecordCount) + ' - ' + sbr009 + ' ' + FormatDateTime(FDate.LongDateFormat, sqNotes.FieldByName('NotesDTMod').AsDateTime) + ' ' + sbr010 + ' ' + FormatDateTime(fmTime, sqNotes.FieldByName('NotesDTMod').AsDateTime);
  end;
end;

procedure TfmMain.dsNotesStateChange(Sender: TObject);
begin
  // Active & deactivate notes menus
  if ((dsSubjects.State in [dsEdit, dsInsert]) or (dsNotes.State in [dsEdit, dsInsert])) then
  begin
    miFileSave.Enabled := True;
    tbFileSave.Enabled := True;
    miFileUndo.Enabled := True;
  end
  else
  begin
    miFileSave.Enabled := False;
    tbFileSave.Enabled := False;
    miFileUndo.Enabled := False;
  end;
  if dsNotes.State in [dsEdit, dsInsert] then
  begin
    miNotesNew.Enabled := False;
    tbNotesNew.Enabled := False;
    miNotesDelete.Enabled := False;
    miNotesEncDecrypt.Enabled := False;
    miNotesSendToWp.Enabled := False;
    miNotesSendToOO.Enabled := False;
    miNotesSendToLO.Enabled := False;
    miNotesSendToBrowser.Enabled := False;
    miNotesLook.Enabled := False;
    pmNotesNew.Enabled := False;
    pmNotesDelete.Enabled := False;
    pmNotesLook.Enabled := False;
  end
  else
  begin
    miNotesNew.Enabled := True;
    tbNotesNew.Enabled := True;
    miNotesDelete.Enabled := True;
    miNotesEncDecrypt.Enabled := True;
    miNotesSendToWp.Enabled := True;
    miNotesSendToOO.Enabled := True;
    miNotesSendToLO.Enabled := True;
    miNotesSendToBrowser.Enabled := True;
    miNotesLook.Enabled := True;
    pmNotesNew.Enabled := True;
    pmNotesDelete.Enabled := True;
    pmNotesLook.Enabled := True;
  end;
end;

procedure TfmMain.sqNotesBeforeInsert(DataSet: TDataSet);
begin
  // Avoid notes when there are no subjects
  if sqSubjects.RecordCount = 0 then
  begin
    MessageDlg(msg003, mtWarning, [mbOK], 0);
    Abort;
  end;
end;

procedure TfmMain.sqNotesBeforePost(DataSet: TDataSet);
begin
  // Update date-time of last editing
  sqNotes.FieldByName('NotesDTMod').AsDateTime := Now;
  sqNotes.FieldByName('NotesDateFormat').AsString := DateOrder;
  // Format Tag field
  sqNotes.FieldByName('NotesTags').AsString := CleanTagsField(dbTags.Text);
  // Save data if necessary and set flag to recreate tags list
  if sqNotes.FieldByName('NotesTags').AsString <> OldTagsValue then
    RecreateTagsList := True;
  if IsMemoModified = True then
  begin
    if sqNotes.FieldByName('NotesCheckPwd').AsString = '' then
      sqNotes.FieldByName('NotesText').AsWideString :=
        SaveRichMemo(0, UTF8Length(dbText.Text), True)
    else
    begin
      dcAES.InitStr(edPassword.Text, TDCP_sha1);
      sqNotes.FieldByName('NotesText').AsWideString :=
        dcAES.EncryptString(SaveRichMemo(0, UTF8Length(dbText.Text), True));
    end;
    IsMemoModified := False;
  end;
  // Save activities
  SaveActivitiesData;
end;

procedure TfmMain.sqNotesBeforeScroll(DataSet: TDataSet);
begin
  // Save data before scroll if richedit is modified
  if IsMemoModified = True then
  begin
    sqNotes.Edit;
    sqNotes.Post;
  end;
end;

procedure TfmMain.sqNotesAfterPost(DataSet: TDataSet);
begin
  // Recreate tags list if necessary
  // To complete data saving; no problem if it is repeated in SaveAllData
  sqNotes.ApplyUpdates;
  if RecreateTagsList = True then
  begin
    CreateTagsList;
    RecreateTagsList := False;
  end;
  // Load title and date in title grid
  LoadTitleDateGrid;
  // Save the current tags value
  OldTagsValue := sqNotes.FieldByName('NotesTags').AsString;
  // File changed
  flFileChanged := True;
  // Reload undo text
  StoreUndoData;
end;

procedure TfmMain.sqNotesAfterDelete(DataSet: TDataSet);
begin
  // Load title and date in title grid
  // To complete data saving; no problem if it is repeated in SaveAllData
  sqNotes.ApplyUpdates;
  ClearUndoData;
  // Load title and date in title grid
  LoadTitleDateGrid;
end;

procedure TfmMain.sqNotesAfterInsert(DataSet: TDataSet);
var
  myGUID: TGUID;
begin
  // Insert link value
  sqNotes.FieldByName('ID_Subjects').AsInteger :=
    sqSubjects.FieldByName('IDSubjects').AsInteger;
  // Insert today's date
  sqNotes.FieldByName('NotesDate').AsDateTime := Date;
  // Create UID and sort
  CreateGUID(myGUID);
  sqNotes.FieldByName('NotesUID').AsString :=
    Copy(GUIDToString(myGUID), 2, Length(GUIDToString(myGUID)) - 2);
  sqNotes.FieldByName('NotesSort').AsInteger :=
    sqNotes.FieldByName('IDNotes').AsInteger;
  LoadActivitiesData;
  // To sync title grid
  sqNotes.Post;
  ClearUndoData;
end;

procedure TfmMain.sqNotesAfterScroll(DataSet: TDataSet);
begin
  // load Text in dbText
  dbText.Clear;
  if sqNotes.RecordCount > 0 then
  begin
    if sqNotes.FieldByName('NotesCheckPwd').AsString <> '' then
    begin
      ShowPasswordInput;
    end
    else
    begin
      HidePasswordInput;
      LoadRichMemo;
    end;
    LoadActivitiesData;
  end
  else
  begin
    HidePasswordInput;
    grActGrid.Clean(1, 1, grActGrid.ColCount - 1,
      grActGrid.RowCount - 1, [gzNormal]);
    grActGrid.Row := 1;
    LastRowAct := 1;
    meActNotes.Clear;
  end;
  // To update the attachment list
  lbAttNames.Items.Text := sqNotes.FieldByName('NotesAttName').AsString;
  lbListAttach.Caption := stAttachments + ' [' + IntToStr(lbAttNames.Items.Count) + ']';
  // Save the current tags value
  OldTagsValue := sqNotes.FieldByName('NotesTags').AsString;
  // Update title grid
  SyncTitleDateGrid;
end;

procedure TfmMain.sqNotesBeforeDelete(DataSet: TDataSet);
var
  AttDir: string;
  i: integer;
  myStringList: TStringList;
begin
  // The check for delete is in the function menu and not here
  // because the notes must be deleted without confirmation
  // when deleting a subject.
  // Delete attachment
  AttDir := ExtractFileNameWithoutExt(sqNotes.FileName);
  if sqNotes.FieldByName('NotesAttName').AsString <> '' then
    try
      if DirectoryExistsUTF8(AttDir) = False then
      begin
        MessageDlg(msg030, mtWarning, [mbOK], 0);
        Abort;
      end;
      myStringList := TStringList.Create;
      myStringList.Text := sqNotes.FieldByName('NotesAttName').AsString;
      for i := 0 to myStringList.Count - 1 do
        DeleteFileUTF8(AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip');
      myStringList.Free;
      if IsDirectoryEmpty(AttDir) = True then
        DeleteDirectory(AttDir, False);
    except
      MessageDlg(msg035, mtWarning, [mbOK], 0);
      Abort;
    end;
  // Delete images
  if DirectoryExistsUTF8(AttDir) = True then
  begin
    i := 0;
    try
      while FileExistsUTF8(AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg') = True do
      begin
        DeleteFileUTF8(AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
        Inc(i);
      end;
      if IsDirectoryEmpty(AttDir) = True then
        DeleteDirectory(AttDir, False);
    except
      MessageDlg(msg035, mtWarning, [mbOK], 0);
      Abort;
    end;
  end;
  // Save the UID of the record to be deleted for sync
  with sqDelRec do
  begin
    // DelRec must be opened and closed each time to read the modification
    // made by a sync operation by another possible session of mynotex.
    Open;
    Append;
    sqDelRec.FieldByName('DelRecUID').AsString :=
      sqNotes.FieldByName('NotesUID').AsString;
    sqDelRec.FieldByName('DelRecDTMod').AsDateTime := Now;
    Post;
    ApplyUpdates;
    Close;
  end;
  // Set the flag to recreate tags list
  if sqNotes.FieldByName('NotesTags').AsString <> '' then
    RecreateTagsList := True;
end;

procedure TfmMain.EditNotesDataset;
begin
  // Puts the notes dataset in Edit or in Insert
  // because the user has modified RichMemo
  if dsNotes.State in [dsBrowse] then
  begin
    if sqNotes.RecordCount > 0 then
    begin
      sqNotes.Edit;
    end
    else
      sqNotes.Insert;
  end;
  IsMemoModified := True;
end;

// ************************* FIND PROCEDURES **********************************

procedure TfmMain.sqFindAfterScroll(DataSet: TDataSet);
begin
  // Locate Subjects and Notes while scrolling on the grid
  LocateFromGrid;
end;

// *****************************************************************************
// ************************* MENU PROCEDURES *********************************
// *****************************************************************************

procedure TfmMain.miFileNewClick(Sender: TObject);
begin
  // Create a new file
  sdSaveDialog.Title := cpt001;
  sdSaveDialog.Filter := OpenSaveDlgFilter + '|*.mnt';
  sdSaveDialog.DefaultExt := '.mnt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute = True then
  begin
    // No show text only
    DisableShowTextOnly;
    CreateDataTables(sdSaveDialog.FileName);
    OpenDataTables(sdSaveDialog.FileName);
  end;
end;

procedure TfmMain.miFileOpenClick(Sender: TObject);
begin
  // Open a file
  odOpenDialog.Title := cpt022;
  odOpenDialog.Filter := cpt023;
  odOpenDialog.DefaultExt := '.mnt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute = True then
  begin
    // No show text only
    DisableShowTextOnly;
    OpenDataTables(odOpenDialog.FileName);
  end;
end;

procedure TfmMain.miFileCloseClick(Sender: TObject);
begin
  // Close the file
  SaveAllData;
  CloseDataTables;
end;

procedure TfmMain.miFileSaveClick(Sender: TObject);
begin
  // Save all data
  SaveAllData;
end;

procedure TfmMain.miFileUndoClick(Sender: TObject);
begin
  // Undo editing
  if MessageDlg(msg004, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end;
  if dsSubjects.State in [dsEdit, dsInsert] then
  begin
    sqSubjects.Cancel;
  end;
  if dsNotes.State in [dsEdit, dsInsert] then
  begin
    sqNotes.Cancel;
    // To reload the text of the note
    sqNotesAfterScroll(nil);
  end;
end;

procedure TfmMain.miFileUpdateClick(Sender: TObject);
var
  IDSub, IDNote: integer;
begin
  //Update data
  SaveAllData;
  IDSub := sqSubjects.FieldByName('IDSubjects').AsInteger;
  IDNote := sqNotes.FieldByName('IDNotes').AsInteger;
  sqSubjects.Close;
  // To avoid that the text is loaded
  IsTextToLoad := False;
  sqSubjects.Open;
  // sqNotes is closed and reopened in the sqSubjects.AfterScroll event
  // if there are subjects; otherwise nothing happens
  sqSubjects.Locate('IDSubjects', IDSub, []);
  sqNotes.Locate('IDNotes', IDNote, []);
  // Now load the text
  IsTextToLoad := True;
  // The locate function do not activate sqNotesAfterScroll
  // if the current record is the right one, so...
  sqNotesAfterScroll(nil);
end;

procedure TfmMain.miFileCopyAsClick(Sender: TObject);
var
  SearchRec: TSearchRec;
  AttOrigDir, AttDestDir: string;
begin
  // Copy current file
  sdSaveDialog.Title := cpt002;
  sdSaveDialog.Filter := OpenSaveDlgFilter + '|*.mnt';
  sdSaveDialog.DefaultExt := '.mnt';
  sdSaveDialog.FileName := '';
  if sdSaveDialog.Execute then
  begin
    CopyFile(sqSubjects.FileName, sdSaveDialog.FileName);
    AttOrigDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
    if DirectoryExistsUTF8(AttOrigDir) = True then
      try
        try
          AttDestDir := ExtractFileNameWithoutExt(sdSaveDialog.FileName);
          CreateDirUTF8(AttDestDir);
          // faSysFile (= normal file) to avoid that also the directory is found
          if FindFirst(AttOrigDir + DirectorySeparator + '*', faSysFile, SearchRec) = 0 then
            repeat
              CopyFile(AttOrigDir + DirectorySeparator + SearchRec.Name,
                AttDestDir + DirectorySeparator + SearchRec.Name);
            until
              FindNext(SearchRec) <> 0;
        except
          MessageDlg(msg037, mtWarning, [mbOK], 0);
        end;
      finally
        FindClose(SearchRec);
      end;
  end;
end;

procedure TfmMain.miFileImportClick(Sender: TObject);
begin
  // Import and export procedure;
  // it is called also by miFileExport and FileHTML menù
  // No show text only
  DisableShowTextOnly;
  if ((Sender = miFileImport) or (Sender = miFileExport)) then
  begin
    odOpenDialog.Title := cpt022;
    odOpenDialog.Filter := cpt023;
    odOpenDialog.DefaultExt := '.mnt';
    odOpenDialog.FileName := '';
    if odOpenDialog.Execute = False then
      Abort
    else if odOpenDialog.FileName = sqSubjects.FileName then
    begin
      MessageDlg(msg005, mtWarning, [mbOK], 0);
      Abort;
    end;
  end
  else if Sender = miFileHTML then
  begin
    sdSaveDialog.Title := cpt007;
    sdSaveDialog.Filter := cpt035;
    sdSaveDialog.DefaultExt := '.html';
    sdSaveDialog.FileName := '';
    if sdSaveDialog.Execute = False then
      Abort;
  end;
  with fmImpExp do
  begin
    if Sender = miFileImport then
    begin
      Caption := cpt003;
      lbSubjects.Caption := cpt004;
      bnImpExp.Caption := cpt005;
      bnClose.Caption := cpt018;
      cbNoExpDate.Caption := cpt041;
      cbNoExpDate.Visible := False;
      cbDeleteData.Caption := cpt006;
      cbDeleteData.Checked := False;
      cbDeleteData.Visible := True;
      cbSelDeselAll.Caption := cpt036;
      cbSelDeselAll.Checked := False;
      sqReadSubjects.FileName := odOpenDialog.FileName;
      sqReadNotes.FileName := odOpenDialog.FileName;
      sqReadDelRec.FileName := odOpenDialog.FileName;
      sqWriteSubjects.FileName := sqSubjects.FileName;
      sqWriteNotes.FileName := sqSubjects.FileName;
      sqWriteDelRec.FileName := sqSubjects.FileName;
    end
    else if Sender = miFileExport then
    begin
      Caption := cpt007;
      lbSubjects.Caption := cpt008;
      bnImpExp.Caption := cpt009;
      bnClose.Caption := cpt018;
      cbNoExpDate.Caption := cpt041;
      cbNoExpDate.Visible := False;
      cbDeleteData.Caption := cpt010;
      cbDeleteData.Checked := False;
      cbDeleteData.Visible := True;
      cbSelDeselAll.Caption := cpt036;
      cbSelDeselAll.Checked := False;
      sqReadSubjects.FileName := sqSubjects.FileName;
      sqReadNotes.FileName := sqSubjects.FileName;
      sqReadDelRec.FileName := sqSubjects.FileName;
      sqWriteSubjects.FileName := odOpenDialog.FileName;
      sqWriteNotes.FileName := odOpenDialog.FileName;
      sqWriteDelRec.FileName := odOpenDialog.FileName;
    end
    else if Sender = miFileHTML then
    begin
      Caption := cpt007;
      lbSubjects.Caption := cpt008;
      bnImpExp.Caption := cpt009;
      bnClose.Caption := cpt018;
      cbNoExpDate.Caption := cpt041;
      cbNoExpDate.Visible := True;
      cbDeleteData.Caption := cpt010;
      cbDeleteData.Checked := False;
      cbDeleteData.Visible := False;
      cbSelDeselAll.Caption := cpt036;
      cbSelDeselAll.Checked := False;
      sqReadSubjects.FileName := sqSubjects.FileName;
      sqReadNotes.FileName := sqSubjects.FileName;
    end;
    sqReadSubjects.TableName := 'Subjects';
    sqReadSubjects.PrimaryKey := 'IDSubjects';
    if miSubjectOrderTitle.Checked = True then
      sqReadSubjects.SQL :=
        'Select * from Subjects order by SubjectsName collate nocase, IDSubjects'
    else
      sqReadSubjects.SQL :=
        'Select * from Subjects order by SubjectsSort';
    sqReadNotes.TableName := 'Notes';
    sqReadNotes.PrimaryKey := 'IDNotes';
    if ((Sender = miFileImport) or (Sender = miFileExport)) then
    begin
      sqWriteSubjects.TableName := 'Subjects';
      sqWriteSubjects.PrimaryKey := 'IDSubjects';
      sqWriteNotes.TableName := 'Notes';
      sqWriteNotes.PrimaryKey := 'IDNotes';
      sqWriteDelRec.TableName := 'DelRec';
      sqWriteDelRec.PrimaryKey := 'IDDelRec';
      sqReadDelRec.TableName := 'DelRec';
      sqReadDelRec.PrimaryKey := 'IDDelRec';
    end;
    try
      sqReadSubjects.Open;
      if sqReadSubjects.RecordCount = 0 then
      begin
        if Sender = miFileImport then
          MessageDlg(msg060, mtWarning, [mbOK], 0)
        else
          MessageDlg(msg006, mtWarning, [mbOK], 0);
        sqReadSubjects.Close;
        Exit;
      end;
      // sqReadNotes will be opened later with the proper sql statement
      if ((Sender = miFileImport) or (Sender = miFileExport)) then
      begin
        sqWriteSubjects.Open;
        sqWriteNotes.Open;
        sqWriteDelRec.Open;
        sqReadDelRec.Open;
      end;
    except
      if ((Sender = miFileImport) or (Sender = miFileExport)) then
      begin
        sqReadSubjects.Close;
        sqWriteSubjects.Close;
        sqWriteNotes.Close;
        sqWriteDelRec.Close;
        sqReadDelRec.Close;
        MessageDlg(msg007, mtWarning, [mbOK], 0);
        Abort;
      end
      else
      begin
        sqReadSubjects.Close;
        MessageDlg(msg025, mtWarning, [mbOK], 0);
        Abort;
      end;
    end;
    cbReadSubjects.Items.Clear;
    while not sqReadSubjects.EOF do
    begin
      cbReadSubjects.Items.Add(sqReadSubjects.FieldByName('SubjectsName').AsString);
      sqReadSubjects.Next;
    end;
    sqReadSubjects.First;
    ShowModal;
  end;
end;

procedure TfmMain.miFileZimClick(Sender: TObject);
var
  sqZimSubjects, sqZimNotes: TSqlite3Dataset;
  slTextPage, slFilesList, slAttList: TStringList;
  stOutputDir: string;
  i: integer;
begin
  if MessageDlg(msg083, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Exit;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    slFilesList := TStringList.Create;
    slTextPage := TStringList.Create;
    slAttList := TStringList.Create;
    sqZimSubjects := TSqlite3Dataset.Create(Self);
    sqZimNotes := TSqlite3Dataset.Create(Self);
    stOutputDir := SpaceToLine(ExtractFileNameWithoutExt(sqSubjects.FileName) + '_Zim');
    CreateDirUTF8(stOutputDir);
    if DirectoryExistsUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName)) = True then
    begin
      CreateDirUTF8(stOutputDir + DirectorySeparator + 'ExportZimFiles');
      slFilesList := FindAllFiles(ExtractFileNameWithoutExt(sqSubjects.FileName) + DirectorySeparator, '*.jpeg;*.jpg;*.png;*.zip', True);
      for i := 0 to slFilesList.Count - 1 do
      begin
        CopyFile(slFilesList[i],
          stOutputDir + DirectorySeparator + 'ExportZimFiles' + DirectorySeparator + ExtractFileName(slFilesList[i]));
      end;
    end;
    sqZimSubjects.FileName := sqSubjects.FileName;
    sqZimSubjects.TableName := 'Subjects';
    sqZimSubjects.PrimaryKey := 'IDSubjects';
    sqZimSubjects.SQL := 'Select * from Subjects';
    sqZimSubjects.Open;
    sqZimNotes.FileName := sqNotes.FileName;
    sqZimNotes.TableName := 'Notes';
    sqZimNotes.PrimaryKey := 'IDNotes';
    try
      while not sqZimSubjects.EOF do
      begin
        slTextPage.Clear;
        slTextPage.Add('Content-Type: text/x-zim-wiki');
        slTextPage.Add('Wiki-Format: zim 0.4');
        slTextPage.Add('Creation-Date: ' + FormatDateTime('yyyy-mm-dd"T"hh:nn:ss', sqZimSubjects.FieldByName('SubjectsDTMod').AsDateTime));
        slTextPage.Add('');
        slTextPage.Add('====== ' + sqZimSubjects.FieldByName('SubjectsName').AsString + ' ======');
        slTextPage.Add('Created ' + FormatDateTime('dddd dd mmmm yyyy', sqZimSubjects.FieldByName('SubjectsDTMod').AsDateTime));
        slTextPage.Add('');
        slTextPage.Add(sqZimSubjects.FieldByName('SubjectsComments').AsString);
        slTextPage.SaveToFile(SpaceToLine(stOutputDir + DirectorySeparator + sqZimSubjects.FieldByName('SubjectsName').AsString + '.txt'));
        CreateDirUTF8(SpaceToLine(stOutputDir + DirectorySeparator + sqZimSubjects.FieldByName('SubjectsName').AsString));
        sqZimNotes.SQL := 'Select * from Notes where ' + 'Notes.ID_Subjects = ' + sqZimSubjects.FieldByName('IDSubjects').AsString;
        sqZimNotes.Open;
        while not sqZimNotes.EOF do
        begin
          ;
          slTextPage.Clear;
          slTextPage.Add('Content-Type: text/x-zim-wiki');
          slTextPage.Add('Wiki-Format: zim 0.4');
          slTextPage.Add('Creation-Date: ' + FormatDateTime('yyyy-mm-dd"T"hh:nn:ss', sqZimNotes.FieldByName('NotesDTMod').AsDateTime));
          slTextPage.Add('');
          slTextPage.Add('====== ' + sqZimNotes.FieldByName('NotesTitle').AsString + ' ======');
          slTextPage.Add('Created ' + FormatDateTime('dddd dd mmmm yyyy', sqZimNotes.FieldByName('NotesDate').AsDateTime));
          if sqZimNotes.FieldByName('NotesAttName').AsString <> '' then
          begin
            slTextPage.Add('');
            slAttList.Text := sqZimNotes.FieldByName('NotesAttName').AsString;
            for i := 0 to slAttList.Count - 1 do
            begin
              slAttList[i] := ExtractFileNameWithoutExt(slAttList[i]) + '.zip';
              slTextPage.Add('[[..' + DirectorySeparator + '..' + DirectorySeparator + 'ExportZimFiles' + DirectorySeparator + sqZimNotes.FieldByName('NotesUID').AsString + '-' + slAttList[i] + '|' + slAttList[i] + ']]');
            end;
          end;
          slTextPage.Add('');
          slTextPage.Add(ExpTextToZim(sqZimNotes.FieldByName('NotesText').AsString));
          slTextPage.SaveToFile(SpaceToLine(stOutputDir + DirectorySeparator + sqZimSubjects.FieldByName('SubjectsName').AsString + DirectorySeparator + sqZimNotes.FieldByName('NotesTitle').AsString + '.txt'));
          sqZimNotes.Next;
        end;
        sqZimNotes.Close;
        sqZimSubjects.Next;
        Application.ProcessMessages;
      end;
      MessageDlg(msg024, mtInformation, [mbOK], 0);
    except
      MessageDlg(msg084, mtWarning, [mbOK], 0);
    end;
  finally
    slTextPage.Free;
    slFilesList.Free;
    slAttList.Free;
    sqZimSubjects.Free;
    sqZimNotes.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miConvertTomboyClick(Sender: TObject);
var
  DataPath, AppName: string;
begin
  // Convert from Tomboy and GNote
  SaveAllData;
  if Sender = miConvertTomboy then
  begin
    AppName := 'Tomboy';
    DataPath := GetEnvironmentVariable('HOME') + DirectorySeparator + '.local/share/tomboy';
  end
  else if Sender = miConvertGNote then
  begin
    AppName := 'GNote';
    DataPath := GetEnvironmentVariable('HOME') + DirectorySeparator + '.local/share/gnote';
  end;
  if MessageDlg(msg049 + ' ' + AppName + '?', mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  ConvertFromTomboyGnote(DataPath);
end;

procedure TfmMain.miFilePrinterSetupClick(Sender: TObject);
begin
  // Printer setup
  pdPrintDialog.Execute;
end;

procedure TfmMain.miFileOpenLast1Click(Sender: TObject);
begin
  // Open last1 file
  // No show text only
  DisableShowTextOnly;
  if FileExistsUTF8(LastDatabase1) then
  begin
    OpenDataTables(LastDatabase1);
  end
  else
  begin
    MessageDlg(msg008, mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileOpenLast2Click(Sender: TObject);
begin
  // Open last2 file
  // No show text only
  DisableShowTextOnly;
  if FileExistsUTF8(LastDatabase2) then
  begin
    OpenDataTables(LastDatabase2);
  end
  else
  begin
    MessageDlg(msg008, mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileOpenLast3Click(Sender: TObject);
begin
  // Open last3 file
  // No show text only
  DisableShowTextOnly;
  if FileExistsUTF8(LastDatabase3) then
  begin
    OpenDataTables(LastDatabase3);
  end
  else
  begin
    MessageDlg(msg008, mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileOpenLast4Click(Sender: TObject);
begin
  // Open last4 file
  // No show text only
  DisableShowTextOnly;
  if FileExistsUTF8(LastDatabase4) then
  begin
    OpenDataTables(LastDatabase4);
  end
  else
  begin
    MessageDlg(msg008, mtWarning, [mbOK], 0);
  end;
end;

procedure TfmMain.miFileExitClick(Sender: TObject);
begin
  // Exit
  SaveAllData;
  Close;
end;

procedure TfmMain.miSubjectNewClick(Sender: TObject);
begin
  // Create a new subject
  // No show text only
  DisableShowTextOnly;
  sqSubjects.Append;
  // The record is saved in AfterInsert event
end;

procedure TfmMain.miSubjectDeleteClick(Sender: TObject);
begin
  // Delete a subject and the linked notes
  // ApplyUpdates is in sqSubjects.AfterDelete, so that it can work also
  // for deletion with Ctrl+Canc from the grid
  // No show text only
  DisableShowTextOnly;
  sqSubjects.Delete;
end;

procedure TfmMain.miSubjectCommentsClick(Sender: TObject);
begin
  // Add comments to subjects
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  // In this case, the menu item should not be enabled; anyway...
  if sqSubjects.RecordCount = 0 then
    Abort;
  fmCommentsSubjects.Caption := cpt019;
  fmCommentsSubjects.lbSubName.Caption := cpt020;
  fmCommentsSubjects.lbComments.Caption := cpt021;
  fmCommentsSubjects.bnSubCommCancel.Caption := cpt018;
  fmCommentsSubjects.ShowModal;
end;

procedure TfmMain.miSubjectLookClick(Sender: TObject);
begin
  // Open look form
  // Add comments to subjects
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  // In this case, the menu item should not be enabled; anyway...
  if sqSubjects.RecordCount = 0 then
    Abort;
  // Add the translation captions...
  fmLook.flSubjectForm := True;
  fmLook.Caption := cpt056;
  fmLook.lbColBack.Caption := cpt057;
  fmLook.lbColFont.Caption := cpt058;
  fmLook.bnColDef1.Caption := cpt059;
  fmLook.bnColDef2.Caption := cpt060;
  fmLook.bnColDef3.Caption := cpt061;
  fmLook.bnNoCol.Caption := cpt062;
  fmLook.lbColLorem1.Caption := cpt063;
  fmLook.lbColLorem2.Caption := cpt064;
  fmLook.lbColLorem3.Caption := cpt065;
  fmLook.bnColCancel.Caption := cpt018;
  fmLook.ShowModal;
end;

procedure TfmMain.miSubjectOrderTitleClick(Sender: TObject);
var
  iIDSub: integer;
begin
  // Change order of notes; this events is triggered also
  // by miSubjectOrderCustom
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  with sqSubjects do
  begin
    iIDSub := FieldByName('IDSubjects').AsInteger;
    Close;
    if miSubjectOrderTitle.Checked = True then
      SQL := 'Select * from Subjects order by SubjectsName collate nocase, IDSubjects'
    else
      SQL := 'Select * from Subjects order by SubjectsSort';
    Open;
    Locate('IDSubjects', iIDSub, []);
  end;
  bnSubUp.Enabled := miSubjectOrderCustom.Checked;
  bnSubDown.Enabled := miSubjectOrderCustom.Checked;
end;

procedure TfmMain.miNotesNewClick(Sender: TObject);
begin
  // New note
  // No show text only
  DisableShowTextOnly;
  sqNotes.Append;
  dbTitle.SetFocus;
end;

procedure TfmMain.miNotesDeleteClick(Sender: TObject);
begin
  // Delete note
  // No show text only
  DisableShowTextOnly;
  if MessageDlg(msg009, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end
  else
  begin
    sqNotes.Delete;
    sqNotes.ApplyUpdates;
  end;
  // Recreate tags list if necessary
  if RecreateTagsList = True then
  begin
    CreateTagsList;
    RecreateTagsList := False;
    OldTagsValue := '';
  end;
end;

procedure TfmMain.miNotesUndoClick(Sender: TObject);
begin
  // Replace note text with the stored one
  GetUndoData;
end;

procedure TfmMain.miOrderByDateClick(Sender: TObject);
begin
  // Change order of notes; this events is triggered also
  // by miOrderByTitle and miOrderCustom
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  with sqNotes do
  begin
    Close;
    if miOrderByTitle.Checked = True then
    begin
      SQL := 'Select * from Notes where ID_Subjects = ' + sqSubjects.FieldByName('IDSubjects').AsString + ' order by NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      SQL := 'Select * from Notes where Notes.ID_Subjects = ' + sqSubjects.FieldByName('IDSubjects').AsString + ' order by NotesDate, IDNotes';
    end
    else
    begin
      SQL := 'Select * from Notes where Notes.ID_Subjects = ' + sqSubjects.FieldByName('IDSubjects').AsString + ' order by NotesSort';
    end;
    // To show the last note if order by note is active
    if miOrderByDate.Checked = True then
    begin
      // sqNotes.DisableControls does not work, so...
      IsTextToLoad := False;
      Open;
      IsTextToLoad := True;
      // Call to last with recordcount = 0 raises an error
      if RecordCount > 0 then
        Last;
    end
    else
    begin
      Open;
    end;
  end;
  // This is necessary to change immediately the order of the notes
  // without scrolling the subjects
  miFileUpdateClick(nil);
  bnNotesUp.Enabled := miOrderCustom.Checked;
  bnNotesDown.Enabled := miOrderCustom.Checked;
end;

procedure TfmMain.miNotesEncDecryptClick(Sender: TObject);
begin
  // Encrypt and decrypt text
  SaveAllData;
  if pnPassword.Visible = True then
  begin
    MessageDlg(msg042, mtInformation, [mbOK], 0);
    Abort;
  end;
  // The text must be encrypted
  if sqNotes.FieldByName('NotesCheckPwd').AsString = '' then
  begin
    with fmEncryption do
    begin
      Caption := fmMain.cpt030;
      lbEncrypt.Caption := fmMain.cpt031;
      cbShowChar.Caption := fmMain.cpt032;
      bnClose.Caption := fmMain.cpt018;
      bnEncrypt.Caption := fmMain.cpt030;
      edPwd1.Clear;
      edPwd2.Clear;
      ShowModal;
    end;
  end
  // The text must be decrypted
  else
  begin
    if MessageDlg(msg041, mtConfirmation, [mbOK, mbCancel], 0) = mrOk then
      try
        Screen.Cursor := crHourGlass;
        Application.ProcessMessages;
        sqNotes.Edit;
        dcAES.InitStr(edPassword.Text, TDCP_sha1);
        sqNotes.FieldByName('NotesText').AsWideString :=
          dcAES.DecryptString(sqNotes.FieldByName('NotesText').AsWideString);
        fmMain.sqNotes.FieldByName('NotesCheckPwd').AsString := '';
        fmMain.sqNotes.Post;
        fmMain.sqNotes.ApplyUpdates;
      finally
        Screen.Cursor := crDefault;
      end;
  end;
end;

procedure TfmMain.miNotesImagesClick(Sender: TObject);
begin
  // Add a picture
  opOpenPictureDlg.Title := cpt037;
  if opOpenPictureDlg.Execute = True then
    AddImage(opOpenPictureDlg.FileName);
end;

procedure TfmMain.miNotesMoveClick(Sender: TObject);
var
  OrigSQL: string;
begin
  // Move a note
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  // In this case, the menu item should not be enabled; anyway...
  if sqNotes.RecordCount = 0 then
    Abort;
  fmMoveNote.Caption := cpt016;
  fmMoveNote.lbMoveNoteSub.Caption := cpt017;
  fmMoveNote.bnMoveNote.Caption := cpt016;
  fmMoveNote.bnCloseNote.Caption := cpt018;
  fmMoveNote.sqMoveSubjects.FileName :=
    sqSubjects.FileName;
  fmMoveNote.sqMoveSubjects.TableName := 'Subjects';
  fmMoveNote.sqMoveSubjects.PrimaryKey := 'IDSubjects';
  OrigSQL := sqSubjects.SQL;
  // Current subject is excluded
  OrigSQL := StringReplace(OrigSQL, 'order by', 'where IDSubjects <> ' + IntToStr(sqSubjects.FieldByName('IDSubjects').AsInteger) + ' order by', [rfIgnoreCase]);
  fmMoveNote.sqMoveSubjects.SQL := OrigSQL;
  fmMoveNote.sqMoveSubjects.Open;
  if fmMoveNote.sqMoveSubjects.RecordCount = 0 then
  begin
    fmMoveNote.sqMoveSubjects.Close;
    MessageDlg(msg043, mtInformation, [mbOK], 0);
    Exit;
  end
  else
  begin
    fmMoveNote.ShowModal;
  end;
end;

procedure TfmMain.miNotesSendToOOClick(Sender: TObject);
begin
  // Send to OpenOffice.org Writer
  tbOpenNote.Tag := 0;
  tbOpenNoteClick(nil);
end;

procedure TfmMain.miNotesSendToLOClick(Sender: TObject);
begin
  // Send to LibreOffice Writer
  tbOpenNote.Tag := 1;
  tbOpenNoteClick(nil);
end;

procedure TfmMain.miNotesSendToBrowserClick(Sender: TObject);
begin
  // Send to browser
  tbOpenNote.Tag := 2;
  tbOpenNoteClick(nil);
end;

procedure TfmMain.miNotesInsertClick(Sender: TObject);
var
  myZip: TUnZipper;
  myList, stXML, stFileOrig, stStyleSheet: TStringList;
  i, idxXML: integer;
  stOutput, stOutputWithTags, ssName, ssCodes: string;
  flCopy: boolean;
  Fp: TFontParams;
begin
  // Insert file in a new note
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  odOpenDialog.Title := cpt024;
  odOpenDialog.Filter := cpt025;
  odOpenDialog.DefaultExt := '.odt';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute = False then
    Abort;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  // Open a text file
  if UpperCase(ExtractFileExt(odOpenDialog.FileName)) = '.TXT' then
    try
      try
        myList := TStringList.Create;
        myList.LoadFromFile(odOpenDialog.FileName);
        stOutput := myList.Text;
      except
        MessageDlg(msg028, mtWarning, [mbOK], 0);
        Screen.Cursor := crDefault;
        Abort;
      end;
    finally
      myList.Free;
    end
  // Open Writer file and extract content.xml
  else if UpperCase(ExtractFileExt(odOpenDialog.FileName)) = '.ODT' then
  begin
    myZip := TUnZipper.Create;
    myList := TStringList.Create;
    stXML := TStringList.Create;
    stFileOrig := TStringList.Create;
    myList.Add('content.xml');
    myZip.OutputPath := GetTempDir;
    myZip.FileName := odOpenDialog.FileName;
    try
      try
        myZip.UnZipFiles(myList);
        // GetTempDir is /tmp
        stFileOrig.LoadFromFile(GetTempDir + DirectorySeparator + 'content.xml');
        DeleteFile(GetTempDir + DirectorySeparator + 'content.xml');
      except
        MessageDlg(msg028, mtWarning, [mbOK], 0);
        Screen.Cursor := crDefault;
        Abort;
      end;
    finally
      myZip.Free;
      myList.Free;
    end;
    //Parse XML file
    // Select only styles section
    stXML.Text := Copy(stFileOrig.Text, Pos('<office:automatic-styles>', stFileOrig.Text), Pos('</office:automatic-styles>', stFileOrig.Text) - Pos('<office:automatic-styles>', stFileOrig.Text) + Length('</office:automatic-styles>'));
    stStyleSheet := TStringList.Create;
    // Get paragraph stiles
    for i := 1 to 10000 do
    begin
      idxXML := Pos('<style:style style:name="P' + IntToStr(i) + '" style:family="paragraph"', stXML.Text);
      if idxXML > 0 then
      begin
        stStyleSheet.Add('P' + IntToStr(i) + '-');
        stXML.Text := Copy(stXML.Text, idxXML, Length(stXML.Text) - idxXML);
        // Bold
        if Pos('bold', stXML.Text) > 0 then
          if Pos('bold', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 'b';
        // Italic
        if Pos('italic', stXML.Text) > 0 then
          if Pos('italic', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 'i';
        // Underline
        if Pos('text-underline-style="solid"', stXML.Text) > 0 then
          if Pos('text-underline-style="solid"', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 'u';
        // Strikethrough
        if Pos('text-line-through-style="solid"', stXML.Text) > 0 then
          if Pos('text-line-through-style="solid"', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 's';
      end
      else
      begin
        Break;
      end;
    end;
    // Get text stiles
    for i := 1 to 10000 do
    begin
      idxXML := Pos('<style:style style:name="T' + IntToStr(i) + '" style:family="text">', stXML.Text);
      if idxXML > 0 then
      begin
        stStyleSheet.Add('T' + IntToStr(i) + '-');
        stXML.Text := Copy(stXML.Text, idxXML, Length(stXML.Text) - idxXML);
        // Bold
        if Pos('bold', stXML.Text) > 0 then
          if Pos('bold', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 'b';
        // Italic
        if Pos('italic', stXML.Text) > 0 then
          if Pos('italic', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 'i';
        // Underline
        if Pos('text-underline-style="solid"', stXML.Text) > 0 then
          if Pos('text-underline-style="solid"', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 'u';
        // Strikethrough
        if Pos('text-line-through-style="solid"', stXML.Text) > 0 then
          if Pos('text-line-through-style="solid"', stXML.Text) < Pos('/>', stXML.Text) then
            stStyleSheet.Strings[stStyleSheet.Count - 1] :=
              stStyleSheet.Strings[stStyleSheet.Count - 1] + 's';
      end
      else
      begin
        Break;
      end;
    end;
    // Select only text section
    // #3 will be replaced with < and #4 with >
    stXML.Text := Copy(stFileOrig.Text, Pos('<office:body>', stFileOrig.Text), Pos('</office:body>', stFileOrig.Text) - Pos('<office:body>', stFileOrig.Text) + Length('</office:body>'));
    for i := 0 to stStyleSheet.Count - 1 do
    begin
      ssName := Copy(stStyleSheet.Strings[i], 0, Pos('-', stStyleSheet.Strings[i]) - 1);
      ssCodes := Copy(stStyleSheet.Strings[i], Pos('-', stStyleSheet.Strings[i]) + 1, Length(stStyleSheet.Strings[i]) - Pos('-', stStyleSheet.Strings[i]));
      if ssCodes <> '' then
      begin
        // Change paragraph style
        while Pos('<text:p text:style-name="' + ssName + '">', stXML.Text) > 0 do
        begin
          // change to <*******> all the </text:p> *before* the string that will be changed
          while ((Pos('</text:p>', stXML.Text) > 0) and (Pos('</text:p>', stXML.Text) < Pos('<text:p text:style-name="' + ssName + '">', stXML.Text))) do
            stXML.Text := StringReplace(stXML.Text, '</text:p>', '<*******>', []);
          // change the string
          stXML.Text := StringReplace(stXML.Text, '<text:p text:style-name="' + ssName + '">', #3 + ssCodes + #4, []);
          // the first </text:p> is the one just after the replaced string
          stXML.Text := StringReplace(stXML.Text, '</text:p>', #3 + '/' + ssCodes + #4 + LineEnding, []);
        end;
        // Restore then </text:p> tag
        stXML.Text := StringReplace(stXML.Text, '<*******>', '</text:p>', [rfReplaceAll]);
        // Change text style
        while Pos('<text:span text:style-name="' + ssName + '">', stXML.Text) > 0 do
        begin
          // change to <*******> all the </text:span> *before* the string that will be changed
          while ((Pos('</text:span>', stXML.Text) > 0) and (Pos('</text:span>', stXML.Text) < Pos('<text:span text:style-name="' + ssName + '">', stXML.Text))) do
            stXML.Text := StringReplace(stXML.Text, '</text:span>', '<*******>', []);
          // change the string
          stXML.Text := StringReplace(stXML.Text, '<text:span text:style-name="' + ssName + '">', #3 + ssCodes + #4, []);
          // the first </text:span> is the one just after the replaced string
          stXML.Text := StringReplace(stXML.Text, '</text:span>', #3 + '/' + ssCodes + #4, []);
        end;
        // restore then </text:span> tag
        stXML.Text := StringReplace(stXML.Text, '<*******>', '</text:span>', [rfReplaceAll]);
      end;
    end;
    stStyleSheet.Free;
    stFileOrig.Free;

    // Replace the tags for footnotes
    stXML.Text := StringReplace(stXML.Text, '<text:note ', ' [<', [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '</text:note>', ']', [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '</text:note-citation>', '. ', [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '</text:p></text:note-body>', '', [rfReplaceAll]);
    // Replace the tags which implies a CR
    stXML.Text := StringReplace(stXML.Text, 'text:name="Illustration"/>', '>', []);
    stXML.Text := StringReplace(stXML.Text, 'text:name="Table"/>', '>', []);
    stXML.Text := StringReplace(stXML.Text, 'text:name="Text"/>', '>', []);
    stXML.Text := StringReplace(stXML.Text, 'text:name="Drawing"/>', '>', []);
    // Empty paragraph
    stXML.Text := StringReplace(stXML.Text, '"/>', '>' + LineEnding, [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '<text:h', LineEnding + '<', [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '</text:h>', LineEnding + LineEnding, [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '<text:line-break/>', LineEnding, [rfReplaceAll]);
    stXML.Text := StringReplace(stXML.Text, '</text:p>', LineEnding, [rfReplaceAll]);

    // Clear other HTML tags
    // Following cycle for...next is much slower with a Tlist than a String: so...
    stOutputWithTags := stXML.Text;
    stXML.Free;
    flCopy := True;
    for i := 1 to Length(stOutputWithTags) do
    begin
      if stOutputWithTags[i] = '<' then
        flCopy := False
      else if stOutputWithTags[i] = '>' then
        flCopy := True
      else if flCopy = True then
        stOutput := stOutput + stOutputWithTags[i];
      Application.ProcessMessages;
    end;
    // Delete possibile CR at the beginning
    if Copy(stOutput, 0, 1) = LineEnding then
      stOutput := Copy(stOutput, 2, Length(stOutput) - 2);
    // Restore ‹ and › char (in HTML are &lt; and &gt;), not as < and >
    // to avoid confusion with html tags
    stOutput := StringReplace(stOutput, '&lt;', '‹', [rfReplaceAll]);
    stOutput := StringReplace(stOutput, '&gt;', '›', [rfReplaceAll]);
    stOutput := StringReplace(stOutput, '&apos;', '''', [rfReplaceAll]);
    // Replace #3 with < and #4 with >; they are html tags from Writer text
    stOutput := StringReplace(stOutput, #3, '<', [rfReplaceAll]);
    stOutput := StringReplace(stOutput, #4, '>', [rfReplaceAll]);
    stOutput := StringReplace(stOutput, '<s>', '<strike>', [rfReplaceAll]);
    stOutput := StringReplace(stOutput, '</s>', '</strike>', [rfReplaceAll]);
  end;
  // Get the default font
  Fp.Color := clBlack;
  Fp.Name := DefFontName;
  Fp.Size := DefFontSize;
  Fp.Style := [];
  // Save text
  sqNotes.Append;
  // After insert there is a post, so...
  sqNotes.Edit;
  sqNotes.FieldByName('NotesTitle').AsString :=
    ExtractFileNameOnly(odOpenDialog.FileName);
  sqNotes.FieldByName('NotesText').AsString :=
    '<font color="#000000"' + ' size="' + IntToStr(Fp.Size) + '"' + ' face="' + Fp.Name + '">' + StOutput + '</font>';
  sqNotes.Post;
  sqNotes.ApplyUpdates;
  LoadRichMemo;
  Screen.Cursor := crDefault;
end;

procedure TfmMain.miAttachNewClick(Sender: TObject);
begin
  // Insert attachment relate with the current note
  // No show text only
  DisableShowTextOnly;
  odOpenDialog.Title := cpt026;
  odOpenDialog.Filter := cpt027;
  odOpenDialog.DefaultExt := '';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute = False then
    Abort;
  CreateAttachment(odOpenDialog.FileName);
end;

procedure TfmMain.miAttachOpenClick(Sender: TObject);
var
  myUnZip: TUnZipper;
  myList: TStringList;
  AttDir, OrigFileName, OutDir: string;
begin
  // Open and Save As attachment (miAttachOpen and miAttachSaveAs)
  // No show text only
  DisableShowTextOnly;
  AttDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
  if DirectoryExistsUTF8(AttDir) = False then
  begin
    MessageDlg(msg030, mtWarning, [mbOK], 0);
    Abort;
  end;
  if lbAttNames.ItemIndex < 0 then
  begin
    MessageDlg(msg031, mtWarning, [mbOK], 0);
    Abort;
  end;
  if ((Sender = miAttachSaveAs) or (Sender = pmAttSaveAs)) then
  begin
    // Title of the TSelectDirectoryDialog do not work, up to now, anyway...
    sdSelDirDialog.Title := cpt054;
    if sdSelDirDialog.Execute then
      OutDir := sdSelDirDialog.FileName
    else
      Abort;
  end
  // The other Sender could be miAttachOpen or lbAttNames (double clic)
  else
  begin
    OutDir := GetTempDir;
  end;
  try
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    try
      myUnZip := TUnZipper.Create;
      myList := TStringList.Create;
      AttDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
      OrigFileName := lbAttNames.Items.ValueFromIndex[lbAttNames.ItemIndex];
      myList.Add(OrigFileName);
      myUnZip.OutputPath := OutDir;
      myUnZip.FileName := AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(OrigFileName) + '.zip';
      myUnZip.UnZipFiles(myList);
      if ((Sender <> miAttachSaveAs) and (Sender <> pmAttSaveAs)) then
        OpenDocument(OutDir + DirectorySeparator + OrigFileName);
    except
      MessageDlg(msg032, mtWarning, [mbOK], 0);
    end;
  finally
    myUnZip.Free;
    myList.Free;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miAttachDeleteClick(Sender: TObject);
var
  AttDir, OrigFileName: string;
begin
  // Delete attachment
  // No show text only
  DisableShowTextOnly;
  AttDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
  if DirectoryExistsUTF8(AttDir) = False then
  begin
    if MessageDlg(msg030, mtWarning, [mbOK, mbCancel], 0) = mrOk then
    begin
      lbAttNames.Items.Delete(lbAttNames.ItemIndex);
      sqNotes.Edit;
      sqNotes.FieldByName('NotesAttName').AsString := lbAttNames.Items.Text;
      sqNotes.Post;
      sqNotes.ApplyUpdates;
    end;
    Abort;
  end;
  if lbAttNames.ItemIndex < 0 then
  begin
    MessageDlg(msg031, mtWarning, [mbOK], 0);
    Abort;
  end;
  if MessageDlg(msg034, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  try
    OrigFileName := lbAttNames.Items.ValueFromIndex[lbAttNames.ItemIndex];
    DeleteFileUTF8(AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(OrigFileName) + '.zip');
    lbAttNames.Items.Delete(lbAttNames.ItemIndex);
    sqNotes.Edit;
    sqNotes.FieldByName('NotesAttName').AsString := lbAttNames.Items.Text;
    sqNotes.Post;
    sqNotes.ApplyUpdates;
    if IsDirectoryEmpty(AttDir) = True then
      DeleteDirectory(AttDir, False);
    lbListAttach.Caption := stAttachments + ' [' + IntToStr(lbAttNames.Items.Count) + ']';
  except
    MessageDlg(msg035, mtWarning, [mbOK], 0);
    lbListAttach.Caption := stAttachments + ' [' + IntToStr(lbAttNames.Items.Count) + ']';
  end;
end;

procedure TfmMain.miTagsRenameClick(Sender: TObject);
var
  stMessage, stTag, OldValue, NewValue: string;
  sqTags: TSqlite3Dataset;
  IDNote: integer;
  ResultInput: boolean;
begin
  //Rename or remove a tag in the archive in use
  // Event also for miTagsRemove
  SaveAllData;
  if Sender = miTagsRename then
    stMessage := msg044
  else
    stMessage := msg045;
  if MessageDlg(stMessage, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Exit;
  if lbTagsNames.ItemIndex > -1 then
    stTag := UTF8Copy(lbTagsNames.Items[lbTagsNames.ItemIndex], 1, UTF8Pos('[', lbTagsNames.Items[lbTagsNames.ItemIndex]) - 2)
  else
    stTag := '';
  ResultInput := InputQuery(msg046, msg047, stTag);
  if ResultInput = False then
    Exit;
  OldValue := stTag;
  stTag := '';
  if Sender = miTagsRename then
  begin
    ResultInput := InputQuery(msg046, msg048, stTag);
    if ResultInput = False then
      Exit;
    NewValue := stTag;
  end;
  try
    Screen.Cursor := crHourGlass;
    sqTags := TSqlite3Dataset.Create(Self);
    sqTags.FileName := sqNotes.FileName;
    sqTags.TableName := 'Notes';
    sqTags.PrimaryKey := 'IDNotes';
    sqTags.SQL := 'Select IDNotes, NotesTags from Notes where NotesTags not null';
    sqTags.Open;
    while not sqTags.EOF do
    begin
      if ((sqTags.FieldByName('NotesTags').AsString = OldValue) or (UTF8Pos(OldValue + ', ', sqTags.FieldByName('NotesTags').AsString) > 0) or (UTF8Pos(', ' + OldValue, sqTags.FieldByName('NotesTags').AsString) > 0)) then
      begin
        sqTags.Edit;
        if Sender = miTagsRename then
          sqTags.FieldByName('NotesTags').AsString :=
            StringReplace(sqTags.FieldByName('NotesTags').AsString, OldValue, NewValue, [])
        else if Sender = miTagsRemove then
        begin
          sqTags.FieldByName('NotesTags').AsString :=
            StringReplace(sqTags.FieldByName('NotesTags').AsString, OldValue + ', ', '', []);
          sqTags.FieldByName('NotesTags').AsString :=
            StringReplace(sqTags.FieldByName('NotesTags').AsString, ', ' + OldValue, '', []);
          sqTags.FieldByName('NotesTags').AsString :=
            StringReplace(sqTags.FieldByName('NotesTags').AsString, OldValue, '', []);
        end;
        sqTags.Post;
        sqTags.ApplyUpdates;
      end;
      Application.ProcessMessages;
      sqTags.Next;
    end;
    if sqNotes.RecordCount > 0 then
    begin
      IDNote := sqNotes.FieldByName('IDNotes').AsInteger;
      sqNotes.Close;
      IsTextToLoad := False;
      sqNotes.Open;
      sqNotes.Locate('IDNotes', IDNote, []);
      IsTextToLoad := True;
      sqNotesAfterScroll(nil);
      CreateTagsList;
    end;
  finally
    Screen.Cursor := crDefault;
    sqTags.Free;
  end;
end;

procedure TfmMain.miNotesLookClick(Sender: TObject);
begin
  // Open look form
  // Add comments to subjects
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  // In this case, the menu item should not be enabled; anyway...
  if sqNotes.RecordCount = 0 then
    Abort;
  // Add the translation captions...
  fmLook.flSubjectForm := False;
  fmLook.Caption := cpt056;
  fmLook.lbColBack.Caption := cpt057;
  fmLook.lbColFont.Caption := cpt058;
  fmLook.bnColDef1.Caption := cpt059;
  fmLook.bnColDef2.Caption := cpt060;
  fmLook.bnColDef3.Caption := cpt061;
  fmLook.bnNoCol.Caption := cpt062;
  fmLook.lbColLorem1.Caption := cpt063;
  fmLook.lbColLorem2.Caption := cpt064;
  fmLook.lbColLorem3.Caption := cpt065;
  fmLook.bnColCancel.Caption := cpt018;
  fmLook.ShowModal;
end;

procedure TfmMain.miNotesFindClick(Sender: TObject);
begin
  // Open Find notes panel
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  if miNotesShowActivities.Checked = True then
  begin
    miNotesShowActivitiesClick(nil);
  end;
  miNotesFind.Checked := not miNotesFind.Checked;
  tbFind.Down := miNotesFind.Checked;
  if miNotesFind.Checked then
  begin
    // No only text on search
    miNotesShowOnlyText.Enabled := False;
    // No change order during search
    miNotesOrderBy.Enabled := False;
    miOrderByDate.Enabled := False;
    miOrderByTitle.Enabled := False;
    miOrderCustom.Enabled := False;
    pnFind.Visible := True;
    pnFindText.Visible := True;
    // spSplitterFind must be made visible after pnFind to stay on top
    spSplitterFind.Visible := True;
    edLocateText.Clear;
    edFindText.Clear;
    edFindText.SetFocus;
  end
  else
  begin
    miNotesShowOnlyText.Enabled := True;
    miNotesOrderBy.Enabled := True;
    miOrderByDate.Enabled := True;
    miOrderByTitle.Enabled := True;
    miOrderCustom.Enabled := True;
    spSplitterFind.Visible := False;
    pnFind.Visible := False;
    pnFindText.Visible := False;
    lbFound.Caption := cpt011;
    sqFind.Close;
    bnFind2.Enabled := False;
    meSearchCond.Clear;
  end;
end;

procedure TfmMain.miNotesPrintClick(Sender: TObject);
begin
  // Print the current note
  SaveAllData;
  if MessageDlg(msg062, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  tbOpenNoteClick(miNotesPrint);
end;

procedure TfmMain.miNotesShowOnlyTextClick(Sender: TObject);
begin
  // Show only memo text
  miNotesShowOnlyText.Checked := not miNotesShowOnlyText.Checked;
  SaveAllData;
  if miNotesShowOnlyText.Checked = True then
  begin
    grTitles.Visible := False;
    pnGridSubjects.Visible := False;
    pnBtnSubjects.Visible := False;
    spSplitterSubjects.Visible := False;
    spSplitterNotes.Visible := False;
    spSplitterSubComm.Visible := False;
    pnNotesTop.Visible := False;
    pnSubNotesGrid.Visible := False;
    spSplitterAtt.Visible := False;
    pnAttachments.Visible := False;
    dbText.SetFocus;
    tmTimerSave.Enabled := True;
  end
  else
  begin
    grTitles.Visible := True;
    pnGridSubjects.Visible := True;
    pnBtnSubjects.Visible := True;
    spSplitterSubjects.Visible := True;
    spSplitterNotes.Visible := True;
    spSplitterSubComm.Visible := True;
    pnNotesTop.Visible := True;
    pnSubNotesGrid.Visible := True;
    spSplitterAtt.Visible := True;
    pnAttachments.Visible := True;
    tmTimerSave.Enabled := False;
  end;
end;

procedure TfmMain.miNotesShowActivitiesClick(Sender: TObject);
begin
  // Show activities
  if miNotesFind.Checked = True then
  begin
    miNotesFindClick(nil);
  end;
  miNotesShowActivities.Checked := not miNotesShowActivities.Checked;
  pnActivities.Visible := miNotesShowActivities.Checked;
  spActivities.Visible := miNotesShowActivities.Checked;
end;

procedure TfmMain.miNotesShowCalClick(Sender: TObject);
begin
  // Show calendar
  SaveAllData;
  fmCalendar.Caption := cpt066;
  fmCalendar.lbSelAct.Caption := cpt067;
  fmCalendar.lbStartDate.Caption := cpt068;
  fmCalendar.lbEndDate.Caption := cpt069;
  fmCalendar.lbResources.Caption := cpt070;
  fmCalendar.bnMonthDiary.Caption := cpt071;
  fmCalendar.bnTodayDiary.Caption := cpt072;
  fmCalendar.bnExportDiary.Caption := cpt073;
  fmCalendar.rgCalRange.Items[0] := cpt089;
  fmCalendar.rgCalRange.Items[1] := cpt090;
  fmCalendar.stAll := cpt084;
  fmCalendar.stNoRes := cpt085;
  fmCalendar.stNoDate := cpt086;
  fmCalendar.stResources := cpt087;
  fmCalendar.stCost := cpt088;
  fmCalendar.ShowModal;
end;

procedure TfmMain.bnNotesDownClick(Sender: TObject);
var
  sqMove: TSqlite3Dataset;
  iIDNotes, iSort1, iSort2: integer;
begin
  // Move Down note
  if sqNotes.RecNo < sqNotes.RecordCount then
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      iIDNotes := sqNotes.FieldByName('IDNotes').AsInteger;
      sqMove := TSqlite3Dataset.Create(Self);
      sqMove.FileName := sqNotes.FileName;
      sqMove.TableName := 'Notes';
      sqMove.PrimaryKey := 'IDNotes';
      sqMove.AutoIncrementKey := True;
      sqMove.SQL := sqNotes.SQL;
      sqMove.Open;
      sqMove.Locate('IDNotes', iIDNotes, []);
      iSort1 := sqMove.FieldByName('NotesSort').AsInteger;
      sqMove.Next;
      iSort2 := sqMove.FieldByName('NotesSort').AsInteger;
      sqMove.Edit;
      sqMove.FieldByName('NotesSort').AsInteger := iSort1;
      sqMove.Post;
      sqMove.ApplyUpdates;
      // Change data of the current record in sqNotes
      // to make it aware of the change
      sqNotes.Edit;
      sqNotes.FieldByName('NotesSort').AsInteger := iSort2;
      sqNotes.Post;
      sqNotes.ApplyUpdates;
      IsTextToLoad := False;
      sqNotes.RefetchData;
      sqNotes.Locate('IDNotes', iIDNotes, []);
      IsTextToLoad := True;
      sqNotesAfterScroll(nil);
    finally
      sqMove.Free;
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.bnNotesUpClick(Sender: TObject);
var
  sqMove: TSqlite3Dataset;
  iIDNotes, iSort1, iSort2: integer;
begin
  // Move up note
  if sqNotes.RecNo > 1 then
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      iIDNotes := sqNotes.FieldByName('IDNotes').AsInteger;
      sqMove := TSqlite3Dataset.Create(Self);
      sqMove.FileName := sqNotes.FileName;
      sqMove.TableName := 'Notes';
      sqMove.PrimaryKey := 'IDNotes';
      sqMove.AutoIncrementKey := True;
      sqMove.SQL := sqNotes.SQL;
      sqMove.Open;
      sqMove.Locate('IDNotes', iIDNotes, []);
      iSort1 := sqMove.FieldByName('NotesSort').AsInteger;
      sqMove.Prior;
      iSort2 := sqMove.FieldByName('NotesSort').AsInteger;
      sqMove.Edit;
      sqMove.FieldByName('NotesSort').AsInteger := iSort1;
      sqMove.Post;
      sqMove.ApplyUpdates;
      // Change data of the current record in sqNotes
      // to make it aware of the change
      sqNotes.Edit;
      sqNotes.FieldByName('NotesSort').AsInteger := iSort2;
      sqNotes.Post;
      IsTextToLoad := False;
      sqNotes.ApplyUpdates;
      sqNotes.RefetchData;
      sqNotes.Locate('IDNotes', iIDNotes, []);
      IsTextToLoad := True;
      sqNotesAfterScroll(nil);
    finally
      sqMove.Free;
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.bnSubDownClick(Sender: TObject);
var
  sqMove: TSqlite3Dataset;
  iIDSub, iSort1, iSort2: integer;
begin
  // Move down subject
  if sqSubjects.RecNo < sqSubjects.RecordCount then
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      iIDSub := sqSubjects.FieldByName('IDSubjects').AsInteger;
      sqMove := TSqlite3Dataset.Create(Self);
      sqMove.FileName := sqSubjects.FileName;
      sqMove.TableName := 'Subjects';
      sqMove.PrimaryKey := 'IDSubjects';
      sqMove.AutoIncrementKey := True;
      sqMove.SQL := sqSubjects.SQL;
      sqMove.Open;
      sqMove.Locate('IDSubjects', iIDSub, []);
      iSort1 := sqMove.FieldByName('SubjectsSort').AsInteger;
      sqMove.Next;
      iSort2 := sqMove.FieldByName('SubjectsSort').AsInteger;
      sqMove.Edit;
      sqMove.FieldByName('SubjectsSort').AsInteger := iSort1;
      sqMove.Post;
      sqMove.ApplyUpdates;
      // Change data of the current record in sqSubjects
      // to make it aware of the change
      sqSubjects.Edit;
      sqSubjects.FieldByName('SubjectsSort').AsInteger := iSort2;
      sqSubjects.Post;
      sqSubjects.ApplyUpdates;
      IsTextToLoad := False;
      sqSubjects.RefetchData;
      sqSubjects.Locate('IDSubjects', iIDSub, []);
      IsTextToLoad := True;
      sqNotesAfterScroll(nil);
    finally
      sqMove.Free;
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.bnSubUpClick(Sender: TObject);
var
  sqMove: TSqlite3Dataset;
  iIDSub, iSort1, iSort2: integer;
begin
  // Move up subject
  if sqSubjects.RecNo > 1 then
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      iIDSub := sqSubjects.FieldByName('IDSubjects').AsInteger;
      sqMove := TSqlite3Dataset.Create(Self);
      sqMove.FileName := sqSubjects.FileName;
      sqMove.TableName := 'Subjects';
      sqMove.PrimaryKey := 'IDSubjects';
      sqMove.AutoIncrementKey := True;
      sqMove.SQL := sqSubjects.SQL;
      sqMove.Open;
      sqMove.Locate('IDSubjects', iIDSub, []);
      iSort1 := sqMove.FieldByName('SubjectsSort').AsInteger;
      sqMove.Prior;
      iSort2 := sqMove.FieldByName('SubjectsSort').AsInteger;
      sqMove.Edit;
      sqMove.FieldByName('SubjectsSort').AsInteger := iSort1;
      sqMove.Post;
      sqMove.ApplyUpdates;
      // Change data of the current record in sqSubjects
      // to make it aware of the change
      sqSubjects.Edit;
      sqSubjects.FieldByName('SubjectsSort').AsInteger := iSort2;
      sqSubjects.Post;
      sqSubjects.ApplyUpdates;
      IsTextToLoad := False;
      sqSubjects.RefetchData;
      sqSubjects.Locate('IDSubjects', iIDSub, []);
      IsTextToLoad := True;
      sqNotesAfterScroll(nil);
    finally
      sqMove.Free;
      Screen.Cursor := crDefault;
    end;
end;

procedure TfmMain.miToolsSyncDoClick(Sender: TObject);
var
  IntFile, ExtFile: string;
  DelRecInt, DelRecExt, ChnRecInt, ChnRecExt: integer;
  OldIDSubjects, OldIDNotes: integer;
begin
  // Sync the current db
  if SyncFolder = '' then
    Exit;
  if FileExistsUTF8(SyncFolder + DirectorySeparator + ExtractFileName(sqSubjects.FileName)) = False then
    Exit;
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  //Make a copy of the current db
  if FileExistsUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName) + '.syn') then
  begin
    DeleteFileUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName) + '.syn');
  end;
  CopyFile(sqSubjects.FileName, ExtractFileNameWithoutExt(sqSubjects.FileName) + '.syn');
  // Record ID of curretn Subject and Note
  OldIDSubjects := sqSubjects.FieldByName('IDSubjects').AsInteger;
  OldIDNotes := sqNotes.FieldByName('IDNotes').AsInteger;
  // Set files name
  IntFile := sqSubjects.FileName;
  ExtFile := SyncFolder + DirectorySeparator + ExtractFileName(sqSubjects.FileName);
  DelRecInt := 0;
  DelRecExt := 0;
  ChnRecInt := 0;
  ChnRecExt := 0;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sbStatusBar.Panels[0].Text := ' ' + sbr005;
  // Check if the database structure must be upgraded
  try
    MakeTableUpgrade(ExtFile);
  except
    MessageDlg(msg007, mtWarning, [mbOK], 0);
    dbText.Clear;
    Abort;
  end;
  try
    try
      // Sync DelRec tables
      // Read from external, write to internal
      SyncDelRec(ExtFile, IntFile);
      // Read from internal, write to external
      SyncDelRec(IntFile, ExtFile);
      // Purge notes and subjects
      // Purge internal tables
      DelRecInt := SyncDelSubjectsNotes(ExtFile, IntFile);
      // Purge external tables
      DelRecExt := SyncDelSubjectsNotes(IntFile, ExtFile);
      // Sync notes and subjects
      // Sync internal tables
      ChnRecInt := SyncUpdateSubjectsNotes(ExtFile, IntFile);
      // Sync external tables
      ChnRecExt := SyncUpdateSubjectsNotes(IntFile, ExtFile);
      // Delete backup file
      if FileExistsUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName) + '.syn') then
      begin
        DeleteFileUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName) + '.syn');
      end;
      // Update all
      sqSubjects.Close;
      // Avoid loading text now
      IsTextToLoad := False;
      sqSubjects.Open;
      // sqNotes is closed and reopened in the sqSubjects.AfterScroll event
      // if there are subjects; otherwise nothing happens
      // Try to reach the last Subject and note selected
      sqSubjects.Locate('IDSubjects', OldIDSubjects, []);
      sqNotes.Locate('IDNotes', OldIDNotes, []);
      // Now text can be loaded
      IsTextToLoad := True;
      sqNotesAfterScroll(nil);
      dbText.SelStart := 0;
      Screen.Cursor := crDefault;
      sbStatusBar.Panels[0].Text := ' ' + sbr006;
      if flNoSyncMsg = False then
      begin
        if ((DelRecInt > 0) or (DelRecExt > 0) or (ChnRecInt > 0) or (ChnRecExt > 0)) then
          MessageDlg(msg010 + LineEnding + msg011 + ' ' + IntToStr(DelRecInt) + LineEnding + msg012 + ' ' + IntToStr(ChnRecInt) + LineEnding + LineEnding + msg013 + LineEnding + msg011 + ' ' + IntToStr(DelRecExt) + LineEnding + msg012 + ' ' + IntToStr(ChnRecExt),
            mtInformation, [mbOK], 0);
      end;
    except
      sbStatusBar.Panels[0].Text := ' ' + sbr007;
      MessageDlg(msg014 + LineEnding + ExtractFileNameWithoutExt(sqSubjects.FileName) + '.syn',
        mtWarning, [mbOK], 0);
    end;
  finally
    sqSubjects.EnableControls;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miToolsCompactClick(Sender: TObject);
var
  IDSub, IDNote: integer;
begin
  // Compact the database
  SaveAllData;
  // No show text only
  DisableShowTextOnly;
  if MessageDlg(msg015, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
  begin
    Abort;
  end;
  if FileExistsUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName) + '.bak') then
  begin
    DeleteFileUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName) + '.bak');
  end;
  CopyFile(sqSubjects.FileName,
    ExtractFileNameWithoutExt(sqSubjects.FileName) + '.bak');
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    if sqSubjects.RecordCount > 0 then
      IDSub := sqSubjects.FieldByName('IDSubjects').AsInteger
    else
      IDSub := 0;
    if sqNotes.RecordCount > 0 then
      IDNote := sqNotes.FieldByName('IDNotes').AsInteger
    else
      IDNote := 0;
    sqSubjects.Close;
    sqNotes.Close;
    try
      sqToolsTables.Close;
      sqToolsTables.FileName := sqSubjects.FileName;
      sqToolsTables.ExecSQL('Vacuum');
      // To avoid that the text is loaded
      IsTextToLoad := False;
      sqSubjects.Open;
      if sqSubjects.RecordCount > 0 then
        sqSubjects.Locate('IDSubjects', IDSub, []);
      // sqNotes is opened in the sqSubjects.AfterScroll event,
      // but if there are no subjects the dataset is not opened; so...
      if sqSubjects.RecordCount = 0 then
      begin
        sqNotes.SQL := 'Select * from Notes where IDNotes = -1';
        sqNotes.Open;
      end;
      if sqNotes.RecordCount > 0 then
        sqNotes.Locate('IDNotes', IDNote, []);
      // To avoid that the text is loaded
      IsTextToLoad := True;
      // Now load the text
      sqNotesAfterScroll(nil);
      Screen.Cursor := crDefault;
      MessageDlg(msg016 + LineEnding + ExtractFileNameWithoutExt(sqSubjects.FileName) + '.bak.', mtInformation, [mbOK], 0);
    except
      MessageDlg(msg017 + LineEnding + ExtractFileNameWithoutExt(sqSubjects.FileName) + '.bak.', mtWarning, [mbOK], 0);
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.miToolsEncryptGPGClick(Sender: TObject);
var
  Proc: TProcess;
  stFileNameNoSpace: string;
begin
  // Encrypt with GPG
  odOpenDialog.Title := cpt026;
  odOpenDialog.Filter := cpt027;
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute = True then
  begin
    if FileExistsUTF8(odOpenDialog.FileName + '.pgp') = True then
    begin
      MessageDlg(msg063, mtWarning, [mbOK], 0);
      Abort;
    end;
    InputQuery(msg064, msg065, False, stRecipient);
    if stRecipient = '' then
      Abort;
    stFileNameNoSpace :=
      StringReplace(odOpenDialog.FileName, ' ', '\ ', [rfReplaceAll]);
    try
      Proc := TProcess.Create(nil);
      Proc.Options := Proc.Options + [poWaitOnExit, poUsePipes];
      Proc.CommandLine := 'sh -c "gpg2 --batch ' + ' --armor --output ' + stFileNameNoSpace + '.pgp' + ' --recipient ' + stRecipient + ' --encrypt --sign ' + stFileNameNoSpace + '"';
      Proc.Execute;
      Proc.Free;
      Screen.Cursor := crDefault;
      if FileExistsUTF8(odOpenDialog.FileName + '.pgp') = False then
        MessageDlg(msg068, mtWarning, [mbOK], 0);
    except
      Screen.Cursor := crDefault;
      MessageDlg(msg068, mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.miToolsDecryptGPGClick(Sender: TObject);
var
  Proc: TProcess;
  stFileNameNoSpace: string;
begin
  // Decrypt with GPG
  odOpenDialog.Title := cpt042;
  odOpenDialog.Filter := cpt043;
  odOpenDialog.DefaultExt := '.pgp';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute = True then
  begin
    if FileExistsUTF8(ExtractFileNameWithoutExt(odOpenDialog.FileName)) = True then
    begin
      MessageDlg(msg069, mtWarning, [mbOK], 0);
      Abort;
    end;
    stFileNameNoSpace :=
      StringReplace(odOpenDialog.FileName, ' ', '\ ', [rfReplaceAll]);
    try
      Proc := TProcess.Create(nil);
      Proc.Options := Proc.Options + [poWaitOnExit, poUsePipes];
      Proc.CommandLine := 'sh -c "gpg2 --batch ' + ' --armor --output ' + ExtractFileNameWithoutExt(stFileNameNoSpace) + ' --decrypt ' + stFileNameNoSpace + '"';
      Proc.Execute;
      Proc.Free;
      Screen.Cursor := crDefault;
      if FileExistsUTF8(ExtractFileNameWithoutExt(odOpenDialog.FileName)) = False then
        MessageDlg(msg070, mtWarning, [mbOK], 0);
    except
      Screen.Cursor := crDefault;
      MessageDlg(msg070, mtWarning, [mbOK], 0);
    end;
  end;
end;

procedure TfmMain.miToolsAlarmClick(Sender: TObject);
begin
  // Show the alarm form
  fmSetAlarm.Caption := cpt098;
  fmSetAlarm.cbAlarm.Caption := cpt099;
  fmSetAlarm.ShowModal;
end;

procedure TfmMain.miToolsOptionsClick(Sender: TObject);
begin
  // Show options
  fmOptions.cbOptionsActivateTray.Checked := flTrayIcon;
  fmOptions.cbOptionsNoMsg.Checked := flNoSyncMsg;
  fmOptions.cbOptionsNoChar.Checked := flNoCharCount;
  fmOptions.cbOptionsNoAutosave.Checked := flNoAutosave;
  fmOptions.edOptParSpace.Text := IntToStr(ParagraphSpace);
  fmOptions.edOptLineSpace.Text := IntToStr(LineSpace);
  if flAutosync = True then
    fmOptions.cbOptionsAutosync.Checked := True;
  fmOptions.cbOptionsOpenLastFile.Checked := flOpenLastFile;
  fmOptions.tbOptionsFormTrans.Position := TranspForm;
  fmOptions.lbOptionsFont.Caption :=
    cpt044 + ' ' + DefFontName + ', ' + IntToStr(DefFontSize) + ' ' + cpt045;
  fmOptions.bnOptionsFont.Caption := cpt051;
  fmOptions.lbOptionsSyncDir.Caption := cpt046 + ' ' + SyncFolder;
  fmOptions.bnOptionsSyncDir.Caption := cpt052;
  fmOptions.lbOptionsFormColor.Caption := cpt047 + ' ' + ColorToString(Color);
  fmOptions.bnOptionsFormColor.Caption := cpt053;
  fmOptions.bnOptionsFormClDef.Caption := cpt062;
  fmOptions.lbOptionsFormTrans.Caption := cpt055;
  fmOptions.cbOptionsActivateTray.Caption := cpt048;
  fmOptions.cbOptionsOpenLastFile.Caption := cpt049;
  fmOptions.cbOptionsAutosync.Caption := cpt050;
  fmOptions.cbOptionsNoMsg.Caption := cpt074;
  fmOptions.cbOptionsNoChar.Caption := cpt075;
  fmOptions.lbOptionsColor.Caption := cpt076;
  fmOptions.bnOptFont1.Caption := cpt077;
  fmOptions.bnOptBack1.Caption := cpt078;
  fmOptions.bnOptFont2.Caption := cpt079;
  fmOptions.bnOptBack2.Caption := cpt080;
  fmOptions.bnOptFont3.Caption := cpt081;
  fmOptions.bnOptBack3.Caption := cpt082;
  fmOptions.lbOptionDef1.Caption := cpt063;
  fmOptions.lbOptionDef2.Caption := cpt064;
  fmOptions.lbOptionDef3.Caption := cpt065;
  fmOptions.cbOptionsNoAutosave.Caption := cpt095;
  fmOptions.lbOptionsLine.Caption := cpt096;
  fmOptions.lbOptionsPar.Caption := cpt097;
  fmOptions.Caption := cpt083;
  fmOptions.ShowModal;
end;

procedure TfmMain.miToolsLanguageClick(Sender: TObject);
begin
  // Install language file
  SaveAllData;
  odOpenDialog.Title := cpt028;
  odOpenDialog.Filter := cpt029;
  odOpenDialog.DefaultExt := '.lng';
  odOpenDialog.FileName := '';
  if odOpenDialog.Execute = True then
  begin
    CopyFile(odOpenDialog.FileName,
      GetEnvironmentVariable('HOME') + '/.config' + DirectorySeparator + 'mynotex' + DirectorySeparator + 'translation-' + VersMyNt);
    // Load and activate translation
    Translation;
    // Update status bar
    if sqNotes.Active = False then
      sbStatusBar.Panels[0].Text := ' ' + sbr008
    else if sqNotes.RecordCount = 0 then
      sbStatusBar.Panels[0].Text := ' ' + sbr002
    else
      sbStatusBar.Panels[0].Text :=
        ' ' + sbr003 + ' ' + IntToStr(sqNotes.RecNo) + ' ' + sbr004 + ' ' + IntToStr(sqNotes.RecordCount) + ' - ' + sbr009 + ' ' + FormatDateTime(FDate.LongDateFormat, sqNotes.FieldByName('NotesDTMod').AsDateTime) + ' ' + sbr010 + ' ' + FormatDateTime('hh:nn', sqNotes.FieldByName('NotesDTMod').AsDateTime);
  end;
end;

procedure TfmMain.miHelpClick(Sender: TObject);
begin
  // Show pdf manual
  // e.g. manual-mynotex-it.pdf
  if FileExists(InstallDir + 'manual-mynotex-' + LowerCase(Copy(GetEnvironmentVariable('LANG'), 1, 2) + '.pdf')) then
    OpenDocument(InstallDir + 'manual-mynotex-' + LowerCase(Copy(GetEnvironmentVariable('LANG'), 1, 2) + '.pdf'))
  else
    // Show the official English manual
    OpenDocument(InstallDir + 'manual-mynotex-en.pdf');
end;

procedure TfmMain.miLicenceClick(Sender: TObject);
begin
  // Show copyright
  fmCopyright.lbCopyrightAuthor1.Caption := cpt033;
  fmCopyright.lbCopyrightSite.Caption := cpt034;
  fmCopyright.lbCopyrightForum.Caption := cpt040;
  fmCopyright.ShowModal;
end;

procedure TfmMain.tiTrayIconClick(Sender: TObject);
begin
  // Tray icon management
  if fmMain.WindowState = wsMinimized then
  begin
    fmMain.Show;
    if FmMaximized = True then
      fmMain.WindowState := wsMaximized
    else
      fmMain.WindowState := wsNormal;
  end
  else
  begin
    fmMain.WindowState := wsMinimized;
  end;
end;

// *****************************************************************************
// ************************* RICHMEMO PROCEDURES *******************************
// *****************************************************************************

function TfmMain.SaveRichMemo(FromIdx, ToIdx: integer; SavePictures: boolean): WideString;
var
  FpNew, FpOld: TFontParams;
  i: integer;
  StText, StOrig: WideString;
  flColBack: boolean = False;
  flAlign: boolean = False;
  stAlign: string = '';
  flIndented: boolean = False;
  slPosImg: TStringList;
  AttDir: string;
begin
  // Create text with HTML tags from RichMemo content
  if dbText.Text = '' then
  begin
    Result := '';
    Exit;
  end;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  FpOld.Color := clBlack;
  FpOld.Name := DefFontName;
  FpOld.Size := -1; // To force the specification of font as first tag
  FpOld.Style := [];
  FpOld.BackColor := clWhite;
  FpOld.Alignment := []; // To force the specification of align
  FpOld.Indented := DefaultIndent;
  StOrig := dbText.Text;
  // Save images
  i := -1;
  if SavePictures = True then
  begin
    AttDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
    if DirectoryExistsUTF8(AttDir) = False then
      CreateDirUTF8(AttDir);
    try
      // slPosImg is created in the TRichMemo component code,
      // in the GetImagePosInText event
      slPosImg := dbText.GetImagePosInText;
      if slPosImg.Count > 0 then
      begin
        for i := 0 to slPosImg.Count - 1 do
        begin
          dbText.SaveImageToFile(StrToInt(slPosImg[i]) - 1,
            AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
          StOrig := UTF8Copy(StOrig, 1, StrToInt(slPosImg[i]) - 1) + #2 + UTF8Copy(StOrig, StrToInt(slPosImg[i]), UTF8Length(StOrig));
          Inc(ToIdx);
        end;
      end;
    finally
      slPosImg.Free;
    end;
    // Delete unused images
    while FileExistsUTF8(AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i + 1) + '.jpeg') = True do
      try
        DeleteFileUTF8(AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i + 1) + '.jpeg');
        Inc(i);
      except
        MessageDlg(msg035, mtWarning, [mbOK], 0);
      end;
    if fmMain.IsDirectoryEmpty(AttDir) = True then
      DeleteDirectory(AttDir, False);
  end
  else
  begin
    try
      // slPosImg is created in the TRichMemo component code,
      // in the GetImagePosInText event
      slPosImg := dbText.GetImagePosInText;
      if slPosImg.Count > 0 then
      begin
        for i := 0 to slPosImg.Count - 1 do
        begin
          StOrig := UTF8Copy(StOrig, 1, StrToInt(slPosImg[i]) - 1) + #2 + UTF8Copy(StOrig, StrToInt(slPosImg[i]), UTF8Length(StOrig));
        end;
      end;
    finally
      slPosImg.Free;
    end;
  end;
  // To avoid confusion with html tags
  StOrig := StringReplace(StOrig, '<', #5, [rfReplaceAll]);
  StOrig := StringReplace(StOrig, '>', #6, [rfReplaceAll]);
  StText := '';
  try
    for i := FromIdx to ToIdx do
    begin
      dbText.GetTextAttributes(i, FpNew);
      if AreFontParamsEqual(FpNew, FpOld) = False then
      begin
        // See notes on AreFontParamsEqual for following code
        if ((FpNew.Size < 2) or (FpNew.Size > 128)) then
          FpNew.Size := DefFontSize;
        if ((FpNew.Color <> FpOld.Color) or (FpNew.Size <> FpOld.Size) or (FpNew.Name <> FpOld.Name)) then
        begin
          StText := StText + '</font><font' + ' color="#' + IntToHex(Red(FpNew.Color), 2) + IntToHex(Green(FpNew.Color), 2) + IntToHex(Blue(FpNew.Color), 2) + '"' +
            // Converted from 1 to 7 on html export
            ' size="' + IntToStr(FpNew.Size - ZoomFontSize) + '"' + ' face="' + FpNew.Name + '">';
        end;
        if FpNew.BackColor <> FpOld.BackColor then
        begin
          if flColBack = True then
            StText := StText + '</span>';
          if FpNew.BackColor <> clWhite then
          begin
            StText := StText + '<span style=' + '"background: #' + IntToHex(Red(FpNew.BackColor), 2) + IntToHex(Green(FpNew.BackColor), 2) + IntToHex(Blue(FpNew.BackColor), 2) + '">';
            flColBack := True;
          end
          else
            flColBack := False;
        end;
        // Align tag must be inserted at the beginning of each line
        // to be exported in LibreOffice;
        // here we set the flag, below we insert it in the text
        if FpNew.Alignment <> FpOld.Alignment then
        begin
          if flAlign = True then
            stAlign := '</p>'
          else
            stAlign := '';
          if FpNew.Alignment = [trLeft] then
            stAlign := stAlign + '<p align=left>'
          else if FpNew.Alignment = [trCenter] then
            stAlign := stAlign + '<p align=center>'
          else if FpNew.Alignment = [trRight] then
            stAlign := stAlign + '<p align=right>'
          else if FpNew.Alignment = [trJustified] then
            stAlign := stAlign + '<p align=justify>';
          flAlign := True;
        end;
        // Indented final tag must be inserted at the end of the line;
        // here we set the flag, below we insert it in the text
        if FpNew.Indented <> FpOld.Indented then
        begin
          if FpNew.Indented > DefaultIndent then
          begin
            StText := StText + '<blockquote>';
            flIndented := False;
          end
          else
            flIndented := True;
        end;
        if FpNew.Style <> FPOld.Style then
        begin
          if ((fsBold in FpNew.Style) and not (fsBold in FpOld.Style)) then
            StText := StText + '<b>';
          if (not (fsBold in FpNew.Style) and (fsBold in FpOld.Style)) then
            StText := StText + '</b>';
          if ((fsItalic in FpNew.Style) and not (fsItalic in FpOld.Style)) then
            StText := StText + '<i>';
          if (not (fsItalic in FpNew.Style) and (fsItalic in FpOld.Style)) then
            StText := StText + '</i>';
          if ((fsUnderline in FpNew.Style) and not (fsUnderline in FpOld.Style)) then
            StText := StText + '<u>';
          if (not (fsUnderline in FpNew.Style) and (fsUnderline in FpOld.Style)) then
            StText := StText + '</u>';
          if ((fsStrikeOut in FpNew.Style) and not (fsStrikeOut in FpOld.Style)) then
            StText := StText + '<strike>';
          if (not (fsStrikeOut in FpNew.Style) and (fsStrikeOut in FpOld.Style)) then
            StText := StText + '</strike>';
        end;
        FpOld := FpNew;
      end;
      // Align tag must be inserted at the beginning of each line;
      // here we insert the flag on a CR
      if ((StOrig[i] = LineEnding) or (i = FromIdx)) then
      begin
        StText := StText + stAlign;
      end;
      // Indented final tag must be inserted at the end of the line;
      // here we insert the flag on a CR
      if ((flIndented = True) and (StOrig[i + 1] = LineEnding)) then
      begin
        StText := StText + '</blockquote>';
        flIndented := False;
      end;
      StText := StText + StOrig[i + 1];
    end;
    // To remove the last char, i + 1
    StText := UTF8Copy(StText, 1, UTF8Length(StText) - 1);
    if flColBack = True then
      StText := StText + '</span>';
    StText := StText + '</p>';
    if flIndented = True then
      StText := StText + '</blockquote>';
    StText := StringReplace(StText, '</blockquote>' + LineEnding + '<blockquote>', LineEnding, [rfReplaceAll]);
    if UTF8Copy(StText, 1, 7) = '</font>' then
      StText := UTF8Copy(StText, 8, UTF8Length(StText) - 7);
    StText := StText + '</font>';
    // To avoid a partial formatting of indented paragraph,
    // which would add a CR in hmtl exportation
    // insert this line: StText :=
    // StringReplace(StText, '</blockquote><blockquote>', '', [rfReplaceAll]);
    // Create img tag
    i := 0;
    while UTF8Pos(#2, StText) > 0 do
    begin
      if SavePictures = True then
        StText := StringReplace(StText, #2, '<IMG SRC="' + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg">', [])
      else
        StText := StringReplace(StText, #2, '', []);
      Inc(i);
    end;
    Result := StText;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.LoadRichMemo;
var
  FpNew: TFontParams;
  i, n, idxStart, idxLength, idxStartTag, idxEndTag: integer;
  StTags, StSubTag: string;
  StData, StOrig, StDest: WideString;
  IsTag, IsSubTag: boolean;
  AttDir: string = '';
begin
  // Load HTML text from db and format it in RichMemo
  // Note that RichMemo.text cannot be changed without reset formatting.
  if IsTextToLoad = False then
    Exit;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  StDest := '';
  if sqNotes.FieldByName('NotesCheckPwd').AsString = '' then
    StData := sqNotes.FieldByName('NotesText').AsWideString
  else
  begin
    dcAES.InitStr(edPassword.Text, TDCP_sha1);
    StData := dcAES.DecryptString(sqNotes.FieldByName('NotesText').AsWideString);
  end;
  // Copy text without tags in RichText
  StOrig := StData;
  IsTag := False;
  dbText.Clear;
  for i := 1 to Length(StOrig) do
  begin
    if StOrig[i] = '<' then
      IsTag := True
    else if StOrig[i] = '>' then
      IsTag := False
    else if IsTag = False then
      StDest := StDest + StOrig[i];
  end;
  // Restore < and > insert by user, not as html tags
  StDest := StringReplace(StDest, #5, '<', [rfReplaceAll]);
  StDest := StringReplace(StDest, #6, '>', [rfReplaceAll]);
  dbText.Text := StDest;
  if DirectoryExistsUTF8(ExtractFileNameWithoutExt(sqSubjects.FileName)) = True then
    AttDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
  // Format the text in the RichText
  try
    StTags := '';
    idxStart := 0;
    for i := 1 to Length(StOrig) do
    begin
      if StOrig[i] = '<' then
      begin
        IsTag := True;
        StTags := StTags + StOrig[i];
      end
      else if StOrig[i] = '>' then
      begin
        StTags := StTags + StOrig[i]; // Later will be removed
        // Load images
        if ((AttDir <> '') and (UTF8Pos('<IMG SRC=', StTags) > 0)) then
          try
            dbText.LoadImageFromFile(idxStart, AttDir + DirectorySeparator + Copy(StTags, 11, 49), 1);
            // In the dbText text a character is added as picture
            Inc(idxStart);
          except;
          end
        // Tag <font ...
        else if UTF8Pos('<font color="', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</font>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          // Get the color, always of 6 chars
          idxStartTag := UTF8Pos(' color="#', StTags) + 9;
          FpNew.Color := StringToColor('$00' + UTF8Copy(StTags, idxStartTag + 4, 2) + UTF8Copy(StTags, idxStartTag + 2, 2) + UTF8Copy(StTags, idxStartTag + 0, 2));
          // Get the size
          idxStartTag := UTF8Pos(' size="', StTags) + 7;
          idxEndTag := idxStartTag;
          while StTags[idxEndTag] <> '"' do
            Inc(idxEndTag);
          idxEndTag := idxEndTag - idxStartTag;
          FpNew.Size := StrToInt(UTF8Copy(StTags, idxStartTag, idxEndTag)) + ZoomFontSize;
          // Get the name
          idxStartTag := UTF8Pos(' face="', StTags) + 7;
          idxEndTag := idxStartTag;
          while StTags[idxEndTag] <> '"' do
            Inc(idxEndTag);
          idxEndTag := idxEndTag - idxStartTag;
          FpNew.Name := UTF8Copy(StTags, idxStartTag, idxEndTag);
          // Set the font properties
          FpNew.Changed := [fiName, fiSize, fiColor];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <b>
        else if UTF8Pos('<b>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</b>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Style := [fsBold];
          FpNew.Changed := [fiBold];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <i>
        else if UTF8Pos('<i>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</i>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Style := [fsItalic];
          FpNew.Changed := [fiItalic];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <u>
        else if UTF8Pos('<u>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</u>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Style := [fsUnderline];
          FpNew.Changed := [fiUnderline];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <strike>
        else if UTF8Pos('<strike>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</strike>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Style := [fsStrikeOut];
          FpNew.Changed := [fiStrike];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <p align=left>
        else if UTF8Pos('<p align=left>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</p>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Alignment := [trLeft];
          FpNew.Changed := [fiAlignment];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <p align=right>
        else if UTF8Pos('<p align=right>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</p>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Alignment := [trRight];
          FpNew.Changed := [fiAlignment];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <p align=center>
        else if UTF8Pos('<p align=center>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</p>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Alignment := [trCenter];
          FpNew.Changed := [fiAlignment];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <p align=justify>
        else if UTF8Pos('<p align=justify>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</p>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Alignment := [trJustified];
          FpNew.Changed := [fiAlignment];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <span style= ...
        else if UTF8Pos('<span style="', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</span>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          // Get the color, always of 6 chars
          idxStartTag := UTF8Pos('"background: #', StTags) + 14;
          FpNew.BackColor := StringToColor('$00' + UTF8Copy(StTags, idxStartTag + 4, 2) + UTF8Copy(StTags, idxStartTag + 2, 2) + UTF8Copy(StTags, idxStartTag + 0, 2));
          // Set the font properties
          FpNew.Changed := [fiBackcolor];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end
        // Tag <blockquote>
        else if UTF8Pos('<blockquote>', StTags) > 0 then
        begin
          idxLength := 0;
          StSubTag := '';
          for n := i to Length(StOrig) do
          begin
            if StOrig[n] = '<' then
            begin
              IsSubTag := True;
              StSubTag := StSubTag + StOrig[n];
            end
            else if StOrig[n] = '>' then
            begin
              StSubTag := StSubTag + StOrig[n];
              if StSubTag = '</blockquote>' then
                Break;
              IsSubTag := False;
              StSubTag := '';
            end
            else
            begin
              if IsSubTag = False then
                Inc(idxLength)
              else
                StSubTag := StSubTag + StOrig[n];
            end;
          end;
          FpNew.Indented := WidthInden;
          FpNew.Changed := [fiIndented];
          dbText.SetTextAttributes(idxStart, idxLength, FpNew);
        end;
        IsTag := False;
        StTags := '';
      end
      else
      begin
        if IsTag = True then
          StTags := StTags + StOrig[i]
        else
          Inc(idxStart);
      end;
    end;
  finally
    SetCharCount(flNoCharCount);
    dbText.SelStart := 0;
    // Necessary to let the cursor move to char 0
    Application.ProcessMessages;
    if UTF8Length(dbText.Text) = 0 then
      ClearUndoData
    else
      StoreUndoData;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.tbCutClick(Sender: TObject);
begin
  // Cut to clipboard
  dbText.CutToClipboard;
end;

procedure TfmMain.tbCopyClick(Sender: TObject);
begin
  // Copy to clipboard
  dbText.CopyToClipboard;
end;

procedure TfmMain.tbCopyHtmlClick(Sender: TObject);
var
  Fid: TClipboardFormat;
  Strm: TMemoryStream;
  StHTML: string;
  flNumList: boolean = False;
  FnSize, HTMLFnSize, NumList, IniIndent, EndIndent: integer;
begin
  // Copy to clipboard in html format
  Strm := TMemoryStream.Create;
  StHTML := SaveRichMemo(dbText.SelStart, dbText.SelStart + dbText.SelLength, False);
  // Restore < and >
  StHTML := StringReplace(StHTML, #5, '<', [rfReplaceAll]);
  StHTML := StringReplace(StHTML, #6, '>', [rfReplaceAll]);
  StHTML := '<html><head><meta http-equiv="Content-Type" ' + 'content="text/html; charset=UTF-8"></head><body>' + StHTML + ' </body>';
  // Set font size in html format (from 1 to 7)
  for FnSize := 6 to 72 do
  begin
    if FnSize < 7 then
      HTMLFnSize := 1
    else if FnSize < 10 then
      HTMLFnSize := 2
    else if FnSize < 13 then
      HTMLFnSize := 3
    else if FnSize < 17 then
      HTMLFnSize := 4
    else if FnSize < 21 then
      HTMLFnSize := 5
    else if FnSize < 25 then
      HTMLFnSize := 6
    else
      HTMLFnSize := 7;
    StHTML := StringReplace(StHTML, 'size="' + IntToStr(FnSize) + '" face="', 'size="' + IntToStr(HTMLFnSize) + '" face="', [rfIgnoreCase, rfReplaceAll]);
  end;
  // Set the tags for lists
  IniIndent := 0;
  EndIndent := 0;
  while UTF8Pos('<blockquote>', StHTML) > 0 do
  begin
    flNumList := False;
    IniIndent := UTF8Pos('<blockquote>', StHTML) + 13;
    if UTF8Pos('</blockquote>', StHTML) - 1 > 0 then
      EndIndent := UTF8Pos('</blockquote>', StHTML) - 1
    else
      // If copying an indented text not to the end of it
      EndIndent := UTF8Length(StHTML);
    while IniIndent < EndIndent do
    begin
      if ((UTF8Copy(StHTML, IniIndent, 3) = '>1.') or (UTF8Copy(StHTML, IniIndent, 3) = '>A.') or (UTF8Copy(StHTML, IniIndent, 3) = '>a.')) then
        flNumList := True;
      if StHTML[IniIndent] = LineEnding then
        StHTML[IniIndent] := #3; // A code to be replaced with <il>
      Inc(IniIndent);
    end;
    StHTML := StringReplace(StHTML, #3, '</li><li>', [rfReplaceAll]);
    if flNumList = True then
    begin
      StHTML := StringReplace(StHTML, '<blockquote>', '<p align=left><ol><li>', []);
      StHTML := StringReplace(StHTML, '</blockquote>', '</li></ol></p>', []);
    end
    else
    begin
      StHTML := StringReplace(StHTML, '<blockquote>', '<p align=left><ul><li>', []);
      StHTML := StringReplace(StHTML, '</blockquote>', '</li></ul></p>', []);
    end;
  end;
  // A list has been processed
  if IniIndent > 0 then
  begin
    // Delete numbers, letters and bullets
    for NumList := 1 to 100 do
    begin
      StHTML := StringReplace(StHTML, IntToStr(NumList) + '.' + #9, '', [rfReplaceAll]);
    end;
    for NumList := 65 to 90 do
    begin
      StHTML :=
        StringReplace(StHTML, char(NumList) + '.' + #9, '', [rfReplaceAll]);
    end;
    for NumList := 97 to 122 do
    begin
      StHTML :=
        StringReplace(StHTML, char(NumList) + '.' + #9, '', [rfReplaceAll]);
    end;
    StHTML := StringReplace(StHTML, '•' + #9, '', [rfReplaceAll]);
    // Remove alignment tags within a list to avoid empty rows
    // Alignment is always left
    StHTML := StringReplace(StHTML, '<li><p align=left>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li><p align=center>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li><p align=right>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li><p align=justify>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li></p><p align=left>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li></p><p align=center>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li></p><p align=right>', '<li>', [rfReplaceAll]);
    StHTML := StringReplace(StHTML, '<li></p><p align=justify>', '<li>', [rfReplaceAll]);
  end;
  StHTML := StringReplace(StHTML, LineEnding, '<p>', [rfReplaceAll]);
  Strm.WriteBuffer(Pointer(StHTML)^, Length(StHTML));
  Strm.Position := 0;
  Clipboard.Clear;
  Fid := Clipboard.FindFormatID('text/html');
  if Fid = 0 then
    Fid := RegisterClipboardFormat('text/html');
  Clipboard.AddFormat(Fid, Strm);
  Strm.Free;
end;

procedure TfmMain.pmTextCopyLatexClick(Sender: TObject);
var
  Fid: TClipboardFormat;
  Strm: TMemoryStream;
  StLatex, StLatexNoCode: string;
  NumList, IniIndent, EndIndent, i: integer;
  flNumList, IsSubTag: boolean;
begin
  // Copy to clipboard in html format
  Strm := TMemoryStream.Create;
  StLatex := SaveRichMemo(dbText.SelStart, dbText.SelStart + dbText.SelLength, False);
  StLatex := StringReplace(StLatex, '<i>', '\textit{', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '</i>', '}', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '<b>', '\textbf{', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '</b>', '}', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '<u>', '\underline{', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '</u>', '}', [rfReplaceAll]);
  // Set possibile misplacements of brakets
  StLatex := StringReplace(StLatex, '\textit{}', '}\textit{', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\textbf{}', '}\textbf{', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\underline{}', '}\underline{', [rfReplaceAll]);
  while UTF8Pos('<p align=', StLatex) > 0 do
  begin
    // Add Latex code only for center and right; left and justified comes as justified
    if UTF8Copy(StLatex, UTF8Pos('<p align=', StLatex) + 9, 6) = 'center' then
    begin
      StLatex := StringReplace(StLatex, '<p align=center>', '\begin{center}', [rfIgnoreCase]);
      StLatex := StringReplace(StLatex, '</p>', '\end{center}', [rfIgnoreCase]);
    end
    else if UTF8Copy(StLatex, UTF8Pos('<p align=', StLatex) + 9, 5) = 'right' then
    begin
      StLatex := StringReplace(StLatex, '<p align=right>', '\begin{flushright}', [rfIgnoreCase]);
      StLatex := StringReplace(StLatex, '</p>', '\end{flushright}', [rfIgnoreCase]);
    end
    else if UTF8Copy(StLatex, UTF8Pos('<p align=', StLatex) + 9, 4) = 'left' then
    begin
      StLatex := StringReplace(StLatex, '<p align=left>', '', [rfIgnoreCase]);
      StLatex := StringReplace(StLatex, '</p>', '', [rfIgnoreCase]);
    end
    else if UTF8Copy(StLatex, UTF8Pos('<p align=', StLatex) + 9, 7) = 'justify' then
    begin
      StLatex := StringReplace(StLatex, '<p align=justify>', '', [rfIgnoreCase]);
      StLatex := StringReplace(StLatex, '</p>', '', [rfIgnoreCase]);
    end;
  end;
  // Put the paragraphs tags within the characters tags
  StLatex := StringReplace(StLatex, '}' + LineEnding + '\end{center}', '\end{center}}' + LineEnding, [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '}' + LineEnding + '\end{flushright}', '\end{flushright}}' + LineEnding, [rfReplaceAll]);
  // Set the tags for lists
  IniIndent := 0;
  EndIndent := 0;
  while UTF8Pos('<blockquote>', StLatex) > 0 do
  begin
    flNumList := False;
    IniIndent := UTF8Pos('<blockquote>', StLatex) + 12;
    EndIndent := UTF8Pos('</blockquote>', StLatex) - 1;
    if UTF8Copy(StLatex, IniIndent, 2) = '1.' then
      flNumList := True;
    while IniIndent < EndIndent do
    begin
      if StLatex[IniIndent] = LineEnding then
        StLatex[IniIndent] := #3; // A code to be replaced with \items
      Inc(IniIndent);
    end;
    StLatex := StringReplace(StLatex, #3, '\item ', [rfReplaceAll]);
    if flNumList = True then
    begin
      StLatex := StringReplace(StLatex, '<blockquote>', LineEnding + LineEnding + '\begin{enumerate}' + LineEnding + LineEnding + '\item ', []);
      StLatex := StringReplace(StLatex, '</blockquote>', LineEnding + LineEnding + '\end{enumerate}' + LineEnding + LineEnding, []);
    end
    else
    begin
      StLatex := StringReplace(StLatex, '<blockquote>', LineEnding + LineEnding + '\begin{itemize}' + LineEnding + LineEnding + '\item ', []);
      StLatex := StringReplace(StLatex, '</blockquote>', LineEnding + LineEnding + '\end{itemize}' + LineEnding + LineEnding, []);
    end;
  end;
  // A list has been processed
  if IniIndent > 0 then
  begin
    // Delete bullets and numbers
    for NumList := 1 to 100 do
    begin
      StLatex := StringReplace(StLatex, IntToStr(NumList) + '.' + #9, '', [rfReplaceAll]);
      StLatex := StringReplace(StLatex, IntToStr(NumList) + '. ', '', [rfReplaceAll]);
    end;
    StLatex := StringReplace(StLatex, '•' + #9, '', [rfReplaceAll]);
    StLatex := StringReplace(StLatex, '• ', '', [rfReplaceAll]);
  end;
  while UTF8Pos('"', stLatex) > 0 do
  begin
    StLatex := StringReplace(StLatex, '"', '“', []);
    StLatex := StringReplace(StLatex, '"', '”', []);
  end;
  StLatex := StringReplace(StLatex, '\item', LineEnding + '\item', [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\item ' + LineEnding, LineEnding, [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\item \begin{itemize}', LineEnding + LineEnding + '\begin{itemize}' + LineEnding + LineEnding, [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\item \begin{enumerate}', LineEnding + LineEnding + '\begin{enumerate}' + LineEnding + LineEnding, [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\item \end{itemize}', LineEnding + LineEnding + '\end{itemize}' + LineEnding + LineEnding, [rfReplaceAll]);
  StLatex := StringReplace(StLatex, '\item \end{enumerate}', LineEnding + LineEnding + '\end{enumerate}' + LineEnding + LineEnding, [rfReplaceAll]);
  StLatex := StringReplace(StLatex, LineEnding + LineEnding + LineEnding, LineEnding + LineEnding, [rfReplaceAll]);
  // Replace todo symbols
  if ((UTF8Pos('', StLatex) > 0) or (UTF8Pos('', StLatex) > 0)) then
  begin
    StLatex := StringReplace(StLatex, '', '\Square \hspace{10pt}', [rfReplaceAll]);
    StLatex := StringReplace(StLatex, '', '\CheckedBox \hspace{10pt}', [rfReplaceAll]);
    StLatex := '% Add \usepackage{wasysym}' + LineEnding + StLatex;
  end;
  // do not use UTF8Length here below because it shortens the string
  IsSubTag := False;
  // Restore < and >
  StLatexNoCode := StringReplace(StLatexNoCode, #5, '<', [rfReplaceAll]);
  StLatexNoCode := StringReplace(StLatex, #6, '>', [rfReplaceAll]);
  StLatexNoCode := '';
  for i := 1 to Length(StLatex) do
  begin
    if StLatex[i] = '<' then
      IsSubTag := True
    else if StLatex[i] = '>' then
      IsSubTag := False
    else if IsSubTag = False then
      StLatexNoCode := StLatexNoCode + StLatex[i];
  end;
  StLatexNoCode := StringReplace(StLatexNoCode, LineEnding + LineEnding + LineEnding, LineEnding + LineEnding, [rfReplaceAll]);
  Strm.WriteBuffer(Pointer(StLatexNoCode)^, Length(StLatexNoCode));
  Strm.Position := 0;
  Clipboard.Clear;
  Fid := Clipboard.FindFormatID('text/plain');
  if Fid = 0 then
    Fid := RegisterClipboardFormat('text/plain');
  Clipboard.AddFormat(Fid, Strm);
  Strm.Free;
end;

procedure TfmMain.pmTextSendAsEmailClick(Sender: TObject);
var
  Proc: TProcess;
  stSub, stBody: string;
begin
  // Send email
  // In Thunderbird it works fine, in Evolution the CR are shown (why?)
  if dbText.Text <> '' then
    try
      stSub := StringReplace(dbTitle.Text, '"', '“', [rfReplaceAll]);
      stBody := StringReplace(dbText.Text, '"', '“', [rfReplaceAll]);
      Proc := TProcess.Create(nil);
      Proc.CommandLine := 'xdg-email --utf8 --subject "' + stSub + '" --body "' + stBody + '"';
      Proc.Execute;
    finally
      Proc.Free;
    end;
end;

procedure TfmMain.tbPasteClick(Sender: TObject);
begin
  // Paste from clipboard
  dbText.PasteFromClipboard;
  MakeLink(CheckLink);
  SetCharCount(flNoCharCount);
end;

procedure TfmMain.tbKindFontChangeClick(Sender: TObject);
var
  fp, fpold: TFontParams;
  i, idxStart, idxLength, idxStep: integer;
begin
  // Apply selected font (also for fdFontSelDialogApplyClicked)
  // There is a selection
  if dbText.SelLength > 0 then
  begin
    idxStart := dbText.SelStart;
    idxLength := dbText.SelLength;
  end
  // There is no selection: the formatting is applied to the current word
  else
  begin
    idxStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, True, True);
    idxLength := dbText.GetWordParagraphStartEnd(dbText.SelStart, True, False) - idxStart;
  end;
  idxStep := idxStart;
  dbText.GetTextAttributes(idxStep, fpold);
  for i := idxStart to idxStart + idxLength - 1 do
  begin
    dbText.GetTextAttributes(i, fp);
    // Change of formatting or last character
    if ((AreFontParamsEqual(fpold, fp) = False) or (i = idxStart + idxLength - 1)) then
    begin
      // Last character
      if i = idxStart + idxLength - 1 then
        dbText.GetTextAttributes(i, fp)
      // Change of formatting
      else
        dbText.GetTextAttributes(i - 1, fp);
      fp.Size := fdFontSelDialog.Font.Size + ZoomFontSize;
      fp.Style := fdFontSelDialog.Font.Style;
      fp.Name := fdFontSelDialog.Font.Name;
      // Last character
      if i = idxStart + idxLength - 1 then
      begin
        fp.Changed := [fiName, fiSize, fiColor];
        dbText.SetTextAttributes(idxStep, i - idxStep + 1, fp);
      end
      // Change of formatting
      else
      begin
        fp.Changed := [fiName, fiSize, fiColor];
        dbText.SetTextAttributes(idxStep, i - idxStep, fp);
        idxStep := i;
        dbText.GetTextAttributes(idxStep, fpold);
      end;
    end;
  end;
  // Put the Notes dataset in Edit and set the flag for user change of RichMemo
  EditNotesDataset;
end;

procedure TfmMain.pmChangeFontKindClick(Sender: TObject);
begin
  // Set font kind for formatting
  if fdFontSelDialog.Execute = True then
    tbKindFontChangeClick(nil);
end;

procedure TfmMain.pmOpenOpenOfficeClick(Sender: TObject);
begin
  // Select the word processor to open text with
  // (valid for OO, LibreOffice and browser)
  if Sender = pmOpenOpenOffice then
    tbOpenNote.Tag := 0
  else if Sender = pmOpenLibreOffice then
    tbOpenNote.Tag := 1
  else if Sender = pmOpenBrowser then
    tbOpenNote.Tag := 2;
  tbOpenNoteClick(nil);
end;

procedure TfmMain.tbOpenNoteClick(Sender: TObject);
var
  Proc: TProcess;
  stCommand, SubNotesText, SubActText, stSpcInd: string;
  // NewHtml, CharHtml: String;
  myFile: TextFile;
  FnSize, HTMLFnSize, NumList, IniIndent, EndIndent, SpIndent, PosName, ImgNum, i: integer;
  flNumList: boolean = False;
  AttReadDir, AttWriteDir: string;
  myHTMLList: TStringList;
  MyPrinter: TPrinter;
begin
  // Open text with word processor or print
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  SaveAllData;
  try
    // Delete possibile file html
    if FileExistsUTF8(GetTempDir + DirectorySeparator + 'MyNotexFile.html') then
      try
        DeleteFileUTF8(GetTempDir + DirectorySeparator + 'MyNotexFile.html');
      except
        MessageDlg(msg035, mtWarning, [mbOK], 0);
      end;
    // Create file html
    AssignFile(myFile, GetTempDir + DirectorySeparator + 'MyNotexFile.html');
    ReWrite(myFile);
    WriteLn(myFile, '<HTML>');
    WriteLn(myFile, '   <HEAD>');
    WriteLn(myFile, '   <meta http-equiv="Content-Type" ' + 'content="text/html; charset=UTF-8">');
    WriteLn(myFile, '   </HEAD>');
    WriteLn(myFile, '   <BODY>');
    SubNotesText := sqNotes.FieldByName('NotesTitle').AsString;
    SubNotesText := StringReplace(SubNotesText, '&', '&amp;', [rfReplaceAll]);
    SubNotesText := StringReplace(SubNotesText, '<', '&lt;', [rfReplaceAll]);
    SubNotesText := StringReplace(SubNotesText, '>', '&gt;', [rfReplaceAll]);
    WriteLn(myFile, '<H1>' + SubNotesText + '</H1>');
    // Add date if not for print
    if Sender <> miNotesPrint then
    begin
      WriteLn(myFile, '<H3>' + sqNotes.FieldByName('NotesDate').AsString + '</H3>');
    end;
    // Add text non encrypted
    if sqNotes.FieldByName('NotesCheckPwd').AsString = '' then
      SubNotesText := sqNotes.FieldByName('NotesText').AsString
    // Add text encrypted: if the password is not set, this function is disabled
    else
    begin
      dcAES.InitStr(edPassword.Text, TDCP_sha1);
      SubNotesText :=
        dcAES.DecryptString(sqNotes.FieldByName('NotesText').AsWideString);
    end;
    // Set font size in html format (from 1 to 7)
    for FnSize := 6 to 72 do
    begin
      if FnSize < 7 then
        HTMLFnSize := 1
      else if FnSize < 10 then
        HTMLFnSize := 2
      else if FnSize < 13 then
        HTMLFnSize := 3
      else if FnSize < 17 then
        HTMLFnSize := 4
      else if FnSize < 21 then
        HTMLFnSize := 5
      else if FnSize < 25 then
        HTMLFnSize := 6
      else
        HTMLFnSize := 7;
      SubNotesText := StringReplace(SubNotesText, 'size="' + IntToStr(FnSize) + '" face="', 'size="' + IntToStr(HTMLFnSize) + '" face="', [rfIgnoreCase, rfReplaceAll]);
    end;
    // Set the tags for lists
    IniIndent := 0;
    EndIndent := 0;
    while UTF8Pos('<blockquote>', SubNotesText) > 0 do
    begin
      flNumList := False;
      IniIndent := UTF8Pos('<blockquote>', SubNotesText) + 13;
      EndIndent := UTF8Pos('</blockquote>', SubNotesText) - 1;
      while IniIndent < EndIndent do
      begin
        if ((UTF8Copy(SubNotesText, IniIndent, 3) = '>1.') or (UTF8Copy(SubNotesText, IniIndent, 3) = '>A.') or (UTF8Copy(SubNotesText, IniIndent, 3) = '>a.')) then
          flNumList := True;
        if SubNotesText[IniIndent] = LineEnding then
          SubNotesText[IniIndent] := #3; // A code to be replaced with <il>
        Inc(IniIndent);
      end;
      SubNotesText := StringReplace(SubNotesText, #3, '</li><li>', [rfReplaceAll]);
      if flNumList = True then
      begin
        SubNotesText := StringReplace(SubNotesText, '<blockquote>', '<p align=left><ol><li>', []);
        SubNotesText := StringReplace(SubNotesText, '</blockquote>', '</li></ol></p>', []);
      end
      else
      begin
        SubNotesText := StringReplace(SubNotesText, '<blockquote>', '<p align=left><ul><li>', []);
        SubNotesText := StringReplace(SubNotesText, '</blockquote>', '</li></ul></p>', []);
      end;
    end;
    // A list has been processed
    if IniIndent > 0 then
    begin
      // Delete numbers, letters and bullets
      for NumList := 1 to 100 do
      begin
        SubNotesText := StringReplace(SubNotesText, IntToStr(NumList) + '.' + #9, '', [rfReplaceAll]);
      end;
      for NumList := 65 to 90 do
      begin
        SubNotesText :=
          StringReplace(SubNotesText, char(NumList) + '.' + #9, '', [rfReplaceAll]);
      end;
      for NumList := 97 to 122 do
      begin
        SubNotesText :=
          StringReplace(SubNotesText, char(NumList) + '.' + #9, '', [rfReplaceAll]);
      end;
      SubNotesText := StringReplace(SubNotesText, '•' + #9, '', [rfReplaceAll]);
      // Remove alignment tags within a list to avoid empty rows
      // Alignment is always left
      SubNotesText := StringReplace(SubNotesText, '<li><p align=left>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li><p align=center>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li><p align=right>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li><p align=justify>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li></p><p align=left>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li></p><p align=center>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li></p><p align=right>', '<li>', [rfReplaceAll]);
      SubNotesText := StringReplace(SubNotesText, '<li></p><p align=justify>', '<li>', [rfReplaceAll]);
    end;
    SubNotesText := StringReplace(SubNotesText, LineEnding, '<p>', [rfReplaceAll]);
    SubNotesText := StringReplace(SubNotesText, #5, '&lt;', [rfReplaceAll]);
    SubNotesText := StringReplace(SubNotesText, #6, '&gt;', [rfReplaceAll]);
    WriteLn(myFile, SubNotesText);
    // Save activities
    if sqNotes.FieldByName('NotesActivities').AsString <> '' then
    begin
      WriteLn(myFile, '<hr width="30%" size=1 color="black">');
      WriteLn(myFile, '<p><table>');
      WriteLn(myFile, '<tr>');
      WriteLn(myFile, '  <th><font size="1">' + lbWBS + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbState + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbActivity + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbStartDate + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbEndDate + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbDuration + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbResources + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbPriority + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbCompletion + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbCost + '</font></th>');
      WriteLn(myFile, '  <th><font size="1">' + lbNotes + '</font></th>');
      WriteLn(myFile, '</tr>');
      SubActText := sqNotes.FieldByName('NotesActivities').AsString;
      WriteLn(myFile, '<tr>');
      // Remove code column of the first row
      SubActText := UTF8Copy(SubActText, UTF8Pos(#9, SubActText) + 1, UTF8Length(SubActText));
      PosName := 0;
      stSpcInd := '';
      while UTF8Length(SubActText) > 0 do
      begin
        if Copy(SubActText, 1, Pos(#9, SubActText) - 1) <> '' then
        begin
          // The field is activity name, so is indented
          if PosName = 2 then
            WriteLn(myFile, '<td><font size="1">' + stSpcInd + Copy(SubActText, 1, Pos(#9, SubActText) - 1) + '</font></td>')
          // The field is Completion
          else if PosName = 8 then
            WriteLn(myFile, '<td><font size="1">' + Copy(SubActText, 1, Pos(#9, SubActText) - 1) + ' %</font></td>')
          // The field is Cost
          else if PosName = 9 then
            WriteLn(myFile, '<td><font size="1">' + Copy(SubActText, 1, Pos(#9, SubActText) - 1) + ' ' + CurrencyString + '</font></td>')
          else
            WriteLn(myFile, '<td><font size="1">' + Copy(SubActText, 1, Pos(#9, SubActText) - 1) + '</font></td>');
        end
        else
          WriteLn(myFile, '<td><font size="1">' + '</font></td>');
        SubActText := UTF8Copy(SubActText, UTF8Pos(#9, SubActText) + 1, UTF8Length(SubActText));
        Inc(PosName);
        if UTF8Copy(SubActText, 1, 1) = LineEnding then
        begin
          SubActText := UTF8Copy(SubActText, 2, UTF8Length(SubActText));
          // Get the code and remove code column
          if CheckNumbers(UTF8Copy(SubActText, 2, UTF8Pos(#9, SubActText) - 2)) = True then
          begin
            SpIndent := StrToInt(UTF8Copy(SubActText, 2, UTF8Pos(#9, SubActText) - 2));
            stSpcInd := '';
            for i := 1 to SpIndent do
              stSpcInd := stSpcInd + '&nbsp;';
          end;
          SubActText :=
            UTF8Copy(SubActText, UTF8Pos(#9, SubActText) + 1, UTF8Length(SubActText));
          WriteLn(myFile, '</tr>' + LineEnding + '<tr>');
          PosName := 0;
        end;
      end;
      WriteLn(myFile, '</tr>');
      WriteLn(myFile, '</table>');
    end;
    WriteLn(myFile, '   </BODY>');
    WriteLn(myFile, '</HTML>');
    CloseFile(myFile);
    // Copy images
    AttReadDir := ExtractFileNameWithoutExt(sqNotes.FileName);
    ImgNum := 0;
    if DirectoryExistsUTF8(AttReadDir) = True then
    begin
      AttWriteDir := GetTempDir + DirectorySeparator;
      while FileExistsUTF8(AttReadDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', ImgNum) + '.jpeg') = True do
      begin
        CopyFile(AttReadDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', ImgNum) + '.jpeg',
          AttWriteDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', ImgNum) + '.jpeg');
        Inc(ImgNum);
      end;
    end;
  except
    Screen.Cursor := crDefault;
    MessageDlg(msg020, mtWarning, [mbOK], 0);
  end;
  if Sender = miNotesPrint then
  begin
    if ImgNum > 0 then
    begin
      MessageDlg(msg071, mtInformation, [mbOK], 0);
      Abort;
    end;
    MyPrinter := Printer;
    if MyPrinter.Printers.Count > 0 then
      try
        try
          myHTMLList := TStringList.Create;
          myHTMLList.LoadFromFile(GetTempDir + DirectorySeparator + 'MyNotexFile.html');
          ipPrint.SetHtmlFromStr(myHTMLList.Text);
          ipPrint.Print(1, ipPrint.GetPrintPageCount);
        except
          MessageDlg(msg072, mtWarning, [mbOK], 0);
        end;
      finally
        Screen.Cursor := crDefault;
        myHTMLList.Free
      end;
  end
  // Create process and run word processor
  else
    try
      if tbOpenNote.Tag = 0 then
        stCommand := 'ooffice -writer -n ' + GetTempDir + DirectorySeparator + 'MyNotexFile.html'
      else if tbOpenNote.Tag = 1 then
        stCommand := 'libreoffice --writer -n ' + GetTempDir + DirectorySeparator + 'MyNotexFile.html'
      else if tbOpenNote.Tag = 2 then
        stCommand := 'xdg-open ' + GetTempDir + DirectorySeparator + 'MyNotexFile.html';
      Proc := TProcess.Create(nil);
      Proc.CommandLine := stCommand;
      Proc.Execute;
      Proc.Free;
    except
      Screen.Cursor := crDefault;
      MessageDlg(msg032, mtWarning, [mbOK], 0);
    end;
  Screen.Cursor := crDefault;
end;

procedure TfmMain.tbColorFontChangeClick(Sender: TObject);
// Works also for tbBackColorFontChange
var
  fp: TFontParams;
  idxStart, idxLength: integer;
begin
  // Change font color for selection
  // There is a selection
  if dbText.SelLength > 0 then
  begin
    idxStart := dbText.SelStart;
    idxLength := dbText.SelLength;
  end
  // There is no selection: the formatting is applied to the current word
  else
  begin
    idxStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, True, True);
    idxLength := dbText.GetWordParagraphStartEnd(dbText.SelStart, True, False) - idxStart;
  end;
  if Sender = tbColorFontChange then
  begin
    fp.Color := fdColorFormatting.Color;
    fp.Changed := [fiColor];
  end
  else
  begin
    fp.BackColor := fdBackColorFormatting.Color;
    fp.Changed := [fiBackcolor];
  end;
  dbText.SetTextAttributes(idxStart, idxLength, fp);
  // Put the Notes dataset in Edit and set the flag for user change of RichMemo
  EditNotesDataset;
end;

procedure TfmMain.pmChangeFontColorClick(Sender: TObject);
begin
  // Set font color for formatting
  if fdColorFormatting.Execute = True then
    tbColorFontChangeClick(tbColorFontChange);
end;

procedure TfmMain.pmChangeFontBackColorClick(Sender: TObject);
begin
  // Set background color for formatting
  if fdBackColorFormatting.Execute = True then
    tbColorFontChangeClick(tbBackColorFontChange);
end;

procedure TfmMain.SetFontFormat(Sender: TObject);
var
  fp, fpOld: TFontParams;
  idxStart, idxLength: integer;
begin
  // Set character formatting
  // Paragraph selection
  if Sender = tbAlignLeft then
    tbAlignLeft.Down := True;
  if Sender = tbAlignCenter then
    tbAlignCenter.Down := True;
  if Sender = tbAlignRight then
    tbAlignRight.Down := True;
  if Sender = tbAlignFill then
    tbAlignFill.Down := True;
  if ((Sender = tbAlignLeft) or (Sender = tbAlignCenter) or (Sender = tbAlignRight) or (Sender = tbAlignFill) or (Sender = tbAlignIndent)) then
  begin
    idxStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True);
    idxLength := dbText.GetWordParagraphStartEnd(dbText.SelStart + dbText.SelLength, False, False) - idxStart;
  end
  // Word selection or user selection
  else
  begin
    // There is a selection
    if dbText.SelLength > 0 then
    begin
      idxStart := dbText.SelStart;
      idxLength := dbText.SelLength;
    end
    // There is no selection
    else
    begin
      idxStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, True, True);
      idxLength := dbText.GetWordParagraphStartEnd(dbText.SelStart, True, False) - idxStart;
    end;
  end;
  // Restore default font
  if ((Sender = tbFontRestore) or (Sender = pmRestoreFont)) then
  begin
    fp.Color := clBlack;
    fp.BackColor := clWhite;
    fp.Indented := DefaultIndent;
    fp.Name := DefFontName;
    // DefFontSize could be = 0
    if DefFontSize < 3 then
      fp.Size := 11
    else
      fp.Size := DefFontSize + ZoomFontSize;
    fp.Style := [];
    fp.Changed := [fiName, fiSize, fiColor, fiBold, fiItalic, fiUnderline, fiIndented, fiStrike, fiBackcolor];
    dbText.SetTextAttributes(idxStart, idxLength, fp);
  end
  else
  begin
    // Set bold
    if Sender = tbFontBold then
    begin
      fp.Changed := [fiBold];
      if tbFontBold.Down = True then
        fp.Style := fp.Style + [fsBold]
      else
        fp.Style := [];
    end
    // Set italic
    else if Sender = tbFontItalic then
    begin
      fp.Changed := [fiItalic];
      if tbFontItalic.Down = True then
        fp.Style := fp.Style + [fsItalic]
      else
        fp.Style := [];
    end
    // Set underline
    else if Sender = tbFontUnderline then
    begin
      fp.Changed := [fiUnderline];
      if tbFontUnderline.Down = True then
        fp.Style := fp.Style + [fsUnderline]
      else
        fp.Style := [];
    end
    // Set strike
    else if Sender = tbFontStrike then
    begin
      fp.Changed := [fiStrike];
      if tbFontStrike.Down = True then
        fp.Style := fp.Style + [fsStrikeOut]
      else
        fp.Style := [];
    end
    // Set left
    else if Sender = tbAlignLeft then
    begin
      fp.Changed := [fiAlignment];
      fp.Alignment := [trLeft];
      tbAlignCenter.Down := False;
      tbAlignRight.Down := False;
      tbAlignFill.Down := False;
    end
    // Set center
    else if Sender = tbAlignCenter then
    begin
      fp.Changed := [fiAlignment];
      fp.Alignment := [trCenter];
      tbAlignLeft.Down := False;
      tbAlignRight.Down := False;
      tbAlignFill.Down := False;
    end
    // Set right
    else if Sender = tbAlignRight then
    begin
      fp.Changed := [fiAlignment];
      fp.Alignment := [trRight];
      tbAlignLeft.Down := False;
      tbAlignCenter.Down := False;
      tbAlignFill.Down := False;
    end
    // Set fill
    else if Sender = tbAlignFill then
    begin
      fp.Changed := [fiAlignment];
      fp.Alignment := [trJustified];
      tbAlignLeft.Down := False;
      tbAlignCenter.Down := False;
      tbAlignRight.Down := False;
    end
    // Set indent
    else if Sender = tbAlignIndent then
    begin
      fp.Changed := [fiIndented];
      dbText.GetTextAttributes(dbText.SelStart, fpOld);
      if fpOld.Indented <= DefaultIndent then
      begin
        fp.Indented := WidthInden;
        tbAlignIndent.Down := True;
      end
      else
      begin
        fp.Indented := DefaultIndent;
        tbAlignIndent.Down := False;
      end;
    end;
    dbText.SetTextAttributes(idxStart, idxLength, fp);
  end;
  EditNotesDataset;
end;

procedure TfmMain.SetHeadings(Sender: TObject);
var
  SelStart, SelEnd: integer;
  fp: TFontParams;
begin
  // Set headings
  SelStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True);
  SelEnd := dbText.GetWordParagraphStartEnd(dbText.SelStart + dbText.SelLength, False, False);
  if (Sender = pmHeading1) then
  begin
    fp.Size := DefFontSize + ZoomFontSize + 2;
    fp.Style := [fsBold];
    fp.Alignment := [trCenter];
    fp.Color := clBlack;
    fp.Changed := [fiSize, fiColor, fiBold, fiItalic, fiUnderline, fiStrike, fiAlignment];
  end;
  if (Sender = pmHeading2) then
  begin
    fp.Size := DefFontSize + ZoomFontSize;
    fp.Style := [fsBold];
    fp.Alignment := [trLeft];
    fp.Color := clBlack;
    fp.Changed := [fiSize, fiColor, fiBold, fiItalic, fiUnderline, fiStrike, fiAlignment];
  end;
  if (Sender = pmHeading3) then
  begin
    fp.Size := DefFontSize + ZoomFontSize;
    fp.Style := [fsItalic];
    fp.Alignment := [trLeft];
    fp.Color := clBlack;
    fp.BackColor := clWhite;
    fp.Changed := [fiSize, fiColor, fiBold, fiItalic, fiUnderline, fiStrike, fiAlignment];
  end;
  if (Sender = pmRestoreFont) then
  begin
    fp.Name := DefFontName;
    fp.Size := DefFontSize + ZoomFontSize;
    fp.Style := [];
    fp.Color := clBlack;
    fp.BackColor := clWhite;
    fp.Indented := DefaultIndent;
    fp.Alignment := [trLeft];
    fp.Changed := [fiName, fiSize, fiColor, fiBold, fiItalic, fiUnderline, fiStrike, fiIndented, fiAlignment, fiBackcolor];
  end;
  dbText.SetTextAttributes(SelStart, SelEnd - SelStart + 1, fp);
  EditNotesDataset;
  ActivateFormatIcons;
end;

procedure TfmMain.pmBullClick(Sender: TObject);
var
  fp: TFontParams;
  idxStart, idxLength: integer;
begin
  // Set list with bullets
  // It work for all pmBull menu items
  idxStart := dbText.GetWordParagraphStartEnd(dbText.SelStart, False, True);
  idxLength := dbText.GetWordParagraphStartEnd(dbText.SelStart, False, False) - idxStart;
  fp.Changed := [fiIndented];
  fp.Indented := WidthInden;
  dbText.SetTextAttributes(idxStart, idxLength, fp);
  if Sender = pmBull1 then
    dbText.ListNumber(dbText.SelStart, WidthInden, '*')
  else if Sender = pmBull2 then
    dbText.ListNumber(dbText.SelStart, WidthInden, '1')
  else if Sender = pmBull3 then
    dbText.ListNumber(dbText.SelStart, WidthInden, 'A')
  else if Sender = pmBull4 then
    dbText.ListNumber(dbText.SelStart, WidthInden, 'a')
  else if Sender = pmNoBull then
    dbText.ListNumber(dbText.SelStart, 0, 'X');
  EditNotesDataset;
end;

procedure TfmMain.dbTextClick(Sender: TObject);
begin
  // On caret move activate or deactivate format icons
  ActivateFormatIcons;
end;

procedure TfmMain.ActivateFormatIcons;
var
  fp: TFontParams;
begin
  // Activate and deactivate format icons
  dbText.GetTextAttributes(dbText.SelStart, fp);
  // Bold icon
  if fsBold in fp.Style then
    tbFontBold.Down := True
  else
    tbFontBold.Down := False;
  // Italic icon
  if fsItalic in fp.Style then
    tbFontItalic.Down := True
  else
    tbFontItalic.Down := False;
  // Underline icon
  if fsUnderline in fp.Style then
    tbFontUnderline.Down := True
  else
    tbFontUnderline.Down := False;
  // Strike icon
  if fsStrikeOut in fp.Style then
    tbFontStrike.Down := True
  else
    tbFontStrike.Down := False;
  // Left icon
  if trLeft in fp.Alignment then
  begin
    tbAlignLeft.Down := True;
    tbAlignCenter.Down := False;
    tbAlignRight.Down := False;
    tbAlignFill.Down := False;
  end;
  // Center icon
  if trCenter in fp.Alignment then
  begin
    tbAlignLeft.Down := False;
    tbAlignCenter.Down := True;
    tbAlignRight.Down := False;
    tbAlignFill.Down := False;
  end;
  // Right icon
  if trRight in fp.Alignment then
  begin
    tbAlignLeft.Down := False;
    tbAlignCenter.Down := False;
    tbAlignRight.Down := True;
    tbAlignFill.Down := False;
  end;
  // Fill icon
  if trJustified in fp.Alignment then
  begin
    tbAlignLeft.Down := False;
    tbAlignCenter.Down := False;
    tbAlignRight.Down := False;
    tbAlignFill.Down := True;
  end;
  // Indent icon
  if fp.Indented > DefaultIndent then
    tbAlignIndent.Down := True
  else
    tbAlignIndent.Down := False;
end;

procedure TfmMain.pmTextSelectAllClick(Sender: TObject);
begin
  //Select all the text
  dbText.SelectAll;
end;

function TfmMain.CheckLink: integer;
var
  SelStart: integer;
  slPosImg: TStringList;
  i, iCorr: integer;
begin
  // Check if the current word is a link
  Result := -1;
  SelStart := dbText.SelStart;
  try
    // Get the possibile picture position in the note
    // slPosImg is created in the TRichMemo component code,
    // in the GetImagePosInText event
    slPosImg := dbText.GetImagePosInText;
    iCorr := 0;
    if slPosImg.Count > 0 then
    begin
      for i := 0 to slPosImg.Count - 1 do
      begin
        if StrToInt(slPosImg[i]) < SelStart then
          iCorr := iCorr - 1;
      end;
      SelStart := SelStart + iCorr;
    end;
  finally
    slPosImg.Free;
  end;
  while ((UTF8Copy(dbText.Text, SelStart, 1) <> ' ') and (UTF8Copy(dbText.Text, SelStart, 1) <> #9) and (UTF8Copy(dbText.Text, SelStart, 1) <> LineEnding) and (SelStart > 0)) do
    SelStart := SelStart - 1;
  if ((UTF8Copy(dbText.Text, SelStart + 1, 1) = '(') or (UTF8Copy(dbText.Text, SelStart + 1, 1) = '{') or (UTF8Copy(dbText.Text, SelStart + 1, 1) = '[')) then
    SelStart := SelStart + 1;
  if (((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 7)) = 'http://') and (UTF8Copy(dbText.Text, SelStart + 8, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 7)) or ((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 8)) = 'https://') and (UTF8Copy(dbText.Text, SelStart + 9, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 8)) or ((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 4)) = 'www.') and (UTF8Copy(dbText.Text, SelStart + 5, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 4)) or ((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 7)) = 'mailto:') and (UTF8Copy(dbText.Text, SelStart + 8, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 7)) or
    ((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 7)) = 'file://') and (UTF8Copy(dbText.Text, SelStart + 8, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 7)) or ((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 6)) = 'mnt://') and (UTF8Copy(dbText.Text, SelStart + 7, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 6)) or ((UTF8LowerCase(UTF8Copy(dbText.Text, SelStart + 1, 1)) = '#') and (UTF8Copy(dbText.Text, SelStart + 2, 1) <> LineEnding) and (UTF8Length(dbText.Text) > SelStart + 1))) then
    Result := SelStart;
end;

procedure TfmMain.MakeLink(SelStart: integer);
var
  SelEnd, i, iCorr: integer;
  fp: TFontParams;
  slPosImg: TStringList;
begin
  // Make current word a link
  if SelStart > -1 then
  begin
    SelEnd := dbText.SelStart - 1;
    if ((UTF8Copy(dbText.Text, SelEnd + 1, 1) = '.') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = ',') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = ';') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = ':') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = '?') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = '!')) then
      SelEnd := SelEnd - 1;
    if ((UTF8Copy(dbText.Text, SelEnd + 1, 1) = ')') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = ']') or (UTF8Copy(dbText.Text, SelEnd + 1, 1) = '}')) then
      SelEnd := SelEnd - 1;
    dbText.GetTextAttributes(SelStart, fp);
    fp.Color := clBlue;
    if UTF8Copy(dbText.Text, SelStart + 1, 1) <> '#' then
      fp.Style := [fsUnderline];
    fp.Changed := [fiColor, fiUnderline];
    try
      // Get the possibile picture position in the note
      // slPosImg is created in the TRichMemo component code,
      // in the GetImagePosInText event
      slPosImg := dbText.GetImagePosInText;
      iCorr := 0;
      if slPosImg.Count > 0 then
      begin
        for i := 0 to slPosImg.Count - 1 do
        begin
          if StrToInt(slPosImg[i]) < SelStart then
            iCorr := iCorr - 1;
        end;
        SelStart := SelStart - iCorr;
      end;
    finally
      slPosImg.Free;
    end;
    dbText.SetTextAttributes(SelStart, SelEnd - SelStart + 1, fp);
  end;
end;

procedure TfmMain.sbStatusBarDrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel; const Rect: TRect);
begin
  with sbStatusBar.Canvas do
  begin
    case Panel.Index of
      0, 1:
      begin
        Brush.Color := sbStatusBar.Color;
        Font.Style := [fsBold];
        if IncMinute(Time, 5) > fmSetAlarm.tpAlarm.Time then
          Font.Color := clRed
        else if IncMinute(Time, 15) > fmSetAlarm.tpAlarm.Time then
          Font.Color := clMaroon
        else
          Font.Color := clGreen;
      end;
    end;
    FillRect(Rect);
    TextRect(Rect, 0, 1, Panel.Text);
  end;
end;

// *****************************************************************************
// ************************* PASSWORD PROCEDURES *******************************
// *****************************************************************************

procedure TfmMain.ShowPasswordInput;
begin
  // Show the password fields
  // If the cursor has moved in a locked record
  // but it is going to move elsewhere, do nothing
  if IsTextToLoad = False then
    Exit;
  DisableShowTextOnly;
  miNotesShowOnlyText.Enabled := False;
  edPassword.Text := '';
  lbPwdBottom.Visible := False;
  pnPassword.Visible := True;
  dbText.Visible := False;
  dbText.Clear;
  dbText.ReadOnly := True;
  if edPassword.Visible = True then
    edPassword.SetFocus;
  // To activate menu items and buttons
  dsNotesDataChange(nil, nil);
end;

procedure TfmMain.HidePasswordInput;
begin
  // Hide the password fields
  dbText.ReadOnly := False;
  dbText.Visible := True;
  pnPassword.Visible := False;
  miNotesShowOnlyText.Enabled := True;
  // To activate menu items and buttons
  dsNotesDataChange(nil, nil);
  // edPassword.Text must still keep the password, so is not cleared
end;

procedure TfmMain.CheckPasswordInput;
begin
  // Check if the password is correct
  dcAES.InitStr(edPassword.Text, TDCP_sha1);
  // The password is correct
  if dcAES.EncryptString(PwdCheckString) = sqNotes.FieldByName('NotesCheckPwd').AsString then
  begin
    HidePasswordInput;
    LoadRichMemo;
  end
  else
  begin
    lbPwdBottom.Visible := True;
    edPassword.SelectAll;
  end;
end;

procedure TfmMain.edPasswordKeyPress(Sender: TObject; var Key: char);
begin
  // Try to load text
  if Key = #13 then
    CheckPasswordInput
  else
    lbPwdBottom.Visible := False;
end;

// *****************************************************************************
// ***************** CONVERT FROM TOMBOY/GNOTE PROCEDURES **********************
// *****************************************************************************

procedure TfmMain.ConvertFromTomboyGnote(DataPath: string);
var
  myFileListBox: TFileListBox;
  sqTomGnoteSubjects, sqTomGnoteNotes: TSqlite3Dataset;
  stText: TStringList;
  stSubject: string;
  i, NtAdded, NtModified: integer;
  myGUID: TGUID;
begin
  // Convert data from Tomboy or GNote notes
  try
    Cursor := crHourGlass;
    Application.ProcessMessages;
    // Create the file list
    myFileListBox := TFileListBox.Create(Self);
    myFileListBox.Directory := DataPath;
    myFileListBox.Mask := '*.note';
    // Create a Subject and Note dataset
    sqTomGnoteSubjects := TSqlite3Dataset.Create(Self);
    sqTomGnoteSubjects.FileName := sqSubjects.FileName;
    sqTomGnoteSubjects.TableName := 'Subjects';
    sqTomGnoteSubjects.PrimaryKey := 'IDSubjects';
    sqTomGnoteSubjects.AutoIncrementKey := True;
    sqTomGnoteSubjects.SQL := 'Select * from Subjects';
    sqTomGnoteSubjects.Open;
    sqTomGnoteNotes := TSqlite3Dataset.Create(Self);
    sqTomGnoteNotes.FileName := sqSubjects.FileName;
    sqTomGnoteNotes.TableName := 'Notes';
    sqTomGnoteNotes.PrimaryKey := 'IDNotes';
    sqTomGnoteNotes.SQL := 'Select * from Notes';
    sqTomGnoteNotes.Open;
    // Create a stringlist to get the text of the note
    stText := TStringList.Create;
    // Set the counters
    NtAdded := 0;
    NtModified := 0;
    // Check each file of Tomboy/Gnote notes
    for i := 0 to myFileListBox.Items.Count - 1 do
    begin
      // The note is not present in the archive in use
      if sqTomGnoteNotes.Locate('NotesUID', ExtractFileNameOnly(myFileListBox.Items[i]), [loCaseInsensitive]) = False then
      begin
        stText.LoadFromFile(myFileListBox.Directory + DirectorySeparator + myFileListBox.Items[i]);
        stSubject := GetTomboyGnoteSubject(stText.Text);
        // If the subject does not exist, create the generic one
        if stSubject = '' then
          stSubject := msg050;
        //The subject is not present in the archive
        if sqTomGnoteSubjects.Locate('SubjectsName', stSubject, [loCaseInsensitive]) = False then
        begin
          sqTomGnoteSubjects.Append;
          sqTomGnoteSubjects.FieldByName('SubjectsName').AsString := stSubject;
          CreateGUID(myGUID);
          sqTomGnoteSubjects.FieldByName('SubjectsUID').AsString :=
            Copy(GUIDToString(myGUID), 2, Length(GUIDToString(myGUID)) - 2);
          sqTomGnoteSubjects.FieldByName('SubjectsDTMod').AsDateTime := Now;
          sqTomGnoteSubjects.Post;
          sqTomGnoteSubjects.ApplyUpdates;
        end;
        // Add the note
        sqTomGnoteNotes.Append;
        sqTomGnoteNotes.FieldByName('ID_Subjects').AsInteger :=
          sqTomGnoteSubjects.FieldByName('IDSubjects').AsInteger;
        sqTomGnoteNotes.FieldByName('NotesTitle').AsString :=
          GetTomboyGnoteTitle(stText.Text);
        sqTomGnoteNotes.FieldByName('NotesDate').AsDateTime :=
          GetTomboyGnoteCreateDate(stText.Text);
        sqTomGnoteNotes.FieldByName('NotesText').AsString :=
          GetTomboyGnoteText(stText.Text);
        sqTomGnoteNotes.FieldByName('NotesUID').AsString :=
          UpperCase(ExtractFileNameOnly(myFileListBox.Items[i]));
        sqTomGnoteNotes.FieldByName('NotesDTMod').AsDateTime :=
          GetTomboyGnoteLastChangeDate(stText.Text);
        sqTomGnoteNotes.Post;
        sqTomGnoteNotes.ApplyUpdates;
        NtAdded := NtAdded + 1;
      end
      // The note is present in the archive in use: check if it must be updated
      else
      begin
        stText.LoadFromFile(myFileListBox.Directory + DirectorySeparator + myFileListBox.Items[i]);
        // Add a second to inner datetime to make sure the comparison is exact
        if sqTomGnoteNotes.FieldByName('NotesDTMod').AsDateTime + StrToTime('00:00:01') < GetTomboyGnoteLastChangeDate(stText.Text) then
        begin
          stText.LoadFromFile(myFileListBox.Directory + DirectorySeparator + myFileListBox.Items[i]);
          sqTomGnoteNotes.Edit;
          sqTomGnoteNotes.FieldByName('NotesTitle').AsString :=
            GetTomboyGnoteTitle(stText.Text);
          sqTomGnoteNotes.FieldByName('NotesText').AsString :=
            GetTomboyGnoteText(stText.Text);
          sqTomGnoteNotes.FieldByName('NotesDTMod').AsDateTime :=
            GetTomboyGnoteLastChangeDate(stText.Text);
          sqTomGnoteNotes.Post;
          sqTomGnoteNotes.ApplyUpdates;
          NtModified := NtModified + 1;
        end;
      end;
      Application.ProcessMessages;
    end;
    // Update data
    miFileUpdateClick(nil);
    sbStatusBar.Panels[0].Text :=
      ' ' + msg051 + ' ' + IntToStr(NtAdded) + ' - ' + msg052 + ' ' + IntToStr(NtModified) + '.';
  finally
    myFileListBox.Free;
    stText.Free;
    sqTomGnoteSubjects.Free;
    sqTomGnoteNotes.Free;
    Cursor := crDefault;
  end;
end;

function TfmMain.GetTomboyGnoteCreateDate(stText: string): TDate;
begin
  // Get creation date
  Result := 0; //Just to initialize the variable
  try
    stText := UTF8Copy(stText, UTF8Pos('<create-date>', stText) + 13, UTF8Pos('</create-date>', stText) - UTF8Pos('<create-date>', stText) - 13);
    if FDate.ShortDateFormat = 'dd-mm-yyyy' then
      stText := UTF8Copy(stText, 9, 2) + FDate.DateSeparator + UTF8Copy(stText, 6, 2) + FDate.DateSeparator + UTF8Copy(stText, 1, 4)
    else if FDate.ShortDateFormat = 'yyyy-mm-dd' then
      stText := UTF8Copy(stText, 1, 4) + FDate.DateSeparator + UTF8Copy(stText, 6, 2) + FDate.DateSeparator + UTF8Copy(stText, 9, 2)
    else if FDate.ShortDateFormat = 'mm-dd-yyyy' then
      stText := UTF8Copy(stText, 6, 2) + FDate.DateSeparator + UTF8Copy(stText, 9, 2) + FDate.DateSeparator + UTF8Copy(stText, 1, 4);
    Result := StrToDateTime(stText, FDate);
  except
    Result := 0;
  end;
end;

function TfmMain.GetTomboyGnoteLastChangeDate(stText: string): TDateTime;
begin
  // Get last change date
  Result := 0; //Just to initialize the variable
  try
    stText := UTF8Copy(stText, UTF8Pos('<last-change-date>', stText) + 18, UTF8Pos('</last-change-date>', stText) - UTF8Pos('<last-change-date>', stText) - 18);
    if FDate.ShortDateFormat = 'dd-mm-yyyy' then
      stText := UTF8Copy(stText, 9, 2) + FDate.DateSeparator + UTF8Copy(stText, 6, 2) + FDate.DateSeparator + UTF8Copy(stText, 1, 4)
    else if FDate.ShortDateFormat = 'yyyy-mm-dd' then
      stText := UTF8Copy(stText, 1, 4) + FDate.DateSeparator + UTF8Copy(stText, 6, 2) + FDate.DateSeparator + UTF8Copy(stText, 9, 2)
    else if FDate.ShortDateFormat = 'mm-dd-yyyy' then
      stText := UTF8Copy(stText, 6, 2) + FDate.DateSeparator + UTF8Copy(stText, 9, 2) + FDate.DateSeparator + UTF8Copy(stText, 1, 4);
    Result := StrToDateTime(stText, FDate);
  except
    Result := 0;
  end;
end;

function TfmMain.GetTomboyGnoteTitle(stText: string): string;
begin
  // Get title
  Result := '';
  stText := UTF8Copy(stText, UTF8Pos('<title>', stText) + 7, UTF8Pos('</title>', stText) - UTF8Pos('<title>', stText) - 7);
  Result := stText;
end;

function TfmMain.GetTomboyGnoteSubject(stText: string): string;
begin
  // Get subject
  Result := '';
  if UTF8Pos('<tag>system:notebook:', stText) > 0 then
  begin
    stText := UTF8Copy(stText, UTF8Pos('<tag>system:notebook:', stText) + 21, UTF8Pos('</tag>', stText) - UTF8Pos('<tag>system:notebook:', stText) - 21);
    Result := stText;
  end;
end;

function TfmMain.GetTomboyGnoteText(stText: string): string;
var
  StOutput: string;
  flCopy: boolean;
  i: integer;
begin
  // Get text
  Result := '';
  stText := UTF8Copy(stText, UTF8Pos('<note-content version', stText) + 21, UTF8Pos('</note-content>', stText) - UTF8Pos('<note-content version', stText) - 21);
  stText := '<' + stText; // After <note-content version there is the number of version
  // Correct tags
  // To avoid that ‹ and › characters are used in the text
  stText := StringReplace(stText, '‹', '«', [rfReplaceAll]);
  stText := StringReplace(stText, '›', '»', [rfReplaceAll]);
  stText := StringReplace(stText, '<bold>', '‹b›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</bold>', '‹/b›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<italic>', '‹i›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</italic>', '‹/i›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<italic>', '‹i›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</italic>', '‹/i›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<highlight>', '‹u›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</highlight>', '‹/u›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<strikethrough>', '‹strike›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</strikethrough>', '‹/strike›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<list-item', '•'#9'<', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<list>', '‹blockquote›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</list>', '‹/blockquote›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<size:small>', '‹/font›‹font color="#000000" size="' + IntToStr(DefFontSize - 2) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<size:large>', '‹/font›‹font color="#000000" size="' + IntToStr(DefFontSize + 2) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<size:huge>', '‹/font›‹font color="#000000" size="' + IntToStr(DefFontSize + 4) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</size:small>', '‹/font›‹font color="#000000" size="' + IntToStr(DefFontSize) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</size:large>', '‹/font›‹font color="#000000" size="' + IntToStr(DefFontSize) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</size:huge>', '‹/font›‹font color="#000000" size="' + IntToStr(DefFontSize) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '<link:url>', '‹/font›‹font color="#0000FF" size="' + IntToStr(DefFontSize) + '" face="' + DefFontName + '"›‹u›', [rfReplaceAll, rfIgnoreCase]);
  stText := StringReplace(stText, '</link:url>', '‹/font›‹/u›‹font color="#000000" size="' + IntToStr(DefFontSize) + '" face="' + DefFontName + '"›', [rfReplaceAll, rfIgnoreCase]);
  // Clear other possible uncorrected HTML tags
  flCopy := True;
  stOutput := '';
  for i := 1 to Length(stText) do
  begin
    if stText[i] = '<' then
      flCopy := False
    else if stText[i] = '>' then
      flCopy := True
    else if flCopy = True then
      stOutput := stOutput + stText[i];
    Application.ProcessMessages;
  end;
  stOutput := StringReplace(stOutput, '‹', '<', [rfReplaceAll]);
  stOutput := StringReplace(stOutput, '›', '>', [rfReplaceAll]);
  Result := '<font color="#000000" size="' + IntToStr(DefFontSize) + '" face="' + DefFontName + '">' + stOutput;
end;


// *****************************************************************************
// ************************* COMMON PROCEDURES *********************************
// *****************************************************************************

procedure TfmMain.apAppPropException(Sender: TObject; E: Exception);
begin
  // Error handling: a more particular check is, for example
  // if E.Message = 'Invalid date format' then
  //  MessageDlg('Date not valid.', mtWarning, [mbOK], 0)
  if ShowErrorMsg = False then
    MessageDlg(msg018, mtWarning, [mbOK], 0)
  else
    MessageDlg(E.Message, mtWarning, [mbOK], 0);
end;

procedure TfmMain.CreateDataTables(DataFileName: string);
var
  mySQL: string;
begin
  // Create new archive
  // If file exists, delete it (the Open dialog doesn't work properly)
  // Close tables
  CloseDataTables;
  // Overwriting not allowed; see notes on code below
  if FileExistsUTF8(DataFileName) = True then
  begin
    MessageDlg(msg019, mtWarning, [mbOK], 0);
    Abort;
  end;
  // Create Data Table and indexes
  with sqToolsTables do
    try
      FileName := DataFileName;
      // Create Subjects Table and indexes
      mySQL := 'CREATE TABLE Subjects (';
      mySQL := MySQL + 'IDSubjects INTEGER PRIMARY KEY, ';
      mySQL := MySQL + 'SubjectsName VARCHAR(200), ';
      mySQL := MySQL + 'SubjectsComments TEXT, ';
      mySQL := MySQL + 'SubjectsBackColor VARCHAR(30), ';
      mySQL := MySQL + 'SubjectsFontColor VARCHAR(30), ';
      mySQL := MySQL + 'SubjectsSort INTEGER, ';
      mySQL := MySQL + 'SubjectsUID VARCHAR(37), ';
      mySQL := MySQL + 'SubjectsDTMod DATETIME);';
      ExecSQL(MySQL);
      mySQL := 'CREATE UNIQUE INDEX idxSubjectsIDSubjects ON Subjects (IDSubjects);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxSubjectsName ON Subjects (SubjectsName);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxSubjectsSort ON Subjects (SubjectsSort);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxSubjectsUID ON Subjects (SubjectsUID);';
      ExecSQL(MySQL);
      // Create Notes Table and indexes
      mySQL := 'CREATE TABLE Notes (';
      mySQL := MySQL + 'IDNotes INTEGER PRIMARY KEY, ';
      mySQL := MySQL + 'ID_Subjects INTEGER, ';
      mySQL := MySQL + 'NotesTitle VARCHAR(200), ';
      mySQL := MySQL + 'NotesDate DATE, ';
      mySQL := MySQL + 'NotesText TEXT, ';
      mySQL := MySQL + 'NotesTags VARCHAR(200), ';
      mySQL := MySQL + 'NotesBackColor VARCHAR(30), ';
      mySQL := MySQL + 'NotesFontColor VARCHAR(30), ';
      mySQL := MySQL + 'NotesSort INTEGER, ';
      mySQL := MySQL + 'NotesAttName TEXT, ';
      mySQL := MySQL + 'NotesActivities TEXT, ';
      mySQL := MySQL + 'NotesUID VARCHAR(37), ';
      mySQL := MySQL + 'NotesDTMod DATETIME, ';
      mySQL := MySQL + 'NotesDateFormat VARCHAR(3), ';
      mySQL := MySQL + 'NotesCheckPwd VARCHAR(200)); ';
      mySQL := MySQL + 'CONSTRAINT fkNotes_Subjects FOREIGN KEY (ID_Subjects) REFERENCES Subjects (IDSubjects));';
      ExecSQL(MySQL);
      mySQL := 'CREATE UNIQUE INDEX idxNotesIDNotes ON Notes (IDNotes);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxNotesID_Subjects ON Notes (ID_Subjects);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxNotesTitle ON Notes (NotesTitle);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxNotesSort ON Notes (NotesSort);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxNotesUID ON Notes (NotesUID);';
      ExecSQL(MySQL);
      // Create Deleted record Table and indexes
      mySQL := 'CREATE TABLE DelRec (';
      mySQL := MySQL + 'IDDelRec INTEGER PRIMARY KEY, ';
      mySQL := MySQL + 'DelRecUID VARCHAR(37), ';
      mySQL := MySQL + 'DelRecDTMod DATETIME); ';
      ExecSQL(MySQL);
      mySQL := 'CREATE UNIQUE INDEX idxDelRecIDDelrec ON Delrec (IDDelRec);';
      ExecSQL(MySQL);
      mySQL := 'CREATE INDEX idxDelRecUID ON Delrec (DelRecUID);';
      ExecSQL(MySQL);
      // Create Tools Table
      mySQL := 'CREATE TABLE Tools (';
      mySQL := MySQL + 'IDTools INTEGER PRIMARY KEY, ';
      mySQL := MySQL + 'IDLastSubject INTEGER, ';
      mySQL := MySQL + 'IDLastNote INTEGER); ';
      ExecSQL(MySQL);
    except
      MessageDlg(msg020, mtWarning, [mbOK], 0);
    end;
end;

procedure TfmMain.MakeTableUpgrade(DataFileName: string);
var
  mySQL: string;
begin
  // Update database structure
  with sqToolsTables do
    try
      // Change Notes
      FileName := DataFileName;
      TableName := 'Notes';
      SQL := 'Select * from Notes where IDNotes = -1';
      Open;
      // Upgrade to 1.0.6
      if FieldCount < 8 then
      begin
        Close;
        mySQL := 'ALTER TABLE Subjects ADD SubjectsComments TEXT';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Notes ADD NotesTags VARCHAR(200)';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Notes ADD NotesAttName TEXT';
        ExecSQL(MySQL);
        Close;
        // Error Message is in OpenTable procedure
      end
      // Upgrade to 1.1.2
      else if FieldCount < 10 then
      begin
        Close;
        mySQL := 'ALTER TABLE Notes ADD NotesCheckPwd VARCHAR(200)';
        ExecSQL(MySQL);
        mySQL := 'CREATE TABLE Tools (';
        mySQL := MySQL + 'IDTools INTEGER PRIMARY KEY, ';
        mySQL := MySQL + 'IDLastSubject INTEGER, ';
        mySQL := MySQL + 'IDLastNote INTEGER); ';
        ExecSQL(MySQL);
        Close;
        // Error Message is in OpenTable procedure
      end
      // Upgrade to 1.3.0
      else if FieldCount < 11 then
      begin
        Close;
        mySQL := 'ALTER TABLE Notes ADD NotesActivities TEXT';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Notes ADD NotesDateFormat VARCHAR(3)';
        ExecSQL(MySQL);
        Close;
        // Error Message is in OpenTable procedure
      end;
      Close;
      // Change Subjects
      FileName := DataFileName;
      TableName := 'Subjects';
      SQL := 'Select * from Subjects where IDSubjects = -1';
      Open;
      // Upgrade to 1.3.1
      if FieldCount < 6 then
      begin
        Close;
        mySQL := 'ALTER TABLE Subjects ADD SubjectsBackColor VARCHAR(30)';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Subjects ADD SubjectsFontColor VARCHAR(30)';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Subjects ADD SubjectsSort INTEGER';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxSubjectsSort ON Subjects (SubjectsSort);';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Notes ADD NotesBackColor VARCHAR(30)';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Notes ADD NotesFontColor VARCHAR(30)';
        ExecSQL(MySQL);
        mySQL := 'ALTER TABLE Notes ADD NotesSort INTEGER';
        ExecSQL(MySQL);
        mySQL := 'CREATE INDEX idxNotesSort ON Notes (NotesSort);';
        ExecSQL(MySQL);
        Close;
        SQL := 'Select * from Subjects';
        PrimaryKey := 'IDSubjects';
        Open;
        while not sqToolsTables.EOF do
        begin
          sqToolsTables.Edit;
          sqToolsTables.FieldByName('SubjectsSort').AsInteger :=
            sqToolsTables.FieldByName('IDSubjects').AsInteger;
          sqToolsTables.Post;
          sqToolsTables.Next;
        end;
        sqToolsTables.ApplyUpdates;
        Close;
        TableName := 'Notes';
        SQL := 'Select * from Notes';
        PrimaryKey := 'IDNotes';
        Open;
        while not sqToolsTables.EOF do
        begin
          sqToolsTables.Edit;
          sqToolsTables.FieldByName('NotesSort').AsInteger :=
            sqToolsTables.FieldByName('IDNotes').AsInteger;
          sqToolsTables.Post;
          sqToolsTables.Next;
        end;
        sqToolsTables.ApplyUpdates;
        // Error Message is in OpenTable procedure
      end
    finally
      PrimaryKey := '';
      Close;
    end;
end;

procedure TfmMain.OpenDataTables(DataFileName: string);
var
  sqSetId: TSqlite3Dataset;
begin
  // Open tables
  // The file could not exist (e.g. when it is loaded automatically at startup)
  if FileExistsUTF8(DataFileName) = False then
    Exit;
  // Check if the database structure must be upgraded
  try
    MakeTableUpgrade(DataFileName);
  except
    MessageDlg(msg007, mtWarning, [mbOK], 0);
    dbText.Clear;
    Abort;
  end;
  // Close tables
  CloseDataTables;
  // Following elements are made visible now because RichText must be
  // visible before opening the Notes table, otherwise it does not format the text
  pnNotes.Visible := True;
  pnGridSubjects.Visible := True;
  spSplitterSubjects.Visible := True;
  try
    with sqNotes do
    begin
      FileName := DataFileName;
      TableName := 'Notes';
      PrimaryKey := 'IDNotes';
    end;
    // sqNotes is opened in the sqSubjects.AfterScroll event
    with sqDelRec do
    begin
      // This table must be opened each time there is a delete of something
      FileName := DataFileName;
      TableName := 'DelRec';
      PrimaryKey := 'IDDelRec';
    end;
    with sqSubjects do
    begin
      FileName := DataFileName;
      TableName := 'Subjects';
      PrimaryKey := 'IDSubjects';
      if miSubjectOrderTitle.Checked = True then
        SQL := 'Select * from Subjects order by SubjectsName collate nocase, IDSubjects'
      else
        SQL := 'Select * from Subjects order by SubjectsSort';
      // To avoid that text is loaded
      IsTextToLoad := False;
      Open;
    end;
    // Get ID of last Subject and Note saved in the file in use
    try
      sqSetId := TSqlite3Dataset.Create(Self);
      sqSetId.FileName := sqNotes.FileName;
      sqSetId.TableName := 'Tools';
      sqSetId.PrimaryKey := 'IDTools';
      sqSetId.Open;
      sqSubjects.Locate('IDSubjects',
        sqSetId.FieldByName('IDLastSubject').AsInteger, []);
      // Now text can be loaded
      IsTextToLoad := True;
      if sqNotes.Active = True then
      begin
        // If there's no need to move the cursor...
        if sqSetId.FieldByName('IDLastNote').AsInteger = sqNotes.FieldByName('IDNotes').AsInteger then
          // ... load the text ...
          sqNotesAfterScroll(nil)
        // ... otherwise move the cursor and load the text...
        else
        if sqNotes.Locate('IDNotes', sqSetId.FieldByName('IDLastNote').AsInteger, []) = False then
          // ... or open it anyway
          sqNotesAfterScroll(nil);
      end;
      sqSetId.Close;
    finally
      sqSetId.Free;
    end;
    // sqNotes is opened in the sqSubjects.AfterScroll event,
    // but if there are no subjects the dataset is not opened; so...
    if sqSubjects.RecordCount = 0 then
    begin
      sqNotes.SQL := 'Select * from Notes where IDNotes = -1';
      sqNotes.Open;
      // Add a subject without asking
      ConfirmNewSubject := False;
      sqSubjects.Append;
      ConfirmNewSubject := True;
      // Everything is saved in sqSubjectsAfterInsert event
    end;
    sqFind.FileName := sqSubjects.FileName;
  except
    MessageDlg(msg007, mtWarning, [mbOK], 0);
    sqNotes.Close;
    sqSubjects.Close;
    dbText.Clear;
    Abort;
  end;
  // Create tags list
  CreateTagsList;
  // Menù items
  miFileClose.Enabled := True;
  miFileUpdate.Enabled := True;
  miFileCopyAs.Enabled := True;
  miFileImport.Enabled := True;
  miFileExport.Enabled := True;
  miFileHTML.Enabled := True;
  miFileZim.Enabled := True;
  miFileConvert.Enabled := True;
  miNotesOrderBy.Enabled := True;
  miOrderByDate.Enabled := True;
  miOrderByTitle.Enabled := True;
  miOrderCustom.Enabled := True;
  miNotesFind.Enabled := True;
  miNotesPrint.Enabled := True;
  tbFind.Enabled := True;
  miNotesMove.Enabled := True;
  miNotesInsert.Enabled := True;
  miNotesTags.Enabled := True;
  miNotesShowOnlyText.Enabled := True;
  miNotesShowActivities.Enabled := True;
  miNotesShowCal.Enabled := True;
  if ((FileExistsUTF8(SyncFolder + DirectorySeparator + ExtractFileName(DataFileName))) and (SyncFolder <> ExtractFileDir(DataFileName))) then
  begin
    miToolsSyncDo.Enabled := True;
    tbToolsSyncDo.Enabled := True;
  end;
  miToolsCompact.Enabled := True;
  miToolsLanguage.Enabled := False;
  // Update last archives menu
  if sqSubjects.FileName = LastDatabase2 then
  begin
    LastDatabase2 := LastDatabase1;
    LastDatabase1 := sqSubjects.FileName;
  end
  else if sqSubjects.FileName = LastDatabase3 then
  begin
    LastDatabase3 := LastDatabase2;
    LastDatabase2 := LastDatabase1;
    LastDatabase1 := sqSubjects.FileName;
  end
  else if sqSubjects.FileName <> LastDatabase1 then
  begin
    //also = LastDatabase4
    LastDatabase4 := LastDatabase3;
    LastDatabase3 := LastDatabase2;
    LastDatabase2 := LastDatabase1;
    LastDatabase1 := sqSubjects.FileName;
  end;
  if LastDatabase1 <> '' then
  begin
    miFileOpenLast1.Caption := ExtractFileNameOnly(LastDatabase1);
    miFileOpenLast1.Visible := True;
    miLine6.Visible := True;
  end;
  if LastDatabase2 <> '' then
  begin
    miFileOpenLast2.Caption := ExtractFileNameOnly(LastDatabase2);
    miFileOpenLast2.Visible := True;
    miLine6.Visible := True;
  end;
  if LastDatabase3 <> '' then
  begin
    miFileOpenLast3.Caption := ExtractFileNameOnly(LastDatabase3);
    miFileOpenLast3.Visible := True;
    miLine6.Visible := True;
  end;
  if LastDatabase4 <> '' then
  begin
    miFileOpenLast4.Caption := ExtractFileNameOnly(LastDatabase4);
    miFileOpenLast4.Visible := True;
    miLine6.Visible := True;
  end;
  // ExtractFileName does not show the path
  fmMain.Caption := 'MyNotex - ' + ExtractFilePath(DataFileName) + ExtractFileNameOnly(DataFileName);
  // Autosync
  if flAutosync = True then
  begin
    if miToolsSyncDo.Enabled = True then
    begin
      miToolsSyncDoClick(nil);
      flFileChanged := False;
    end;
  end;
  // Set default font
  dbText.SetDefaultFont(DefFontName, DefFontSize + ZoomFontSize);
  // Set paragraph space
  dbText.SetParagraphSpace(ParagraphSpace);
  // Set line space
  dbText.SetLineSpace(LineSpace);
  // To activate menu items and buttons
  dsNotesDataChange(nil, nil);
  // File not changed
  flFileChanged := False;
end;

procedure TfmMain.CloseDataTables;
var
  sqSetId: TSqlite3Dataset;
  i: integer;
begin
  // Set ID of Subject and Note in use
  // The first condition is useful to close the software
  // even if the file in use is deleted
  if ((FileExistsUTF8(sqSubjects.FileName)) and (sqSubjects.Active = True)) then
    try
      sqSetId := TSqlite3Dataset.Create(Self);
      sqSetId.FileName := sqNotes.FileName;
      sqSetId.TableName := 'Tools';
      sqSetId.PrimaryKey := 'IDTools';
      sqSetId.Open;
      sqSetId.Edit;
      sqSetId.FieldByName('IDLastSubject').AsInteger :=
        sqSubjects.FieldByName('IDSubjects').AsInteger;
      sqSetId.FieldByName('IDLastNote').AsInteger :=
        sqNotes.FieldByName('IDNotes').AsInteger;
      sqSetId.Post;
      sqSetId.ApplyUpdates;
      sqSetId.Close;
      // Autosync
      if flAutosync = True then
      begin
        if miToolsSyncDo.Enabled = True then
        begin
          if flFileChanged = True then
          begin
            miToolsSyncDoClick(nil);
          end;
        end;
      end;
    finally
      sqSetId.Free;
    end;
  // Close tables
  sqSubjects.Close;
  sqNotes.Close;
  sqSubjects.FileName := '';
  sqNotes.FileName := '';
  // Close search panel and functions
  if miNotesFind.Checked = True then
    miNotesFindClick(nil);
  // No show text only
  DisableShowTextOnly;
  // Close other elements
  pnNotes.Visible := False;
  pnGridSubjects.Visible := False;
  spSplitterSubjects.Visible := False;
  dbText.Clear;
  lbAttNames.Clear;
  // Menù items
  miFileClose.Enabled := False;
  miFileSave.Enabled := False;
  tbFileSave.Enabled := False;
  miFileUndo.Enabled := False;
  miFileUpdate.Enabled := False;
  miFileCopyAs.Enabled := False;
  miFileImport.Enabled := False;
  miFileExport.Enabled := False;
  miFileConvert.Enabled := False;
  miFileHTML.Enabled := False;
  miFileZim.Enabled := False;
  miSubjectNew.Enabled := False;
  tbSubjectNew.Enabled := False;
  miSubjectDelete.Enabled := False;
  miSubjectComments.Enabled := False;
  miSubjectLook.Enabled := False;
  miSubjectOrder.Enabled := False;
  miSubjectOrderTitle.Enabled := False;
  miSubjectOrderCustom.Enabled := False;
  pmSubNew.Enabled := False;
  pmSubDelete.Enabled := False;
  pmSubComments.Enabled := False;
  pmSubLook.Enabled := False;
  miNotesNew.Enabled := False;
  tbNotesNew.Enabled := False;
  miNotesDelete.Enabled := False;
  miNotesUndo.Enabled := False;
  pmNotesNew.Enabled := False;
  pmNotesDelete.Enabled := False;
  pmNotesLook.Enabled := False;
  miNotesOrderBy.Enabled := False;
  miOrderByDate.Enabled := False;
  miOrderByTitle.Enabled := False;
  miOrderCustom.Enabled := False;
  miNotesEncDecrypt.Enabled := False;
  miNotesFind.Enabled := False;
  miNotesPrint.Enabled := False;
  miNotesMove.Enabled := False;
  miNotesImages.Enabled := False;
  miNotesAttach.Enabled := False;
  miAttachNew.Enabled := False;
  miAttachOpen.Enabled := False;
  miAttachSaveAs.Enabled := False;
  miAttachDelete.Enabled := False;
  miNotesLook.Enabled := False;
  pmAttNew.Enabled := False;
  pmAttOpen.Enabled := False;
  pmAttSaveAs.Enabled := False;
  pmAttDelete.Enabled := False;
  miNotesInsert.Enabled := False;
  miNotesSendToWp.Enabled := False;
  miNotesSendToOO.Enabled := False;
  miNotesSendToLO.Enabled := False;
  miNotesSendToBrowser.Enabled := False;
  miNotesTags.Enabled := False;
  miNotesShowOnlyText.Enabled := False;
  miNotesShowActivities.Enabled := False;
  miNotesShowCal.Enabled := False;
  miToolsSyncDo.Enabled := False;
  tbToolsSyncDo.Enabled := False;
  miToolsCompact.Enabled := False;
  miToolsLanguage.Enabled := True;
  tbCut.Enabled := False;
  tbCopy.Enabled := False;
  tbCopyHtml.Enabled := False;
  tbPaste.Enabled := False;
  tbKindFontChange.Enabled := False;
  tbColorFontChange.Enabled := False;
  tbBackColorFontChange.Enabled := False;
  tbFontBold.Enabled := False;
  tbFontItalic.Enabled := False;
  tbFontUnderline.Enabled := False;
  tbFontStrike.Enabled := False;
  tbFontRestore.Enabled := False;
  tbAlignLeft.Enabled := False;
  tbAlignCenter.Enabled := False;
  tbAlignRight.Enabled := False;
  tbAlignFill.Enabled := False;
  tbAlignIndent.Enabled := False;
  tbOpenNote.Enabled := False;
  tbFind.Enabled := False;
  fmMain.Caption := 'MyNotex';
  sbStatusBar.Panels[0].Text := ' ' + sbr008;
  sbStatusBar.Panels[1].Text := '';
  sbStatusBar.Panels[2].Text := '';
  // Clear the bookmarks
  for i := 0 to 9 do
    BookmarkList[i] := '';
  // Clear undo data
  ClearUndoData;
end;

procedure TfmMain.SaveAllData;
begin
  // Save all data
  if dsSubjects.State in [dsInsert, dsEdit] then
  begin
    sqSubjects.Post;
    sqSubjects.ApplyUpdates;
  end;
  if dsNotes.State in [dsInsert, dsEdit] then
  begin
    sqNotes.Post;
    sqNotes.ApplyUpdates;
  end;
  // Restart timer if enabled
  if tmTimerSave.Enabled = True then
  begin
    tmTimerSave.Enabled := False;
    tmTimerSave.Enabled := True;
  end;
end;

procedure TfmMain.FindData(SearchWithin: boolean);
var
  IDFakeRec, TagsList, OldIDNotes, OldIDSubjects: TStringList;
  sqGetFakeID, sqGetOldID: TSqlite3Dataset;
  stOrig, stDest: WideString;
  SqlCoded: string;
  i: integer;
  IsTag: boolean;
  DateStart, DateEnd: TDate;
begin
  // Search from the bottom of the form
  // To be sure to search on last version of data
  SaveAllData;
  if edFindText.Text = '' then
  begin
    MessageDlg(msg021, mtInformation, [mbOK], 0);
    Abort;
  end;
  // Check the single date
  if cbFindKind.ItemIndex = 6 then
  begin
    try
      DateStart := StrToDate(edFindText.Text);
    except
      MessageDlg(msg059, mtError, [mbOK], 0);
      Exit;
    end;
  end;
  // Check the dates range
  if cbFindKind.ItemIndex = 7 then
  begin
    edFindText.Text := StringReplace(edFindText.Text, '  ', ' ', [rfReplaceAll]);
    edFindText.Text := StringReplace(edFindText.Text, ' ,', ',', [rfReplaceAll]);
    edFindText.Text := StringReplace(edFindText.Text, ',  ', ',', [rfReplaceAll]);
    edFindText.Text := StringReplace(edFindText.Text, ', ', ',', [rfReplaceAll]);
    try
      DateStart := StrToDate(UTF8Copy(edFindText.Text, 1, UTF8Pos(',', edFindText.Text) - 1));
    except
      MessageDlg(msg061, mtError, [mbOK], 0);
      Exit;
    end;
    try
      DateEnd := StrToDate(UTF8Copy(edFindText.Text, UTF8Pos(',', edFindText.Text) + 1, UTF8Length(edFindText.Text)));
    except
      MessageDlg(msg061, mtError, [mbOK], 0);
      Exit;
    end;
    if DateEnd < DateStart then
    begin
      MessageDlg(msg061, mtError, [mbOK], 0);
      Exit;
    end;
  end;
  if SearchWithin = False then
    meSearchCond.Clear;
  if meSearchCond.Lines.IndexOf(cbFindKind.Text + ' "' + edFindText.Text + '"') > -1 then
  begin
    Exit;
  end;
  meSearchCond.Lines.Add(cbFindKind.Text + ' "' + edFindText.Text + '"');
  StDest := '';
  // If tag is searched, clean the field
  if cbFindKind.ItemIndex = 5 then
    edFindText.Text := CleanTagsField(edFindText.Text);
  //Get the searched ID
  OldIDSubjects := TStringList.Create;
  OldIDNotes := TStringList.Create;
  if SearchWithin = True then
  begin
    sqGetOldID := TSqlite3Dataset.Create(Self);
    sqGetOldID.FileName := sqNotes.FileName;
    sqGetOldID.SQL := sqFind.SQL;
    sqGetOldID.Open;
    while not sqGetOldID.EOF do
    begin
      if cbFindKind.ItemIndex < 2 then
        OldIDSubjects.Add(sqGetOldID.FieldByName('IDSubjects').AsString)
      else
        OldIDNotes.Add(sqGetOldID.FieldByName('IDNotes').AsString);
      sqGetOldID.Next;
    end;
    sqGetOldID.Close;
    sqGetOldID.Free;
  end;
  if sqFind.Active = True then
  begin
    sqFind.Close;
  end;
  // Subject title begins with
  if cbFindKind.ItemIndex = 0 then
  begin
    sqFind.SQL := 'Select IDSubjects, SubjectsName from Subjects ' + 'where SubjectsName like "' + edFindText.Text + '%" ';
    if OldIDSubjects.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDSubjects.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Subjects.IDSubjects = ' + OldIDSubjects[i];
        if i < OldIDSubjects.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects';
  end
  // Subject title contains
  else if cbFindKind.ItemIndex = 1 then
  begin
    sqFind.SQL := 'Select IDSubjects, SubjectsName from Subjects ' + 'where SubjectsName like "%' + edFindText.Text + '%" ';
    if OldIDSubjects.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDSubjects.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Subjects.IDSubjects = ' + OldIDSubjects[i];
        if i < OldIDSubjects.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects';
  end
  // Text contains
  else if cbFindKind.ItemIndex = 2 then
  begin
    // Note title begins with
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects ' + 'and NotesTitle like "' + edFindText.Text + '%" ';
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end
  else if cbFindKind.ItemIndex = 3 then
  begin
    // Note title contains
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects ' + 'and NotesTitle like "%' + edFindText.Text + '%" ';
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end
  else if cbFindKind.ItemIndex = 4 then
  begin
    // Text contains
    SqlCoded := edFindText.Text;
    SqlCoded := StringReplace(SqlCoded, '<', #5, [rfReplaceAll]);
    SqlCoded := StringReplace(SqlCoded, '>', #6, [rfReplaceAll]);
    // Create a dataset to check if the word looked for is within html tags
    sqGetFakeID := TSqlite3Dataset.Create(Self);
    sqGetFakeID.FileName := sqNotes.FileName;
    sqGetFakeID.PrimaryKey := sqNotes.PrimaryKey;
    sqGetFakeID.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate, NotesText from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects ' + 'and Notes.NotesCheckPwd is null ' + 'and NotesText like "%' + SqlCoded + '%" ';
    sqGetFakeID.Open;
    IDFakeRec := TStringList.Create;
    while not sqGetFakeID.EOF do
    begin
      ;
      stOrig := sqGetFakeID.FieldByName('NotesText').AsWideString;
      IsTag := False;
      // Delete html tags
      for i := 1 to UTF8Length(StOrig) do
      begin
        if StOrig[i] = '<' then
          IsTag := True
        else if StOrig[i] = '>' then
          IsTag := False
        else if IsTag = False then
          StDest := StDest + StOrig[i];
      end;
      if UTF8Pos(UpperCase(SqlCoded), UpperCase(StDest)) = 0 then
        IDFakeRec.Add(sqGetFakeID.FieldByName('IDNotes').AsString);
      sqGetFakeID.Next;
    end;
    sqGetFakeID.Close;
    sqGetFakeID.Free;
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate, NotesText from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects ' + 'and Notes.NotesCheckPwd is null ' + 'and NotesText like "%' + SqlCoded + '%" ';
    if IDFakeRec.Count > 0 then
    begin
      for i := 0 to IDFakeRec.Count - 1 do
        sqFind.SQL := sqFind.SQL + ' and Notes.IDNotes <> ' + IDFakeRec[i];
    end;
    IDFakeRec.Free;
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end
  else if cbFindKind.ItemIndex = 5 then
  begin
    // Tags equal to
    TagsList := TStringList.Create;
    TagsList.Text := StringReplace(edFindText.Text, ', ', LineEnding, [rfReplaceAll]);
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects and (';
    for i := 0 to TagsList.Count - 1 do
    begin
      // Sometimes this search is case sensitive (why?), so use UpperCase
      sqFind.SQL := sqFind.SQL + ' ((UPPER(NotesTags) = "' + UpperCase(TagsList[i]) + '") ' +
        // Beginning tag
        'or (UPPER(NotesTags) like "' + UpperCase(TagsList[i]) + ', %") ' +
        // Ending tag
        'or (UPPER(NotesTags) like "%, ' + UpperCase(TagsList[i]) + '") ' +
        // Middle tag
        'or (UPPER(NotesTags) like "%, ' + UpperCase(TagsList[i]) + ', %")) ';
      if i < TagsList.Count - 1 then
        sqFind.SQL := sqFind.SQL + ' or ';
    end;
    TagsList.Free;
    sqFind.SQL := sqFind.SQL + ' )';
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end
  else if cbFindKind.ItemIndex = 6 then
  begin
    // Date equal to
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects and ' + ' Notes.NotesDate = ' + FloatToStr(DateStart);
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end
  else if cbFindKind.ItemIndex = 7 then
  begin
    // Dates between
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects ' + ' and Notes.NotesDate >= ' + FloatToStr(DateStart) + ' and Notes.NotesDate <= ' + FloatToStr(DateEnd);
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end
  else if cbFindKind.ItemIndex = 8 then
  begin
    // Attachments names contains
    sqFind.SQL := 'Select IDSubjects, SubjectsName, IDNotes, NotesTitle, ' + 'NotesDate from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects and (' + ' ((UPPER(NotesAttName) = "' + UpperCase(edFindText.Text) + '") ' +
      // Beginning name
      'or (UPPER(NotesAttName) like "' + UpperCase(edFindText.Text) + '%") ' +
      // Ending name
      'or (UPPER(NotesAttName) like "%' + UpperCase(edFindText.Text) + '") ' +
      // Middle name
      'or (UPPER(NotesAttName) like "%' + UpperCase(edFindText.Text) + '%")) ';
    sqFind.SQL := sqFind.SQL + ' )';
    if OldIDNotes.Count > 0 then
    begin
      sqFind.SQL := sqFind.SQL + ' and (';
      for i := 0 to OldIDNotes.Count - 1 do
      begin
        sqFind.SQL := sqFind.SQL + ' Notes.IDNotes = ' + OldIDNotes[i];
        if i < OldIDNotes.Count - 1 then
          sqFind.SQL := sqFind.SQL + ' or ';
      end;
      sqFind.SQL := sqFind.SQL + ') ';
    end;
    if miOrderByTitle.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesTitle collate nocase, IDNotes';
    end
    else if miOrderByDate.Checked = True then
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesDate, IDNotes';
    end
    else
    begin
      sqFind.SQL := sqFind.SQL + ' ORDER BY SubjectsName collate nocase, IDSubjects, NotesSort, IDNotes';
    end;
  end;
  OldIDSubjects.Free;
  OldIDNotes.Free;
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  grFilter.BeginUpdate;
  sqFind.Open;
  // Subjects
  if cbFindKind.ItemIndex < 2 then
  begin
    grFilter.Columns[0].Width := -1;
    grFilter.Columns[0].Visible := False;
    sqFind.FieldByName('SubjectsName').DisplayLabel := cpt012;
    grFilter.Columns[1].Width := WidthSubCol;
  end
  // Notes
  else if cbFindKind.ItemIndex < 4 then
  begin
    grFilter.Columns[0].Width := -1;
    grFilter.Columns[0].Visible := False;
    sqFind.FieldByName('SubjectsName').DisplayLabel := cpt012;
    grFilter.Columns[1].Width := WidthSubCol;
    grFilter.Columns[2].Width := -1;
    grFilter.Columns[2].Visible := False;
    sqFind.FieldByName('NotesTitle').DisplayLabel := cpt013;
    grFilter.Columns[3].Width := WidthNoteCol;
    sqFind.FieldByName('NotesDate').DisplayLabel := cpt014;
    grFilter.Columns[4].Width := WidthDateCol;
  end
  // Text contains or tag equal to or dates between or attachments names contains
  else if cbFindKind.ItemIndex > 3 then
  begin
    grFilter.Columns[0].Width := -1;
    grFilter.Columns[0].Visible := False;
    sqFind.FieldByName('SubjectsName').DisplayLabel := cpt012;
    grFilter.Columns[1].Width := WidthSubCol;
    grFilter.Columns[2].Width := -1;
    grFilter.Columns[2].Visible := False;
    sqFind.FieldByName('NotesTitle').DisplayLabel := cpt013;
    grFilter.Columns[3].Width := WidthNoteCol;
    sqFind.FieldByName('NotesDate').DisplayLabel := cpt014;
    grFilter.Columns[4].Width := WidthDateCol;
    // text contains
    if cbFindKind.ItemIndex = 4 then
    begin
      grFilter.Columns[5].Width := -1;
      grFilter.Columns[5].Visible := False;
    end;
  end;
  grFilter.EndUpdate;
  // To be ready to search in the text
  if cbFindKind.ItemIndex = 4 then
  begin
    edLocateText.Text := edFindText.Text;
  end;
  lbFound.Caption := cpt015 + ' ' + IntToStr(sqFind.RecordCount);
  // There's no way to set focus on the grfilter grid.
  // Following code does not work
  pnFind.SetFocus;
  grFilter.SetFocus;
  bnFind2.Enabled := True;
  Screen.Cursor := crDefault;
end;

procedure TfmMain.LocateFromGrid;
begin
  // Locate subjects and note from grFilter
  if sqFind.Active = True then
  begin
    sqSubjects.Locate('IDSubjects', sqFind.FieldByName('IDSubjects').AsInteger, []);
    // Search on notes
    if sqFind.FieldCount > 2 then
    begin
      sqNotes.Locate('IDNotes', sqFind.FieldByName('IDNotes').AsInteger, []);
    end;
    // Don't search automatically in the notes text here
    // because the text could be not already loaded
  end;
end;

procedure TfmMain.FindFirstNext(Sender: TObject);
var
  StartPos, ChrPos: longint;
begin
  // Find text in the current note's text
  if edLocateText.Text <> '' then
  begin
    if Sender = bnFindFirst then
    begin
      // Find first
      ChrPos := UTF8Pos(UTF8LowerCase(edLocateText.Text), UTF8LowerCase(dbText.Text));
      if ChrPos > 0 then
      begin
        ChrPos := AdjustPosForImages(ChrPos);
        dbText.SelStart := ChrPos - 1;
        dbText.SelLength := UTF8Length(edLocateText.Text);
        // To show the selection
        dbText.SetFocus;
      end
      else
      begin
        MessageDlg(msg022, mtInformation, [mbOK], 0);
      end;
    end
    else
    begin
      // Find next
      StartPos := dbText.SelStart + dbText.SelLength + 1;
      ChrPos := UTF8Pos(UTF8LowerCase(edLocateText.Text), UTF8Copy(UTF8LowerCase(dbText.Text), StartPos, Length(dbText.Text) - StartPos + 1));
      if ChrPos > 0 then
      begin
        ChrPos := StartPos + ChrPos - 1;
        ChrPos := AdjustPosForImages(ChrPos);
        dbText.SelStart := ChrPos - 1;
        dbText.SelLength := UTF8Length(edLocateText.Text);
        // To show the selection
        dbText.SetFocus;
      end
      else
      begin
        MessageDlg(msg022, mtInformation, [mbOK], 0);
      end;
    end;
  end;
end;

procedure TfmMain.LoadTitleDateGrid;
var
  i: integer;
  sqTitDates: TSqlite3Dataset;
begin
  // Load titles and date in titles grid
  sqTitDates := TSqlite3Dataset.Create(Self);
  with sqTitDates do
    try
      FileName := sqNotes.FileName;
      SQL := sqNotes.SQL;
      Open;
      // Clear and add a line
      grTitles.RowCount := 1;
      for i := 0 to RecordCount - 1 do
      begin
        grTitles.InsertColRow(False, i + 1);
        grTitles.Cells[0, i + 1] := FieldByName('IDNotes').AsString;
        grTitles.Cells[1, i + 1] := FieldByName('NotesTitle').AsString;
        grTitles.Cells[2, i + 1] := FieldByName('NotesDate').AsString;
        if FieldByName('NotesCheckPwd').AsString <> '' then
          grTitles.Cells[3, i + 1] := 'T'
        else
          grTitles.Cells[3, i + 1] := 'F';
        if FieldByName('NotesActivities').AsString <> '' then
          grTitles.Cells[4, i + 1] := 'T'
        else
          grTitles.Cells[4, i + 1] := 'F';
        grTitles.Cells[5, i + 1] := FieldByName('NotesBackColor').AsString;
        grTitles.Cells[6, i + 1] := FieldByName('NotesFontColor').AsString;
        Next;
      end;
      if grTitles.RowCount = 1 then
        grTitles.RowCount := 2;
      SyncTitleDateGrid;
    finally
      sqTitDates.Free;
    end;
end;

procedure TfmMain.SyncTitleDateGrid;
var
  i: integer;
begin
  if sqNotes.Active = True then
  begin
    for i := 1 to grTitles.RowCount - 1 do
    begin
      if grTitles.Cells[0, i] = sqNotes.FieldByName('IDNotes').AsString then
        grTitles.Row := i;
    end;
  end;
end;

procedure TfmMain.grTitlesSelectCell(Sender: TObject; aCol, aRow: integer; var CanSelect: boolean);
begin
  // Select a note from its ID
  if sqNotes.Active = True then
  begin
    sqNotes.Locate('IDNotes', grTitles.Cells[0, aRow], []);
  end;
end;

procedure TfmMain.CreateAttachment(FileNames: array of string);
var
  myZipper: TZipper;
  AttDir: string;
  i, n: integer;
begin
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  try
    AttDir := ExtractFileNameWithoutExt(sqSubjects.FileName);
    if DirectoryExistsUTF8(AttDir) = False then
      CreateDirUTF8(AttDir);
  except
    MessageDlg(msg029, mtWarning, [mbOK], 0);
    Screen.Cursor := crDefault;
    Abort;
  end;
  for i := 0 to Length(FileNames) - 1 do
  begin
    // The file is a directory
    if FileGetAttrUTF8(FileNames[i]) = 48 then
      Continue
    else
      for n := 0 to lbAttNames.Items.Count - 1 do
      begin
        if LowerCase(ExtractFileNameOnly(lbAttNames.Items[n])) = LowerCase(ExtractFileNameOnly(FileNames[i])) then
        begin
          MessageDlg(msg033 + LineEnding + ExtractFileNameOnly(FileNames[i]) + '.*.',
            mtWarning, [mbOK], 0);
          Screen.Cursor := crDefault;
          Abort;
        end;
      end;
    try
      try
        Screen.Cursor := crHourGlass;
        myZipper := TZipper.Create;
        myZipper.FileName := AttDir + DirectorySeparator + sqNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(FileNames[i]) + '.zip';
        myZipper.Entries.AddFileEntry(FileNames[i],
          ExtractFileName(FileNames[i]));
        myZipper.ZipAllFiles;
        lbAttNames.Items.Add(ExtractFileName(FileNames[i]));
        sqNotes.Edit;
        sqNotes.FieldByName('NotesAttName').AsString := lbAttNames.Items.Text;
        sqNotes.Post;
        sqNotes.ApplyUpdates;
      except
        MessageDlg(msg028, mtWarning, [mbOK], 0);
      end;
    finally
      myZipper.Free;
      Screen.Cursor := crDefault;
    end;
  end;
  lbListAttach.Caption := stAttachments + ' [' + IntToStr(lbAttNames.Items.Count) + ']';
  Screen.Cursor := crDefault;
end;

procedure TfmMain.CreateTagsList;
var
  TagsList, TagsFreq, TagsCurr: TStringList;
  sqTags: TSqlite3Dataset;
  i, idxTag: integer;
begin
  // Create tags list
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  sqTags := TSqlite3Dataset.Create(Self);
  sqTags.FileName := sqNotes.FileName;
  sqTags.PrimaryKey := sqNotes.PrimaryKey;
  sqTags.SQL := 'Select IDNotes, NotesTags from Notes where NotesTags not null';
  sqTags.Open;
  if sqTags.RecordCount > 1000 then
    try
      lbTagsNames.Items.Add(msg038);
      Exit;
    finally
      sqTags.Free;
      Screen.Cursor := crDefault;
    end;
  try
    TagsList := TStringList.Create;
    TagsFreq := TStringList.Create;
    TagsCurr := TStringList.Create;
    TagsList.Capacity := 5000;
    TagsFreq.Capacity := 5000;
    while not sqTags.EOF do
    begin
      TagsCurr.Text := StringReplace(sqTags.FieldByName('NotesTags').AsString, ', ', LineEnding, [rfReplaceAll]);
      for i := 0 to TagsCurr.Count - 1 do
      begin
        idxTag := TagsList.IndexOf(TagsCurr[i]);
        if idxTag < 0 then
        begin
          TagsList.Add(TagsCurr[i]);
          TagsFreq.Add('1');
        end
        else
          TagsFreq[idxTag] := IntToStr(StrToInt(TagsFreq[idxTag]) + 1);
      end;
      sqTags.Next;
    end;
    for i := 0 to TagsList.Count - 1 do
      TagsList[i] := TagsList[i] + ' [' + TagsFreq[i] + ']';
    TagsList.Sort;
    lbTagsNames.Clear;
    for i := 0 to TagsList.Count - 1 do
    begin
      if UTF8Pos('@', TagsList[i]) > 0 then
        lbTagsNames.Items.Add(TagsList[i]);
    end;
    for i := 0 to TagsList.Count - 1 do
    begin
      if UTF8Pos('@', TagsList[i]) = 0 then
        lbTagsNames.Items.Add(TagsList[i]);
    end;
  finally
    TagsList.Free;
    TagsFreq.Free;
    TagsCurr.Free;
    sqTags.Free;
    Screen.Cursor := crDefault;
  end;
end;

function TfmMain.CleanTagsField(myField: string): string;
var
  myText: string;
  i: integer;
  TagsList, TagsCurr: TStringList;
begin
  // Clean the tags field from wrong inputs
  if myField <> '' then
  begin
    myText := myField;
    myText := StringReplace(myText, ',', ', ', [rfReplaceAll]);
    myText := StringReplace(myText, ' ,', ',', [rfReplaceAll]);
    while Pos('  ', myText) > 0 do
      myText := StringReplace(myText, '  ', ' ', [rfReplaceAll]);
    if myText[Length(myText)] = ' ' then
      myText := Copy(myText, 0, Length(myText) - 1);
    if myText[1] = ' ' then
      myText := Copy(myText, 2, Length(myText));
    // If something wrong happens later...
    Result := myText;
    // Remove duplicates
    try
      TagsList := TStringList.Create;
      TagsList.Capacity := 500;
      TagsCurr := TStringList.Create;
      TagsCurr.Capacity := 500;
      TagsCurr.Text := StringReplace(myText, ', ', LineEnding, [rfReplaceAll]);
      for i := 0 to TagsCurr.Count - 1 do
      begin
        if TagsList.IndexOf(TagsCurr[i]) < 0 then
          TagsList.Add(TagsCurr[i]);
      end;
      myText := TagsList.Text;
    finally
      TagsList.Free;
      TagsCurr.Free;
    end;
    Result := StringReplace(myText, LineEnding, ', ', [rfReplaceAll]);
    Result := UTF8Copy(Result, 1, UTF8Length(Result) - 2);
  end
  else
    Result := '';
end;

procedure TfmMain.AddImage(FileName: string);
var
  LevResize: single;
  myImage: TPicture;
begin
  // Add image in note
  if FileExistsUTF8(FileName) = False then
    Abort;
  with fmResizeImage do
  begin
    Caption := cpt038;
    rgResizeImage.Caption := cpt038;
    bnSubCommCancel.Caption := cpt018;
    Unit6.LabelCaption := cpt039;
  end;
  try
    myImage := TPicture.Create;
    myImage.LoadFromFile(FileName);
    fmResizeImage.SetImageHeigth(myImage.Height);
    fmResizeImage.SetImageWidth(myImage.Width);
  finally
    myImage.Free;
  end;
  if fmResizeImage.ShowModal = mrOk then
  begin
    Screen.Cursor := crHourGlass;
    Application.ProcessMessages;
    case fmResizeImage.rgResizeImage.ItemIndex of
      0: LevResize := 0.05;
      1: LevResize := 0.1;
      2: LevResize := 0.3;
      3: LevResize := 0.6;
      4: LevResize := 0.8;
      5: LevResize := 1;
      6: LevResize := 1.2;
      7: LevResize := 1.4;
      8: LevResize := 1.6;
      9: LevResize := 1.8;
      10: LevResize := 2.0;
      11: LevResize := 3.0;
      else
        LevResize := 1;
    end;
    try
      dbText.LoadImageFromFile(dbText.SelStart,
        FileName, LevResize);
      EditNotesDataset;
    finally
      Screen.Cursor := crDefault;
    end;
  end;
end;

function TfmMain.AdjustPosForImages(Pos: integer): integer;
var
  slPosImg: TStringList;
  n: integer;
begin
  // Adjust cursor position for possibile images before it
  try
    Result := Pos;
    // slPosImg is created in the TRichMemo component code,
    // in the GetImagePosInText event
    slPosImg := dbText.GetImagePosInText;
    if slPosImg.Count > 0 then
    begin
      for n := 0 to slPosImg.Count - 1 do
      begin
        if StrToInt(slPosImg[n]) <= Pos then
          Inc(Pos);
      end;
    end;
  finally
    Result := Pos;
    slPosImg.Free;
  end;
end;

function TfmMain.GetBullet: integer;
var
  stNum: string;
begin
  // Get bullet of the current line
  Result := 0;
  if Copy(dbText.Lines[dbText.CaretPos.Y], 1, 4) = '•' + #9 then
    Result := 1
  else
  begin
    if Pos('.' + #9, Copy(dbText.Lines[dbText.CaretPos.Y], 1, 6)) > 0 then
    begin
      stNum := Copy(dbText.Lines[dbText.CaretPos.Y], 1, 1);
      if ((stNum >= '0') and (stNum <= '9')) then
        Result := 2
      else if ((stNum >= 'A') and (stNum <= 'Z')) then
        Result := 3
      else if ((stNum >= 'a') and (stNum <= 'z')) then
        Result := 4;
    end;
  end;
end;

procedure TfmMain.ZoomIncrease;
var
  MySelStart: integer;
begin
  // Increase font size
  SaveAllData;
  if ZoomFontSize < 36 then
  begin
    Inc(ZoomFontSize);
    dbText.SetDefaultFont(DefFontName, DefFontSize + ZoomFontSize);
    if dbText.Visible = True then
    begin
      MySelStart := dbText.SelStart;
      LoadRichMemo;
      dbText.SelStart := MySelStart;
    end;
  end;
end;

procedure TfmMain.ZoomDecrease;
var
  MySelStart: integer;
begin
  // Decrease font size
  SaveAllData;
  if ZoomFontSize > 0 then
  begin
    Dec(ZoomFontSize);
    dbText.SetDefaultFont(DefFontName, DefFontSize + ZoomFontSize);
    if dbText.Visible = True then
    begin
      MySelStart := dbText.SelStart;
      LoadRichMemo;
      dbText.SelStart := MySelStart;
    end;
  end;
end;

procedure TfmMain.DisableShowTextOnly;
begin
  // Disable Notes - Show text only
  if miNotesShowOnlyText.Checked = True then
    miNotesShowOnlyTextClick(nil);
end;

function TfmMain.CheckNumbers(Source: string): boolean;
var
  i: integer;
begin
  // Checks if a string contains only numbers
  Result := True; //there are only numbers
  if Length(Source) = 0 then
  begin
    Result := False;
  end
  else
    for i := 1 to Length(Source) do
    begin
      if ((Source[i] < '0') or (Source[i] > '9')) then
      begin
        Result := False;
        Exit;
      end;
    end;
end;

function TfmMain.AreFontParamsEqual(fp1, fp2: TFontParams): boolean;
begin
  // Check if two font params are equal
  Result := False;
  // Sometimes the font size of a text in which it is
  // not explicitly specified is a big number, so..
  if ((fp1.Size < 2) or (fp1.Size > 128)) then
    fp1.Size := DefFontSize;
  if ((fp2.Size < 2) or (fp2.Size > 128)) then
    fp2.Size := DefFontSize;
  if ((fp1.Color = fp2.Color) and (fp1.Name = fp2.Name) and (fp1.Size = fp2.Size) and (fp1.Style = fp2.Style) and (fp1.BackColor = fp2.BackColor) and (fp1.Indented = fp2.Indented) and (fp1.Alignment = fp2.Alignment)) then
    Result := True;
end;

procedure TfmMain.tmTimerSaveTimer(Sender: TObject);
begin
  // Save data on timer
  if flNoAutosave = False then
  begin
    ;
    SaveAllData;
    tmTimerSave.Enabled := False;
    tmTimerSave.Enabled := True;
  end;
end;

procedure TfmMain.tmAlarmTimer(Sender: TObject);
var
  ftTime: string = 'hh:nn am/pm';
  ftLeft: string = 'hh:nn';
begin
  // Check alarm
  if fl24Hour = True then
  begin
    ftTime := 'hh:nn';
  end;
  if fmSetAlarm.tpAlarm.Time <= Time then
  begin
    tmAlarm.Enabled := False;
    fmSetAlarm.cbAlarm.Checked := False;
    sbStatusBar.Panels[1].Text := '';
    fmShowAlarm.lbMessage.Caption :=
      msg081 + ' ' + FormatDateTime(ftTime, Time) + '!';
    fmShowAlarm.ShowModal;
  end
  else
  begin
    sbStatusBar.Panels[1].Text :=
      FormatDateTime(ftTime, fmSetAlarm.tpAlarm.Time) + ' (' + FormatDateTime(ftLeft, Time - IncMinute(fmSetAlarm.tpAlarm.Time, 1)) + ' ' + msg082 + ')';
  end;
end;

function TfmMain.IsDirectoryEmpty(const myDir: string): boolean;
var
  SearchRec: TSearchRec;
begin
  // Check if a directory is empty
  try
    // faSysFile (= normal file) to avoid that also the directory is found
    Result := (FindFirst(myDir + DirectorySeparator + '*', faSysFile, SearchRec) <> 0);
  finally
    FindClose(SearchRec);
  end;
end;

procedure TfmMain.SetCharCount(ToHide: boolean);
begin
  // Count characters
  if ToHide = False then
    sbStatusBar.Panels[2].Text :=
      msg079 + ': ' + FormatFloat('#,##0', dbText.CountChar)
  else
    sbStatusBar.Panels[2].Text := '';
end;

// *****************************************************************************
// ************************* ACTIVITIES PROCEDURES *****************************
// *****************************************************************************

// ********************** GRID ACTIVITIES MENU ITEMS ***************************

procedure TfmMain.tbActIndLeftClick(Sender: TObject);
begin
  // Indent to left
  IndentToLeft(grActGrid.Row);
  EditNotesDataset;
end;

procedure TfmMain.tbActIndRightClick(Sender: TObject);
begin
  // Indent to right
  IndentToRight(grActGrid.Row);
end;

procedure TfmMain.tbActMoveDownClick(Sender: TObject);
begin
  // Move down
  MoveDown(grActGrid.Row);
end;

procedure TfmMain.tbActMoveUpClick(Sender: TObject);
begin
  // Move up
  MoveUp(grActGrid.Row);
end;

procedure TfmMain.tbActInsertRowClick(Sender: TObject);
begin
  // Insert row
  InsertRow(grActGrid.Row);
  // Load notes
  meActNotes.Text := grActGrid.Cells[ColNotes, grActGrid.Row];
end;

procedure TfmMain.tbActDeleteRowClick(Sender: TObject);
begin
  // Delete row
  DeleteRow(grActGrid.Row);
  // Load notes
  meActNotes.Text := grActGrid.Cells[ColNotes, grActGrid.Row];
end;

procedure TfmMain.tbActShowWbsClick(Sender: TObject);
begin
  // Show Wbs
  grActGrid.Columns.Items[ColWbs - 1].Visible := tbActShowWbs.Down;
end;

procedure TfmMain.tbActMoveBackClick(Sender: TObject);
begin
  // Move activities dates backward
  ShiftBackwardActDates;
end;

procedure TfmMain.tbActMoveForClick(Sender: TObject);
begin
  // Move activities dates forward
  ShiftForwardActDates;
end;

procedure TfmMain.tbActCopyGroupClick(Sender: TObject);
begin
  // Copy a group of activities
  CopyActivityGroup;
end;

procedure TfmMain.tbActPasteGroupClick(Sender: TObject);
begin
  // Paste a group of activities
  PasteActivityGroup;
end;

procedure TfmMain.tbActCopyAllClick(Sender: TObject);
begin
  // Copy all activities in clipboard
  CopyAllRows;
end;

procedure TfmMain.tbActMoveFromTextClick(Sender: TObject);
var
  i: integer;
  stActList: TStringList;
begin
  // Move all activities from the text to the grid
  if MessageDlg(msg078, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  try
    stActList := TStringList.Create;
    i := 0;
    while i < dbText.Lines.Count do
    begin
      if UTF8Copy(dbText.Lines[i], 1, UTF8Length('')) = '' then
      begin
        stActList.Add('+' + UTF8Copy(dbText.Lines[i], 3, UTF8Length(dbText.Lines[i])));
        dbText.Lines.Delete(i);
      end
      else if UTF8Copy(dbText.Lines[i], 1, UTF8Length('')) = '' then
      begin
        stActList.Add('-' + UTF8Copy(dbText.Lines[i], 3, UTF8Length(dbText.Lines[i])));
        dbText.Lines.Delete(i);
      end
      else
      begin
        Inc(i);
      end;
    end;
    for i := 0 to stActList.Count - 1 do
    begin
      if UTF8Copy(stActList[i], 1, 1) = '-' then
        grActGrid.Cells[ColState, LastRowAct] := 'Done'
      else
        grActGrid.Cells[ColState, LastRowAct] := 'To do';
      grActGrid.Cells[ColName, LastRowAct] :=
        UTF8Copy(stActList[i], 2, UTF8Length(stActList[i]));
      grActGrid.Cells[ColCode, LastRowAct] := '*1';
      Inc(LastRowAct);
    end;
  finally
    stActList.Free;
    EditNotesDataset;
  end;
end;

procedure TfmMain.tbActClearAllClick(Sender: TObject);
begin
  // Clean activities grid
  if MessageDlg(msg073, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  if MessageDlg(msg074, mtConfirmation, [mbOK, mbCancel], 0) = mrCancel then
    Abort;
  grActGrid.Clean(1, 1, grActGrid.ColCount - 1,
    grActGrid.RowCount - 1, [gzNormal]);
  grActGrid.Row := 1;
  meActNotes.Clear;
  EditNotesDataset;
  LastRowAct := 1;
  lbActDates.Caption := cpt094 + '.';
end;

// ******************* GRID ACTIVITIES EVENTS **********************************

procedure TfmMain.grActGridDrawCell(Sender: TObject; aCol, aRow: integer; aRect: TRect; aState: TGridDrawState);
var
  Simbol: string[1];
  Level: integer;
  TitCol: string;
begin
  // Grid headings
  if aRow = 0 then
  begin
    if aCol = ColID then
      TitCol := lbID;
    if aCol = ColWbs then
      TitCol := lbWBS;
    if aCol = ColState then
      TitCol := lbState;
    if aCol = ColName then
      TitCol := lbActivity;
    if aCol = ColStartDate then
      TitCol := lbStartDate;
    if aCol = ColEndDate then
      TitCol := lbEndDate;
    if aCol = ColDuration then
      TitCol := lbDuration;
    if aCol = ColResources then
      TitCol := lbResources;
    if aCol = ColPriority then
      TitCol := lbPriority;
    if aCol = ColCompletion then
      TitCol := lbCompletion;
    if aCol = ColCost then
      TitCol := lbCost;
    if grActGrid.ColWidths[aCol] < grActGrid.Canvas.TextWidth(TitCol) + 4 then
      grActGrid.ColWidths[aCol] := grActGrid.Canvas.TextWidth(TitCol) + 4;
    grActGrid.Canvas.Rectangle(aRect.Left - 1, aRect.Top - 1,
      aRect.Right + 1, aRect.Bottom + 1);
    grActGrid.Canvas.TextOut((aRect.Right - aRect.Left - grActGrid.Canvas.TextWidth(TitCol)) div 2 + aRect.Left,
      aRect.Top + 2, TitCol);
  end
  else
    grActGrid.Canvas.FillRect(aRect);
  // Row of first level with grey background
  if ((GetCode(grActGrid.Cells[ColCode, aRow]) = 1) and (GetSign(grActGrid.Cells[ColCode, aRow]) <> '*')) then
  begin
    if aCol > ColCode then
    begin
      grActGrid.Canvas.Brush.Color := $F3F3F3;
      grActGrid.Canvas.FillRect(Classes.Rect(aRect.Left, aRect.Top, aRect.Right, aRect.Bottom));
    end;
  end;
  // First col with numbers
  if aCol = ColID then
  begin
    if aRow > 0 then
    begin
      grActGrid.Canvas.Brush.Color := clBtnFace;
      grActGrid.Canvas.TextOut(
        (aRect.Right - aRect.Left - grActGrid.Canvas.TextWidth(IntToStr(aRow))) div 2,
        aRect.Top + 4,
        grActGrid.Cells[ColID, aRow]);
      grActGrid.Canvas.Pen.Color := clSilver;
      grActGrid.Canvas.Rectangle(aRect.Left, aRect.Bottom,
        aRect.Right, aRect.Bottom + 1);
      // Note symbol
      if grActGrid.Cells[ColNotes, aRow] <> '' then
      begin
        if grActGrid.RowHeights[aRow] > -1 then
          with grActGrid.Canvas do
          begin
            Brush.Color := clDkGray;
            FillRect(Classes.Rect(grActGrid.ColWidths[0] - 8, aRect.Top + 8, grActGrid.ColWidths[0] - 2, aRect.Top + 17));
          end;
      end;
      // Row selected symbol
      if grActGrid.Row = aRow then
      begin
        if grActGrid.RowHeights[aRow] > -1 then
          with grActGrid.Canvas do
          begin
            Brush.Color := clDkGray;
            Pen.Color := clDkGray;
            Polygon([Point(aRect.Left + 2, aRect.Top + 8), Point(aRect.Left + 8, aRect.Top + 12), Point(aRect.Left + 2, aRect.Top + 15)]);
          end;
      end;
    end;
  end;
  // Start date
  if ((ACol = ColStartDate) and ((grActGrid.EditorMode = False) or (grActGrid.Row <> aRow) or (grActGrid.Col <> aCol))) then
  begin
    if ((ARow > 0) and (grActGrid.Cells[ACol, ARow] <> '')) then
      try
        grActGrid.Canvas.TextOut(
          aRect.Left + 2, aRect.Top + 3,
          FormatDateTime('ddd ' + FDate.ShortDateFormat, StrToDateTime(grActGrid.Cells[ACol, aRow], FDate)));
      except;
      end;
  end
  // End date
  else if ((ACol = ColEndDate) and ((grActGrid.EditorMode = False) or (grActGrid.Row <> aRow) or (grActGrid.Col <> aCol))) then
  begin
    if ((ARow > 0) and (grActGrid.Cells[ACol, ARow] <> '')) then
      try
        grActGrid.Canvas.TextOut(
          aRect.Left + 2, aRect.Top + 3,
          FormatDateTime('ddd ' + FDate.ShortDateFormat, StrToDateTime(grActGrid.Cells[ACol, aRow], FDate)));
      except;
      end;
  end
  // Completion
  else if ACol = ColCompletion then
  begin
    if grActGrid.Cells[ColCompletion, aRow] <> '' then
    begin
      grActGrid.Canvas.Brush.Color := clSkyBlue;
      grActGrid.Canvas.FillRect(Classes.Rect(aRect.Left + 2, aRect.Top + 2, aRect.Left + 2 + Trunc((aRect.Right - aRect.Left - 4) / 100 * StrToInt(grActGrid.Cells[ColCompletion, aRow])), aRect.Bottom - 2));
      grActGrid.Canvas.Brush.Style := bsClear;
      grActGrid.Canvas.TextOut(
        aRect.Left + 2, aRect.Top + 3,
        grActGrid.Cells[ColCompletion, aRow] + '%');
    end;
  end
  // Cost
  else if ACol = ColCost then
  begin
    if ((ARow > 0) and (grActGrid.Cells[ACol, ARow] <> '')) then
    begin
      grActGrid.Canvas.TextOut(
        aRect.Left + 2, aRect.Top + 3, grActGrid.Cells[ACol, aRow] + ' ' + CurrencyString);
    end;
  end
  // Plus/less and indentation
  else if aCol = ColName then
  begin
    if grActGrid.Cells[ColCode, ARow] <> '' then
    begin
      Simbol := GetSign(grActGrid.Cells[ColCode, aRow]);
      Level := GetCode(grActGrid.Cells[ColCode, aRow]);
      grActGrid.Canvas.TextOut(aRect.Left + IndentLev * Level,
        aRect.Top + 3, grActGrid.Cells[aCol, aRow]);
      if Simbol = '+' then
        ilImageBtActGrid.Draw(grActGrid.Canvas, aRect.Left + IndentLev * Level - 16, aRect.Top + 6, 1)
      else if Simbol = '-' then
        ilImageBtActGrid.Draw(grActGrid.Canvas, aRect.Left + IndentLev * Level - 16, aRect.Top + 6, 0);
    end;
  end
  else if ACol <> ColID then
  begin
    grActGrid.Canvas.TextOut(aRect.Left + 2,
      aRect.Top + 3, grActGrid.Cells[aCol, aRow]);
  end;
end;

procedure TfmMain.grActGridEditingDone(Sender: TObject);
var
  blMod: boolean;
  i: integer;
  myDate: TDateTime;
begin
  // Set proper fields contents in activity grid
  if grActGrid.Cells[grActGrid.Col, grActGrid.Row] = '' then
    if grActGrid.Cells[ColCode, grActGrid.Row] = '' then
      Exit;
  // Clear possible empty dates with separators or wrong dates
  if ((grActGrid.Col = ColStartDate) or (grActGrid.Col = ColEndDate)) then
  begin
    if grActGrid.Cells[grActGrid.Col, grActGrid.Row] <> '' then
    begin
      if grActGrid.Cells[grActGrid.Col, grActGrid.Row] = DateMask then
        grActGrid.Cells[grActGrid.Col, grActGrid.Row] := ''
      else
        try
          myDate := StrToDateTime(grActGrid.Cells[grActGrid.Col, grActGrid.Row], FDate);
        except
          grActGrid.Cells[grActGrid.Col, grActGrid.Row] := '';
        end;
    end;
    // Check dates
    if ((grActGrid.Cells[ColStartDate, grActGrid.Row] <> '') and (grActGrid.Cells[ColEndDate, grActGrid.Row] <> '')) then
    begin
      if StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate) > StrToDateTime(grActGrid.Cells[ColEndDate, grActGrid.Row], FDate) then
      begin
        if grActGrid.Col = ColStartDate then
        begin
          MessageDlg(msg075, mtWarning, [mbOK], 0);
          grActGrid.Cells[ColStartDate, grActGrid.Row] :=
            grActGrid.Cells[ColEndDate, grActGrid.Row];
        end
        else
        begin
          MessageDlg(msg076, mtWarning, [mbOK], 0);
          grActGrid.Cells[ColEndDate, grActGrid.Row] :=
            grActGrid.Cells[ColStartDate, grActGrid.Row];
        end;
      end
      // Update duration
      else
      begin
        grActGrid.Cells[ColDuration, grActGrid.Row] :=
          FormatFloat('#####', StrToDateTime(grActGrid.Cells[ColEndDate, grActGrid.Row], FDate) - StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate) + 1);
      end;
    end;
  end
  // Update end date from duration
  else if grActGrid.Col = ColDuration then
  begin
    if grActGrid.Cells[ColStartDate, grActGrid.Row] = '' then
      grActGrid.Cells[ColStartDate, grActGrid.Row] := DateTimeToStr(Date, FDate);
    grActGrid.Cells[ColEndDate, grActGrid.Row] :=
      DateTimeToStr(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate) + StrToInt(grActGrid.Cells[ColDuration, grActGrid.Row]) - 1, FDate);
  end
  // Set state
  else if grActGrid.Col = ColState then
  begin
    if grActGrid.Columns.Items[ColState - 1].PickList.IndexOf(grActGrid.Cells[ColState, grActGrid.Row]) < 0 then
      grActGrid.Cells[ColState, grActGrid.Row] :=
        grActGrid.Columns.Items[ColState - 1].PickList[0];
    // Set completion from state
    if grActGrid.Cells[ColState, grActGrid.Row] = grActGrid.Columns.Items[ColState - 1].PickList[2] then
      grActGrid.Cells[ColCompletion, grActGrid.Row] := '100'
    else if grActGrid.Cells[ColState, grActGrid.Row] = grActGrid.Columns.Items[ColState - 1].PickList[0] then
      grActGrid.Cells[ColCompletion, grActGrid.Row] := '0';
  end
  // Set state from completion
  else if grActGrid.Col = ColCompletion then
  begin
    if grActGrid.Cells[ColCompletion, grActGrid.Row] = '100' then
      grActGrid.Cells[ColState, grActGrid.Row] :=
        grActGrid.Columns.Items[ColState - 1].PickList[2]
    else if grActGrid.Cells[ColCompletion, grActGrid.Row] = '0' then
      grActGrid.Cells[ColState, grActGrid.Row] :=
        grActGrid.Columns.Items[ColState - 1].PickList[0]
    else
      grActGrid.Cells[ColState, grActGrid.Row] :=
        grActGrid.Columns.Items[ColState - 1].PickList[1];
  end
  else if grActGrid.Col = ColCost then
  begin
    if grActGrid.Cells[ColCost, grActGrid.Row] <> '' then
      try
        grActGrid.Cells[ColCost, grActGrid.Row] :=
          FormatCurr('#,#.##', StrToCurr(StringReplace(grActGrid.Cells[ColCost, grActGrid.Row], ThousandSeparator, '', [])));
      except
        grActGrid.Cells[ColCost, grActGrid.Row] := '';
      end;
  end;
  // Update state if null
  if grActGrid.Cells[ColState, grActGrid.Row] = '' then
    grActGrid.Cells[ColState, grActGrid.Row] :=
      grActGrid.Columns.Items[ColState - 1].PickList[0];
  // Insert code if the row is not empty
  blMod := False;
  for i := 2 to grActGrid.ColCount - 1 do
  begin
    if grActGrid.Cells[i, grActGrid.Row] <> '' then
    begin
      blMod := True;
      Break;
    end;
  end;
  if blMod = True then
  begin
    if grActGrid.Cells[ColCode, grActGrid.Row] = '' then
      grActGrid.Cells[ColCode, grActGrid.Row] := '*1';
  end
  else
  begin
    grActGrid.Cells[ColCode, grActGrid.Row] := '';
  end;
  // Set last row
  SetLastRowGridActivities;
  // Priority
  if grActGrid.Cells[ColPriority, grActGrid.Row] = '' then
    grActGrid.Cells[ColPriority, grActGrid.Row] := '0';
  // Completion
  if grActGrid.Cells[ColCompletion, grActGrid.Row] = '' then
    grActGrid.Cells[ColCompletion, grActGrid.Row] := '0'
  else if StrToInt(grActGrid.Cells[ColCompletion, grActGrid.Row]) > 100 then
    grActGrid.Cells[ColCompletion, grActGrid.Row] := '100';
  //Create WBS
  WriteWbs;
  // Update main activity
  if grActGrid.Col <> ColName then
    UpdateMainAct(grActGrid.Row);
  if IsGridEditing = True then
    EditNotesDataset;
  // Set grid editing flag
  IsGridEditing := False;
  CalcActDates;
end;

procedure TfmMain.grActGridGetEditMask(Sender: TObject; ACol, ARow: integer; var Value: string);
begin
  // Set mask edit for date fields
  if ((ACol = ColStartDate) or (ACol = ColEndDate)) then
    if DateMask = '  -  -    ' then
      Value := '!99-99-9999;1;_'
    else
      Value := '!9999-99-99;1;_';
end;

procedure TfmMain.grActGridKeyDown(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Mask edit in numeric fields
  if ((grActGrid.Col = ColCompletion) or (grActGrid.Col = ColPriority) or (grActGrid.Col = ColDuration)) then
    // system keys and numbers at the top of the keyboard and in the pad
    if not (Key in [9, 13, 37, 38, 39, 40, 46, 48..57, 96..105, 113]) then
      Key := 0
    else if (grActGrid.Col = ColCost) then
      if not (Key in [9, 13, 37, 38, 39, 40, 46, 48..57, 96..105, 113, 188, 190]) then
        Key := 0;
  if ((grActGrid.Col = ColStartDate) and (GetSign(grActGrid.Cells[ColCode, grActGrid.Row]) = '*')) then
  begin
    // Start date with space bar
    if key = 32 then
    begin
      grActGrid.Cells[ColStartDate, grActGrid.Row] :=
        DateTimeToStr(dtCalAct.Date, FDate);
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Increase start date
    else if ((Key = 39) and (Shift = [ssShift])) then
    begin
      if grActGrid.Cells[ColStartDate, grActGrid.Row] = '' then
        grActGrid.Cells[ColStartDate, grActGrid.Row] :=
          DateTimeToStr(dtCalAct.Date, FDate)
      else if grActGrid.Cells[ColDuration, grActGrid.Row] = '' then
        grActGrid.Cells[ColStartDate, grActGrid.Row] :=
          DateTimeToStr(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate) + 1, FDate)
      else
      begin
        if StrToInt(grActGrid.Cells[ColDuration, grActGrid.Row]) > 1 then
          grActGrid.Cells[ColStartDate, grActGrid.Row] :=
            DateTimeToStr(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate) + 1, FDate);
      end;
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Decrease start date
    else if ((Key = 37) and (Shift = [ssShift])) then
    begin
      if grActGrid.Cells[ColStartDate, grActGrid.Row] = '' then
        grActGrid.Cells[ColStartDate, grActGrid.Row] :=
          DateTimeToStr(dtCalAct.Date, FDate)
      else
        grActGrid.Cells[ColStartDate, grActGrid.Row] :=
          DateTimeToStr(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate) - 1, FDate);
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Delete the cell
    else if Key = 46 then
    begin
      grActGrid.Cells[ColStartDate, grActGrid.Row] := '';
      grActGrid.Cells[ColDuration, grActGrid.Row] := '';
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end;
  end
  else if ((grActGrid.Col = ColEndDate) and (GetSign(grActGrid.Cells[ColCode, grActGrid.Row]) = '*')) then
  begin
    // End date with space bar
    if key = 32 then
    begin
      if grActGrid.Cells[ColStartDate, grActGrid.Row] <> '' then
        grActGrid.Cells[ColEndDate, grActGrid.Row] :=
          DateTimeToStr(IncMonth(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate), 1), FDate)
      else
        grActGrid.Cells[ColEndDate, grActGrid.Row] :=
          DateTimeToStr(IncMonth(dtCalAct.Date, 1), FDate);
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Increase end date
    else if ((Key = 39) and (Shift = [ssShift])) then
    begin
      if grActGrid.Cells[ColEndDate, grActGrid.Row] = '' then
      begin
        if grActGrid.Cells[ColStartDate, grActGrid.Row] <> '' then
          grActGrid.Cells[ColEndDate, grActGrid.Row] :=
            DateTimeToStr(IncMonth(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate), 1), FDate)
        else
          grActGrid.Cells[ColEndDate, grActGrid.Row] :=
            DateTimeToStr(IncMonth(dtCalAct.Date, 1), FDate);
      end
      else
        grActGrid.Cells[ColEndDate, grActGrid.Row] :=
          DateTimeToStr(StrToDateTime(grActGrid.Cells[ColEndDate, grActGrid.Row], FDate) + 1, FDate);
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Decrease end date
    else if ((Key = 37) and (Shift = [ssShift])) then
    begin
      if grActGrid.Cells[ColEndDate, grActGrid.Row] = '' then
      begin
        if grActGrid.Cells[ColStartDate, grActGrid.Row] <> '' then
          grActGrid.Cells[ColEndDate, grActGrid.Row] :=
            DateTimeToStr(IncMonth(StrToDateTime(grActGrid.Cells[ColStartDate, grActGrid.Row], FDate), 1), FDate)
        else
          grActGrid.Cells[ColEndDate, grActGrid.Row] :=
            DateTimeToStr(IncMonth(dtCalAct.Date, 1), FDate);
      end
      else if grActGrid.Cells[ColDuration, grActGrid.Row] = '' then
        grActGrid.Cells[ColEndDate, grActGrid.Row] :=
          DateTimeToStr(StrToDateTime(grActGrid.Cells[ColEndDate, grActGrid.Row], FDate) - 1, FDate)
      else
      begin
        if StrToInt(grActGrid.Cells[ColDuration, grActGrid.Row]) > 1 then
          grActGrid.Cells[ColEndDate, grActGrid.Row] :=
            DateTimeToStr(StrToDateTime(grActGrid.Cells[ColEndDate, grActGrid.Row], FDate) - 1, FDate);
      end;
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Delete the cell
    else if Key = 46 then
    begin
      grActGrid.Cells[ColEndDate, grActGrid.Row] := '';
      grActGrid.Cells[ColDuration, grActGrid.Row] := '';
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end;
  end
  // Toggle status with space bar
  else if grActGrid.Col = ColState then
  begin
    if key = 32 then
    begin
      if GetSign(grActGrid.Cells[ColCode, grActGrid.Row]) = '*' then
      begin
        if grActGrid.Cells[ColState, grActGrid.Row] = grActGrid.Columns.Items[ColState - 1].PickList[0] then
          grActGrid.Cells[ColState, grActGrid.Row] :=
            grActGrid.Columns.Items[ColState - 1].PickList[1]
        else if grActGrid.Cells[ColState, grActGrid.Row] = grActGrid.Columns.Items[ColState - 1].PickList[1] then
          grActGrid.Cells[ColState, grActGrid.Row] :=
            grActGrid.Columns.Items[ColState - 1].PickList[2]
        else
          grActGrid.Cells[ColState, grActGrid.Row] :=
            grActGrid.Columns.Items[ColState - 1].PickList[0];
        key := 0;
        // To update the colours
        grActGrid.Repaint;
        key := 0;
        IsGridEditing := True;
        grActGridEditingDone(nil);
      end;
    end;
  end
  else if ((grActGrid.Col = ColDuration) and (GetSign(grActGrid.Cells[ColCode, grActGrid.Row]) = '*')) then
  begin
    if ((Key = 37) and (Shift = [ssShift])) then
    begin
      if StrToInt(grActGrid.Cells[ColDuration, grActGrid.Row]) > 1 then
        grActGrid.Cells[ColDuration, grActGrid.Row] :=
          IntToStr(StrToInt(grActGrid.Cells[ColDuration, grActGrid.Row]) - 1);
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    else if ((Key = 39) and (Shift = [ssShift])) then
    begin
      grActGrid.Cells[ColDuration, grActGrid.Row] :=
        IntToStr(StrToInt(grActGrid.Cells[ColDuration, grActGrid.Row]) + 1);
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end
    // Clear duration
    else if Key = 46 then
    begin
      grActGrid.Cells[ColDuration, grActGrid.Row] := '';
      grActGrid.Cells[ColStartDate, grActGrid.Row] := '';
      grActGrid.Cells[ColEndDate, grActGrid.Row] := '';
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end;
  end
  // Clear cost
  else if ((grActGrid.Col = ColCost) and (GetSign(grActGrid.Cells[ColCode, grActGrid.Row]) = '*')) then
  begin
    // Delete the cell
    if Key = 46 then
    begin
      grActGrid.Cells[ColCost, grActGrid.Row] := '';
      // To update the colours
      grActGrid.Repaint;
      key := 0;
      IsGridEditing := True;
      grActGridEditingDone(nil);
    end;
  end;
  // Indent left
  if ((Key = 37) and (Shift = [ssCtrl])) then
  begin
    IndentToLeft(grActGrid.Row);
    Key := 0;
  end
  // Indent right
  else if ((Key = 39) and (Shift = [ssCtrl])) then
  begin
    IndentToRight(grActGrid.Row);
    Key := 0;
  end
  // Move up
  else if ((Key = 38) and (Shift = [ssCtrl])) then
  begin
    MoveUp(grActGrid.Row);
    Key := 0;
  end
  // Move down
  else if ((Key = 40) and (Shift = [ssCtrl])) then
  begin
    MoveDown(grActGrid.Row);
    Key := 0;
  end
  // Insert row
  else if ((Key = Ord('N')) and (Shift = [ssCtrl, ssShift])) then
  begin
    InsertRow(grActGrid.Row);
    Key := 0;
  end
  // Delete row
  else if ((Key = Ord('D')) and (Shift = [ssCtrl, ssShift])) then
  begin
    DeleteRow(grActGrid.Row);
    Key := 0;
  end
  // Copy groups rows
  else if ((Key = Ord('C')) and (Shift = [ssCtrl, ssShift])) then
  begin
    CopyActivityGroup;
    Key := 0;
  end
  // Paste groups rows
  else if ((Key = Ord('V')) and (Shift = [ssCtrl, ssShift])) then
  begin
    PasteActivityGroup;
    Key := 0;
  end
  // Copy all rows
  else if ((Key = Ord('L')) and (Shift = [ssCtrl, ssShift])) then
  begin
    CopyAllRows;
    Key := 0;
  end;
end;

procedure TfmMain.grActGridKeyUp(Sender: TObject; var Key: word; Shift: TShiftState);
begin
  // Activate chenges on Ctrl + V
  if ((Key = Ord('V')) and (Shift = [ssCtrl])) then
  begin
    IsGridEditing := True;
    grActGridEditingDone(nil);
  end;
end;

procedure TfmMain.grActGridPrepareCanvas(Sender: TObject; aCol, aRow: integer; aState: TGridDrawState);
var
  myDate: TDate;
begin
  // Possible bold
  if grActGrid.Cells[ColCode, aRow] <> '' then
  begin
    if Copy(grActGrid.Cells[ColCode, aRow], 1, 1) <> '*' then
    begin
      if ((aRow > 0) and (aCol > ColWbs)) then
      begin
        grActGrid.Canvas.Font.Style := [fsBold];
      end;
    end;
  end;
  // Possible red or green
  if aCol > 0 then
  begin
    if grActGrid.Cells[ColState, aRow] = grActGrid.Columns.Items[ColState - 1].PickList[2] then
    begin
      grActGrid.Canvas.Font.Color := clGreen;
    end
    else
    begin
      if grActGrid.Cells[ColState, aRow] <> grActGrid.Columns.Items[ColState - 1].PickList[2] then
        try
          if grActGrid.Cells[ColEndDate, aRow] <> '' then
          begin
            if TryStrToDate(grActGrid.Cells[ColEndDate, aRow], myDate, FDate) = True then
            begin
              if myDate < Date then
              begin
                grActGrid.Canvas.Font.Color := clRed;
              end;
            end;
          end;
        except;
        end;
    end;
  end;
end;

procedure TfmMain.grActGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: integer);
var
  Simbol: string[1];
  Level, Dist, i: integer;
begin
  // Expand or compact lines on mouse click
  if grActGrid.MouseCoord(X, Y).X = ColName then
  begin
    Simbol := Copy(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y], 1, 1);
    if ((Simbol = '+') or (Simbol = '-')) then
    begin
      if grActGrid.LeftCol > ColName then
        Abort;
      Dist := 0;
      for i := grActGrid.LeftCol to ColName - 1 do
        Dist := Dist + grActGrid.ColWidths[i];
      Dist := Dist + grActGrid.ColWidths[0];
      Level := StrToInt(Copy(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y], 2, Length(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y])));
      if ((X > Dist + IndentLev * Level - 16) and (X < Dist + IndentLev * Level)) then
      begin
        if GetSign(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y]) = '-' then
          CompactLines(grActGrid.MouseCoord(X, Y).Y,
            GetCode(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y]))
        else if GetSign(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y]) = '+' then
          ExpandLines(grActGrid.MouseCoord(X, Y).Y,
            GetCode(grActGrid.Cells[ColCode, grActGrid.MouseCoord(X, Y).Y]));
        // To avoid editor activation
        Abort;
      end;
    end;
  end;
end;

procedure TfmMain.grActGridSelectCell(Sender: TObject; aCol, aRow: integer; var CanSelect: boolean);
begin
  // Heading rows as read only
  // It must fire only after grid has been completely constructed, so has 13 columns
  if grActGrid.ColCount = 13 then
  begin
    grActGrid.Options := grActGrid.Options + [goEditing];
    if ((GetSign(grActGrid.Cells[ColCode, aRow]) = '+') or (GetSign(grActGrid.Cells[ColCode, aRow]) = '-')) then
    begin
      if ((aCol <> ColName) and (aCol <> ColNotes) and (aCol <> ColResources)) then
        grActGrid.Options := grActGrid.Options - [goEditing];
    end;
  end;
end;

procedure TfmMain.grActGridSelection(Sender: TObject; aCol, aRow: integer);
begin
  // Load notes
  meActNotes.Text := grActGrid.Cells[ColNotes, grActGrid.Row];
  // To show the indicator in the grid
  grActGrid.Update;
end;

procedure TfmMain.grActGridSetEditText(Sender: TObject; ACol, ARow: integer; const Value: string);
begin
  // Set grid editing flag
  IsGridEditing := True;
  EditNotesDataset;
end;

// ******************* GRID ACTIVITIES MEMO EVENTS *****************************

procedure TfmMain.meActNotesChange(Sender: TObject);
begin
  // Update grid notes field and put the dataset in edit
  if grActGrid.Cells[ColNotes, grActGrid.Row] <> StringReplace(meActNotes.Text, #9, '', [rfReplaceAll]) then
  begin
    grActGrid.Cells[ColNotes, grActGrid.Row] :=
      StringReplace(meActNotes.Text, #9, '', [rfReplaceAll]);
    if grActGrid.Cells[ColCode, grActGrid.Row] = '' then
      grActGrid.Cells[ColCode, grActGrid.Row] := '*1';
    EditNotesDataset;
  end;
end;

// ******************* GRID ACTIVITIES Load and Save ***************************

procedure TfmMain.SaveActivitiesData;
var
  x, y: integer;
  myData: string;
begin
  // Save activities
  if sqNotes.State in [dsEdit, dsInsert] then
  begin
    // Empty grid
    if ((LastRowAct = 1) and (grActGrid.Cells[1, 1] = '')) then
      sqNotes.FieldByName('NotesActivities').AsString := ''
    else
    begin
      for y := 1 to LastRowAct do
      begin
        for x := 1 to grActGrid.ColCount - 1 do
        begin
          grActGrid.Cells[x, y] :=
            StringReplace(grActGrid.Cells[x, y], #9, ' ', [rfReplaceAll]);
          myData := myData + grActGrid.Cells[x, y] + #9;
        end;
        myData := myData + LineEnding;
      end;
      if sqNotes.FieldByName('NotesActivities').AsString <> myData then
      begin
        sqNotes.FieldByName('NotesActivities').AsString := myData;
        // To show the grey box
        grActGrid.Refresh;
      end;
    end;
  end;
end;

procedure TfmMain.LoadActivitiesData;
var
  x: integer;
  myData: string;
  stDtFormat: string;
begin
  // Load activities
  stDtFormat := sqNotes.FieldByName('NotesDateFormat').AsString;
  grActGrid.Clean(1, 1, grActGrid.ColCount - 1,
    grActGrid.RowCount - 1, [gzNormal]);
  grActGrid.Row := 1;
  meActNotes.Clear;
  LastRowAct := 1;
  if sqNotes.FieldByName('NotesActivities').AsString <> '' then
  begin
    x := 1;
    myData := sqNotes.FieldByName('NotesActivities').AsString;
    while UTF8Length(myData) > 0 do
    begin
      if ((DateOrder <> stDtFormat) and ((x = ColStartDate) or (x = ColEndDate))) then
      begin
        grActGrid.Cells[x, LastRowAct] :=
          ConvertDateFormat(UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1), stDtFormat, DateOrder);
      end
      else
        grActGrid.Cells[x, LastRowAct] :=
          UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1);
      myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
      if UTF8Copy(myData, 1, 1) = LineEnding then
      begin
        myData := UTF8Copy(myData, 2, UTF8Length(myData));
        x := 1;
        Inc(LastRowAct);
      end
      else
        Inc(x);
    end;
    Dec(LastRowAct);
  end;
  meActNotes.Text := grActGrid.Cells[ColNotes, grActGrid.Row];
  CalcActDates;
end;

// ******************* GRID ACTIVITIES MAIN PROCEDURES *************************

procedure TfmMain.SetLastRowGridActivities;
var
  i: integer;
begin
  //Set last row of the grid
  for i := grActGrid.RowCount - 1 downto 1 do
  begin
    if grActGrid.Cells[ColCode, i] <> '' then
    begin
      LastRowAct := i;
      Break;
    end;
  end;
end;

procedure TfmMain.CompactLines(NumRow: integer; Cod: integer);
var
  i: integer;
begin
  // Compact lines
  grActGrid.Cells[ColCode, NumRow] :=
    '+' + IntToStr(GetCode(grActGrid.Cells[ColCode, NumRow]));
  i := NumRow + 1;
  while GetCode(grActGrid.Cells[ColCode, i]) > Cod do
  begin
    grActGrid.RowHeights[i] := 0;
    if i = LastRowAct then
      Break
    else
      i := i + 1;
  end;
end;

procedure TfmMain.ExpandLines(NumRow: integer; Cod: integer);
var
  i: integer;
begin
  // Expand lines
  grActGrid.Cells[ColCode, NumRow] :=
    '-' + IntToStr(GetCode(grActGrid.Cells[ColCode, NumRow]));
  i := NumRow + 1;
  while GetCode(grActGrid.Cells[ColCode, i]) > Cod do
  begin
    grActGrid.RowHeights[i] := grActGrid.DefaultRowHeight;
    if GetSign(grActGrid.Cells[ColCode, i]) = '+' then
      grActGrid.Cells[ColCode, i] :=
        '-' + IntToStr(GetCode(grActGrid.Cells[ColCode, i]));
    if i = LastRowAct then
      Break
    else
      i := i + 1;
  end;
end;

procedure TfmMain.IndentToRight(ARow: integer);
var
  CodIndent, i: integer;
begin
  // Indent lines to right
  grActGrid.EditorMode := False;
  if ((aRow > 0) and (grActGrid.Cells[ColCode, aRow] <> '') and (grActGrid.Cells[ColCode, aRow - 1] <> '')) then
  begin
    if GetCode(grActGrid.Cells[ColCode, aRow - 1]) + 1 > GetCode(grActGrid.Cells[ColCode, aRow]) then
    begin
      CodIndent := GetCode(grActGrid.Cells[ColCode, aRow]);
      CodIndent := CodIndent + 1;
      grActGrid.Cells[ColCode, aRow] :=
        Copy(grActGrid.Cells[ColCode, aRow], 1, 1) + IntToStr(CodIndent);
      if GetCode(grActGrid.Cells[ColCode, aRow - 1]) < GetCode(grActGrid.Cells[ColCode, aRow]) then
        grActGrid.Cells[ColCode, aRow - 1] :=
          '-' + IntToStr(GetCode(grActGrid.Cells[ColCode, aRow - 1]));
      if aRow < grActGrid.RowCount - 1 then
      begin
        i := aRow + 1;
        while GetCode(grActGrid.Cells[ColCode, i]) >= CodIndent do
        begin
          grActGrid.Cells[ColCode, i] :=
            GetSign(grActGrid.Cells[ColCode, i]) + IntToStr(GetCode(grActGrid.Cells[ColCode, i]) + 1);
          if ((i = LastRowAct) or (i = grActGrid.RowCount - 1)) then
            Break
          else
            i := i + 1;
        end;
      end;
      grActGrid.Repaint;
    end;
  end;
  WriteWbs;
  UpdateMainAct(aRow);
  EditNotesDataset;
end;

procedure TfmMain.IndentToLeft(aRow: integer);
var
  CodIndent, i: integer;
begin
  // Indent lines to left
  grActGrid.EditorMode := False;
  if ((aRow > 0) and (GetCode(grActGrid.Cells[ColCode, aRow]) > 1)) then
  begin
    CodIndent := GetCode(grActGrid.Cells[ColCode, aRow]);
    CodIndent := CodIndent - 1;
    grActGrid.Cells[ColCode, aRow] :=
      Copy(grActGrid.Cells[ColCode, aRow], 1, 1) + IntToStr(CodIndent);
    if aRow < grActGrid.RowCount - 1 then
    begin
      if GetCode(grActGrid.Cells[ColCode, aRow + 1]) > GetCode(grActGrid.Cells[ColCode, aRow]) then
        grActGrid.Cells[ColCode, aRow] :=
          '-' + IntToStr(GetCode(grActGrid.Cells[ColCode, aRow]));
    end;
    if aRow > 1 then
    begin
      if GetCode(grActGrid.Cells[ColCode, aRow - 1]) >= GetCode(grActGrid.Cells[ColCode, aRow]) then
        grActGrid.Cells[ColCode, aRow - 1] :=
          '*' + IntToStr(GetCode(grActGrid.Cells[ColCode, aRow - 1]));
    end;
    if aRow < grActGrid.RowCount - 1 then
    begin
      i := aRow + 1;
      while GetCode(grActGrid.Cells[ColCode, i]) > CodIndent + 1 do
      begin
        grActGrid.Cells[ColCode, i] :=
          GetSign(grActGrid.Cells[ColCode, i]) + IntToStr(GetCode(grActGrid.Cells[ColCode, i]) - 1);
        if ((i = LastRowAct) or (i = grActGrid.RowCount - 1)) then
          Break
        else
          i := i + 1;
      end;
    end;
    grActGrid.Repaint;
  end;
  WriteWbs;
  UpdateMainAct(aRow);
  // Indent to left possible subactivities recursively
  if aRow < grActGrid.RowCount - 1 then
  begin
    if GetCode(grActGrid.Cells[ColCode, aRow + 1]) > 0 then
    begin
      if GetCode(grActGrid.Cells[ColCode, aRow]) = GetCode(grActGrid.Cells[ColCode, aRow + 1]) + 2 then
        IndentToLeft(aRow + 1);
    end;
  end;
  // Update possible previous activities
  if aRow > 2 then
    if GetCode(grActGrid.Cells[ColCode, aRow - 1]) > 1 then
      UpdateMainAct(aRow - 1);
  EditNotesDataset;
end;

procedure TfmMain.MoveUp(aRow: integer);
var
  i, IDDest, CodeCurrAct: integer;
  slSubAct: TStringList;
begin
  // Move up activities
  if ((grActGrid.Cells[ColCode, aRow] = '') or (aRow = 1)) then
    Exit;
  // Check a possible destination ID
  IDDest := 0;
  CodeCurrAct := GetCode(grActGrid.Cells[ColCode, aRow]);
  for i := aRow - 1 downto 1 do
  begin
    if CodeCurrAct > GetCode(grActGrid.Cells[ColCode, i]) then
      Break
    else if CodeCurrAct = GetCode(grActGrid.Cells[ColCode, i]) then
    begin
      IDDest := i;
      Break;
    end;
  end;
  if IDDest = 0 then
    Exit;
  // Collect rows to be moved and move them
  slSubAct := TStringList.Create;
  try
    slSubAct.Add(grActGrid.Rows[aRow].CommaText);
    grActGrid.DeleteRow(aRow);
    while ((aRow <= grActGrid.RowCount - 1) and (CodeCurrAct < GetCode(grActGrid.Cells[ColCode, aRow]))) do
    begin
      if grActGrid.Cells[ColCode, aRow] = '' then
        Break;
      slSubAct.Add(grActGrid.Rows[aRow].CommaText);
      grActGrid.DeleteRow(aRow);
    end;
    if IDDest > grActGrid.RowCount then
      IDDest := grActGrid.RowCount;
    for i := slSubAct.Count - 1 downto 0 do
    begin
      grActGrid.InsertRowWithValues(IDDest, []);
      grActGrid.Rows[IDDest].CommaText := slSubAct[i];
    end;
    grActGrid.Row := IDDest;
    // Reset ID rows
    for i := 1 to LastRowAct do
      grActGrid.Cells[ColID, i] := IntToStr(i);
    WriteWbs;
    EditNotesDataset;
  finally
    slSubAct.Free;
  end;
end;

procedure TfmMain.MoveDown(aRow: integer);
var
  i, IDDest, CodeCurrAct: integer;
  slSubAct: TStringList;
begin
  // Move down activities
  if ((grActGrid.Cells[ColCode, aRow] = '') or (aRow = grActGrid.RowCount - 1)) then
    Exit;
  // Check a possible destination ID
  IDDest := 0;
  CodeCurrAct := GetCode(grActGrid.Cells[ColCode, aRow]);
  for i := aRow + 1 to grActGrid.RowCount - 1 do
  begin
    if CodeCurrAct > GetCode(grActGrid.Cells[ColCode, i]) then
      Break
    else if CodeCurrAct = GetCode(grActGrid.Cells[ColCode, i]) then
    begin
      IDDest := i;
      Break;
    end;
  end;
  if IDDest = 0 then
    Exit;
  // Collect rows to be moved and move them
  slSubAct := TStringList.Create;
  try
    slSubAct.Add(grActGrid.Rows[aRow].CommaText);
    grActGrid.DeleteRow(aRow);
    while ((aRow <= grActGrid.RowCount - 1) and (CodeCurrAct < GetCode(grActGrid.Cells[ColCode, aRow]))) do
    begin
      if grActGrid.Cells[ColCode, aRow] = '' then
        Break;
      slSubAct.Add(grActGrid.Rows[aRow].CommaText);
      grActGrid.DeleteRow(aRow);
    end;
    //Adjust IDDest according to the length of the next block
    if aRow < grActGrid.RowCount - 1 then
    begin
      for i := aRow + 1 to grActGrid.RowCount - 1 do
      begin
        if ((CodeCurrAct >= GetCode(grActGrid.Cells[ColCode, i])) or (grActGrid.Cells[ColCode, i] = '')) then
        begin
          IDDest := i;
          Break;
        end;
      end;
    end;
    if IDDest > grActGrid.RowCount then
      IDDest := grActGrid.RowCount;
    for i := slSubAct.Count - 1 downto 0 do
    begin
      grActGrid.InsertRowWithValues(IDDest, []);
      grActGrid.Rows[IDDest].CommaText := slSubAct[i];
    end;
    grActGrid.Row := IDDest;
    // Reset ID rows
    for i := 1 to LastRowAct do
      grActGrid.Cells[ColID, i] := IntToStr(i);
    WriteWbs;
    EditNotesDataset;
  finally
    slSubAct.Free;
  end;
end;

procedure TfmMain.InsertRow(aRow: integer);
var
  i: integer;
begin
  // Insert an empty row and delete the last one if empty
  if ((aRow < grActGrid.RowCount - 1) and (grActGrid.Cells[ColCode, aRow] <> '')) then
  begin
    grActGrid.InsertColRow(False, aRow);
    grActGrid.Cells[ColCode, aRow] :=
      '*' + IntToStr(GetCode(grActGrid.Cells[ColCode, aRow + 1]));
    grActGrid.Row := aRow;
    if grActGrid.Cells[ColCode, grActGrid.RowCount - 1] = '' then
      grActGrid.DeleteRow(grActGrid.RowCount - 1);
    for i := 1 to grActGrid.RowCount - 1 do
      grActGrid.Cells[ColID, i] := IntToStr(i);
    WriteWbs;
    SetLastRowGridActivities;
    EditNotesDataset;
  end;
end;

procedure TfmMain.DeleteRow(aRow: integer);
var
  i, Code: integer;
begin
  // Delete row
  if grActGrid.Cells[ColCode, aRow] = '' then
    Exit;
  if MessageDlg(msg077, mtConfirmation, [mbOK, mbCancel], 0) = mrOk then
  begin
    Code := GetCode(grActGrid.Cells[ColCode, aRow]);
    grActGrid.DeleteRow(aRow);
    grActGrid.RowCount := NumRow;
    // Delete subactivities
    while Code < GetCode(grActGrid.Cells[ColCode, aRow]) do
    begin
      grActGrid.DeleteRow(aRow);
      grActGrid.RowCount := NumRow;
    end;
    for i := 1 to grActGrid.RowCount - 1 do
      grActGrid.Cells[ColID, i] := IntToStr(i);
    WriteWbs;
    SetLastRowGridActivities;
    EditNotesDataset;
  end;
end;

procedure TfmMain.CopyAllRows;
var
  x, y: integer;
  myData: string;
begin
  // Copy all activities in clipboard
  myData := lbWBS + #9 + lbState + #9 + lbActivity + #9 + lbStartDate + #9 + lbEndDate + #9 + lbDuration + #9 + lbResources + #9 + lbPriority + #9 + lbCompletion + #9 + lbCost + #9 + lbNotes + LineEnding;
  for y := 1 to LastRowAct do
  begin
    for x := 2 to grActGrid.ColCount - 1 do
    begin
      myData := myData + StringReplace(grActGrid.Cells[x, y], LineEnding, ' ', [rfReplaceAll]) + #9;
    end;
    myData := myData + LineEnding;
  end;
  Clipboard.AsText := myData;
end;

procedure TfmMain.CopyActivityGroup;
var
  y, IDStart, IDEnd: integer;
begin
  // Copy activity group in ActivityGroup variable;
  for IDStart := grActGrid.Row downto 1 do
  begin
    if grActGrid.Cells[ColCode, IDStart] = '' then
    begin
      Break;
    end
    else if GetCode(grActGrid.Cells[ColCode, IDStart]) = 1 then
      Break;
  end;
  if grActGrid.Cells[ColCode, IDStart] = '' then
    Inc(IDStart);
  for IDEnd := grActGrid.Row to LastRowAct do
  begin
    if grActGrid.Cells[ColCode, IDEnd] = '' then
    begin
      Break;
    end
    else if ((GetCode(grActGrid.Cells[ColCode, IDEnd]) = 1) and (IDEnd > grActGrid.Row)) then
    begin
      Break;
    end;
  end;
  if IDEnd < LastRowAct then
    Dec(IDEnd);
  ActivityGroup.Clear;
  for y := IDStart to IDEnd do
  begin
    grActGrid.Rows[y].Delimiter := #9;
    ActivityGroup.Add(grActGrid.Rows[y].CommaText);
  end;
  tbActPasteGroup.Enabled := True;
end;

procedure TfmMain.PasteActivityGroup;
var
  x, IDStart: integer;
begin
  // Paste activity group from ActivityGroup variable;
  for IDStart := grActGrid.Row downto 1 do
  begin
    if grActGrid.Cells[ColCode, IDStart] = '' then
    begin
      Break;
    end
    else if GetCode(grActGrid.Cells[ColCode, IDStart]) = 1 then
      Break;
  end;
  for x := ActivityGroup.Count - 1 downto 0 do
  begin
    grActGrid.InsertColRow(False, IDStart);
    grActGrid.Rows[IDStart].Delimiter := #9;
    grActGrid.Rows[IDStart].CommaText := ActivityGroup[x];
  end;
  WriteWbs;
  SetLastRowGridActivities;
  EditNotesDataset;
end;

procedure TfmMain.WriteWbs;
var
  WbsArray: array[0..9] of integer;
  i, n, x: integer;
  StCode: string;
begin
  // Write work breakdown structure
  for i := 0 to 9 do
    WbsArray[i] := 0;
  for i := 0 to LastRowAct do
  begin
    if grActGrid.Cells[ColCode, i] <> '' then
    begin
      if GetCode(grActGrid.Cells[ColCode, i]) < 10 then
      begin
        Inc(WbsArray[GetCode(grActGrid.Cells[ColCode, i]) - 1]);
        for x := GetCode(grActGrid.Cells[ColCode, i]) to 9 do
          WbsArray[x] := 0;
        StCode := '';
        for n := 0 to GetCode(grActGrid.Cells[ColCode, i]) - 1 do
        begin
          if StCode <> '' then
            StCode := StCode + '.';
          StCode := StCode + IntToStr(WbsArray[n]);
        end;
        grActGrid.Cells[ColWbs, i] := StCode;
      end;
    end;
  end;
end;

procedure TfmMain.UpdateMainAct(IDRow: integer);
var
  i, n, ActIndent, IDTop, CompletionTotal, Priority, IncPrior: integer;
  StateTodo, StateStarted, StateDone, CostValue: boolean;
  DateIni, DateEnd: TDateTime;
  Cost: currency;
  stCost: string;
  MyIDList: TStringList;
begin
  // Update main activity of the current one
  if ((IDRow < 2) or (grActGrid.Cells[ColCode, grActGrid.Row] = '')) then
    Exit;
  ActIndent := GetCode(grActGrid.Cells[ColCode, IDRow]);
  if ActIndent = 1 then
    Exit;
  MyIDList := TStringList.Create;
  try
    //  Activities toward top, and their ID in list
    MyIDList.Add(IntToStr(IDRow));
    for i := IDRow - 1 downto 1 do
    begin
      if ActIndent = GetCode(grActGrid.Cells[ColCode, i]) then
        MyIDList.Add(IntToStr(i));
      if ((grActGrid.Cells[ColCode, i] = '') or (ActIndent > GetCode(grActGrid.Cells[ColCode, i]))) then
      begin
        IDTop := i;
        Break;
      end;
    end;
    //  Activities toward bottom, and their ID in list
    for i := IDRow + 1 to LastRowAct do
    begin
      if ActIndent = GetCode(grActGrid.Cells[ColCode, i]) then
        MyIDList.Add(IntToStr(i));
      if ((grActGrid.Cells[ColCode, i] = '') or (ActIndent > GetCode(grActGrid.Cells[ColCode, i]))) then
        Break;
    end;
    // Update main activity
    StateTodo := False;
    StateStarted := False;
    StateDone := False;
    DateIni := 0;
    DateEnd := 0;
    CompletionTotal := 0;
    Priority := 0;
    IncPrior := 0;
    Cost := 0;
    stCost := '';
    CostValue := False;
    for n := 0 to MyIDList.Count - 1 do
    begin
      i := StrToInt(MyIDList[n]);
      if grActGrid.Cells[ColState, i] <> '' then
      begin
        if StateTodo = False then
          if grActGrid.Columns.Items[ColState - 1].PickList.IndexOf(grActGrid.Cells[ColState, i]) = 0 then
            StateTodo := True;
        if StateStarted = False then
          if grActGrid.Columns.Items[ColState - 1].PickList.IndexOf(grActGrid.Cells[ColState, i]) = 1 then
            StateStarted := True;
        if StateDone = False then
          if grActGrid.Columns.Items[ColState - 1].PickList.IndexOf(grActGrid.Cells[ColState, i]) = 2 then
            StateDone := True;
      end;
      if grActGrid.Cells[ColStartDate, i] <> '' then
      begin
        if ((DateIni = 0) or (DateIni > StrToDateTime(grActGrid.Cells[ColStartDate, i], FDate))) then
          DateIni := StrToDateTime(grActGrid.Cells[ColStartDate, i], FDate);
      end;
      if grActGrid.Cells[ColEndDate, i] <> '' then
      begin
        if ((DateEnd = 0) or (DateEnd < StrToDateTime(grActGrid.Cells[ColEndDate, i], FDate))) then
          DateEnd := StrToDateTime(grActGrid.Cells[ColEndDate, i], FDate);
      end;
      if grActGrid.Cells[ColPriority, i] <> '' then
        if StrToInt(grActGrid.Cells[ColPriority, i]) > Priority then
          Priority := StrToInt(grActGrid.Cells[ColPriority, i]);
      if grActGrid.Cells[ColCompletion, i] <> '' then
      begin
        CompletionTotal := CompletionTotal + StrToInt(grActGrid.Cells[ColCompletion, i]);
        Inc(IncPrior);
      end;
      if grActGrid.Cells[ColCost, i] <> '' then
      begin
        CostValue := True;
        stCost := grActGrid.Cells[ColCost, i];
        stCost := StringReplace(stCost, ThousandSeparator, '', []);
        Cost := Cost + StrToCurr(stCost);
      end;
    end;
  finally
    MyIDList.Free;
  end;
  if ((StateTodo = False) and (StateStarted = False)) then
    grActGrid.Cells[ColState, IDTop] :=
      grActGrid.Columns.Items[ColState - 1].PickList[2]
  else if ((StateTodo = True) and (StateStarted = False) and (StateDone = False)) then
    grActGrid.Cells[ColState, IDTop] :=
      grActGrid.Columns.Items[ColState - 1].PickList[0]
  else
    grActGrid.Cells[ColState, IDTop] :=
      grActGrid.Columns.Items[ColState - 1].PickList[1];
  if DateIni > 1 then
    grActGrid.Cells[ColStartDate, IDTop] := DateToStr(DateIni, FDate)
  else
    grActGrid.Cells[ColStartDate, IDTop] := '';
  if DateEnd > 1 then
    grActGrid.Cells[ColEndDate, IDTop] := DateToStr(DateEnd, FDate)
  else
    grActGrid.Cells[ColEndDate, IDTop] := '';
  if ((DateIni > 1) and (DateEnd > 1)) then
    grActGrid.Cells[ColDuration, IDTop] :=
      FormatFloat('#####', DateEnd - DateIni + 1)
  else
    grActGrid.Cells[ColDuration, IDTop] := '';
  grActGrid.Cells[ColPriority, IDTop] := IntToStr(Priority);
  if CompletionTotal > 0 then
    grActGrid.Cells[ColCompletion, IDTop] :=
      IntToStr(CompletionTotal div IncPrior)
  else
    grActGrid.Cells[ColCompletion, IDTop] := '0';
  if CostValue = True then
    grActGrid.Cells[ColCost, IDTop] := FormatCurr('#,#.##', Cost)
  else
    grActGrid.Cells[ColCost, IDTop] := '';
  // Recursion
  if IDTop - 1 > 0 then
    if GetCode(grActGrid.Cells[ColCode, IDTop]) > 1 then
      UpdateMainAct(IDTop);
  grActGrid.Refresh;
end;

procedure TfmMain.ShiftForwardActDates;
var
  i: integer;
begin
  // Shift forward the dates of the activities
  for i := 1 to LastRowAct do
  begin
    if grActGrid.Cells[ColStartDate, i] <> '' then
    begin
      grActGrid.Cells[ColStartDate, i] :=
        DateTimeToStr(StrToDateTime(grActGrid.Cells[ColStartDate, i], FDate) + 1, FDate);
    end;
    if grActGrid.Cells[ColEndDate, i] <> '' then
    begin
      grActGrid.Cells[ColEndDate, i] :=
        DateTimeToStr(StrToDateTime(grActGrid.Cells[ColEndDate, i], FDate) + 1, FDate);
    end;
  end;
  CalcActDates;
  grActGrid.Refresh;
  EditNotesDataset;
end;

procedure TfmMain.ShiftBackwardActDates;
var
  i: integer;
begin
  // Shift backward the dates of the activities
  for i := 1 to LastRowAct do
  begin
    if grActGrid.Cells[ColStartDate, i] <> '' then
    begin
      grActGrid.Cells[ColStartDate, i] :=
        DateTimeToStr(StrToDateTime(grActGrid.Cells[ColStartDate, i], FDate) - 1, FDate);
    end;
    if grActGrid.Cells[ColEndDate, i] <> '' then
    begin
      grActGrid.Cells[ColEndDate, i] :=
        DateTimeToStr(StrToDateTime(grActGrid.Cells[ColEndDate, i], FDate) - 1, FDate);
    end;
  end;
  CalcActDates;
  grActGrid.Refresh;
  EditNotesDataset;
end;

function TfmMain.GetCode(CellVal: string): integer;
begin
  // Get code in ColCode
  if CellVal = '' then
    Result := 0
  else
    Result := StrToInt(Copy(CellVal, 2, Length(CellVal)));
end;

function TfmMain.GetSign(CellVal: string): string;
begin
  // Get sign in ColCode
  if CellVal = '' then
    Result := ''
  else
    Result := Copy(CellVal, 1, 1);
end;

function TfmMain.ConvertDateFormat(Data, FileFormat, SftFormat: string): string;
var
  YMD, MDY, DMY: TFormatSettings;
begin
  // Convert among different date format
  YMD := DefaultFormatSettings;
  YMD.ShortDateFormat := 'yyyy-mm-dd';
  YMD.DateSeparator := '-';
  MDY := DefaultFormatSettings;
  MDY.ShortDateFormat := 'mm-dd-yyyy';
  MDY.DateSeparator := '-';
  DMY := DefaultFormatSettings;
  DMY.ShortDateFormat := 'dd-mm-yyyy';
  DMY.DateSeparator := '-';
  if Data = '' then
    Result := ''
  else if FileFormat = '' then
    Result := Data
  else if FileFormat = SftFormat then
    Result := Data
  else if ((FileFormat = 'YMD') and (SftFormat = 'MDY')) then
    Result := DateToStr(StrToDate(Data, YMD), MDY)
  else if ((FileFormat = 'YMD') and (SftFormat = 'DMY')) then
    Result := DateToStr(StrToDate(Data, YMD), DMY)
  else if ((FileFormat = 'MDY') and (SftFormat = 'DMY')) then
    Result := DateToStr(StrToDate(Data, MDY), DMY)
  else if ((FileFormat = 'MDY') and (SftFormat = 'YMD')) then
    Result := DateToStr(StrToDate(Data, MDY), YMD)
  else if ((FileFormat = 'DMY') and (SftFormat = 'YMD')) then
    Result := DateToStr(StrToDate(Data, DMY), YMD)
  else if ((FileFormat = 'DMY') and (SftFormat = 'MDY')) then
    Result := DateToStr(StrToDate(Data, DMY), MDY);
end;

procedure TfmMain.CalcActDates;
var
  dtStartDate, dtActStartDate, dtEndDate, dtActEndDate: TDate;
  stStartDate, stEndDate: string;
  i: integer;
begin
  dtStartDate := 10000000;
  dtEndDate := 0;
  // Calculate the start and end date of the project
  for i := 1 to LastRowAct do
    try
      if grActGrid.Cells[ColStartDate, i] <> '' then
      begin
        if TryStrToDate(grActGrid.Cells[ColStartDate, i], dtActStartDate, FDate) = True then
        begin
          if dtActStartDate < dtStartDate then
            dtStartDate := dtActStartDate;
        end;
      end;
      if grActGrid.Cells[ColEndDate, i] <> '' then
      begin
        if TryStrToDate(grActGrid.Cells[ColEndDate, i], dtActEndDate, FDate) = True then
        begin
          if dtActEndDate > dtEndDate then
            dtEndDate := dtActEndDate;
        end;
      end;
    except;
    end;
  stStartDate := '';
  stEndDate := '';
  if dtStartDate < 10000000 then
    stStartDate := FormatDateTime(FDate.LongDateFormat, dtStartDate);
  if dtEndDate > 0 then
    stEndDate := FormatDateTime(FDate.LongDateFormat, dtEndDate);
  lbActDates.Caption := '';
  if stStartDate <> '' then
  begin
    if stEndDate = '' then
      lbActDates.Caption := cpt091 + ' ' + stStartDate + '.'
    else
      lbActDates.Caption := cpt091 + ' ' + stStartDate + ' ' + cpt092 + ' ' + stEndDate + '.';
  end
  else if stEndDate <> '' then
    lbActDates.Caption := cpt093 + ' ' + stEndDate + '.'
  else
    lbActDates.Caption := cpt094 + '.';
end;

procedure TfmMain.StoreUndoData;
begin
  // Store note text for possible undo
  iNoteTextPos := dbText.SelStart;
  dbText.GetUndo(iNoteTextPos);
  miNotesUndo.Enabled := True;
end;

procedure TfmMain.GetUndoData;
begin
  // Replace current note text with the one stored for undo
  dbText.SetUndo(iNoteTextPos);
  // To avoid the cursor go to the bottom
  dbText.SetCursorMiddleScreen(dbText.CaretPos.Y);
end;

procedure TfmMain.ClearUndoData;
begin
  // Clear stored undo data of current note text
  iNoteTextPos := 0;
  dbText.ClearUndo;
  miNotesUndo.Enabled := False;
end;

function TfmMain.ExpTextToZim(NoteText: string): string;
var
  i: integer;
  flTag: boolean;
  slResLines: TStringList;
begin
  NoteText := StringReplace(NoteText, '<b> ', '<b>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<b>', '**', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, ' </b>', '</b>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '</b>', '**', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<i> ', '<i>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<i>', '//', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, ' </i>', '</i>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '</i>', '//', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<u> ', '<u>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<u>', '//', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, ' </u>', '</u>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '</u>', '//', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<strike> ', '<strike>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<strike>', '~~', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, ' </strike>', '</strike>', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '</strike>', '~~', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<span style="background', '__<', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '</span>', '__', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<IMG SRC="', '{{..' + DirectorySeparator + '..' + DirectorySeparator + 'ExportZimFiles' + DirectorySeparator, [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '.jpeg">', '.jpeg}}', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '.jpg">', '.jpg}}', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '.png">', '.png}}', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, LineEnding, '', [rfIgnoreCase, rfReplaceAll]);
  NoteText := StringReplace(NoteText, '<p align=', LineEnding + '<', [rfIgnoreCase, rfReplaceAll]);
  flTag := False;
  Result := '';
  for i := 1 to UTF8Length(NoteText) do
  begin
    if UTF8Copy(NoteText, i, 1) = '<' then
      flTag := True
    else if UTF8Copy(NoteText, i, 1) = '>' then
      flTag := False
    else if flTag = False then
      Result := Result + UTF8Copy(NoteText, i, 1);
  end;
  Result := StringReplace(Result, '•' + #8, '• ', [rfIgnoreCase, rfReplaceAll]);
  Result := StringReplace(Result, '.' + #8, '. ', [rfIgnoreCase, rfReplaceAll]);
  try
    slResLines := TStringList.Create;
    slResLines.Text := Result;
    i := 0;
    while i < slResLines.Count do
    begin
      if slResLines[i] = '**' then
      begin
        slResLines[i] := '';
        Inc(i);
        slResLines[i] := '**' + slResLines[i];
      end
      else if slResLines[i] = '//' then
      begin
        slResLines[i] := '';
        Inc(i);
        slResLines[i] := '//' + slResLines[i];
      end
      else if slResLines[i] = '~~' then
      begin
        slResLines[i] := '';
        Inc(i);
        slResLines[i] := '~~' + slResLines[i];
      end
      else
      begin
        Inc(i);
      end;
    end;
    Result := slResLines.Text;
  finally
    slResLines.Free;
  end;
end;

function TfmMain.SpaceToLine(stName: string): string;
begin
  Result := StringReplace(stName, ' ', '_', [rfReplaceAll]);
end;

// *****************************************************************************
// ************************* COMMON SYNC PROCEDURES ****************************
// *****************************************************************************

procedure TfmMain.SyncDelRec(ReadFile, WriteFile: string);
var
  sqSyncRead, sqSyncWrite: TSqlite3Dataset;
begin
  // Purge and sync table DelRec
  try
    sqSyncRead := TSqlite3Dataset.Create(Self);
    sqSyncWrite := TSqlite3Dataset.Create(Self);
    sqSyncRead.FileName := ReadFile;
    sqSyncRead.TableName := 'DelRec';
    sqSyncRead.PrimaryKey := 'IDDelRec';
    sqSyncRead.AutoIncrementKey := True;
    sqSyncRead.Open;
    sqSyncWrite.FileName := WriteFile;
    sqSyncWrite.TableName := 'DelRec';
    sqSyncWrite.PrimaryKey := 'IDDelRec';
    sqSyncWrite.AutoIncrementKey := True;
    sqSyncWrite.Open;
    // Purge from records older than DelayDays days
    while not sqSyncRead.EOF do
    begin
      if sqSyncRead.FieldByName('DelRecDTMod').AsDateTime < Now - DelayDays then
      begin
        sqSyncRead.Delete;
        sqSyncRead.ApplyUpdates;
      end
      else
      begin
        sqSyncRead.Next;
      end;
      Application.ProcessMessages;
    end;
    // Sync
    sqSyncRead.First;
    while not sqSyncRead.EOF do
    begin
      if sqSyncWrite.Locate('DelRecUID', sqSyncRead.FieldByName('DelRecUID').AsString, []) = False then
      begin
        sqSyncWrite.Append;
        sqSyncWrite.FieldByName('DelRecUID').AsString :=
          sqSyncRead.FieldByName('DelRecUID').AsString;
        sqSyncWrite.FieldByName('DelRecDTMod').AsDateTime :=
          sqSyncRead.FieldByName('DelRecDTMod').AsDateTime;
        sqSyncWrite.Post;
        sqSyncWrite.ApplyUpdates;
      end;
      sqSyncRead.Next;
      Application.ProcessMessages;
    end;
  finally
    sqSyncRead.Free;
    sqSyncWrite.Free;
  end;
end;

function TfmMain.SyncDelSubjectsNotes(ReadFile, WriteFile: string): integer;
var
  DelRecords, i: integer;
  sqSyncRead, sqSyncWrite: TSqlite3Dataset;
  AttDir: string;
  myStringList: TStringList;
begin
  try
    sqSyncRead := TSqlite3Dataset.Create(Self);
    sqSyncWrite := TSqlite3Dataset.Create(Self);
    DelRecords := 0;
    // Delete Notes in DelRec
    sqSyncRead.FileName := ReadFile;
    sqSyncRead.TableName := 'DelRec';
    sqSyncRead.PrimaryKey := 'IDDelRec';
    sqSyncRead.AutoIncrementKey := True;
    sqSyncRead.Open;
    sqSyncWrite.FileName := WriteFile;
    sqSyncWrite.TableName := 'Notes';
    sqSyncWrite.PrimaryKey := 'IDNotes';
    sqSyncWrite.AutoIncrementKey := True;
    sqSyncWrite.Open;
    while not sqSyncWrite.EOF do
    begin
      if sqSyncRead.Locate('DelRecUID', sqSyncWrite.FieldByName('NotesUID').AsString, []) = True then
      begin
        // Delete attachment
        if sqSyncWrite.FieldByName('NotesAttName').AsString <> '' then
          try
            AttDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
            if DirectoryExistsUTF8(AttDir) = False then
            begin
              MessageDlg(msg030, mtWarning, [mbOK], 0);
              Abort;
            end;
            myStringList := TStringList.Create;
            myStringList.Text := sqSyncWrite.FieldByName('NotesAttName').AsString;
            for i := 0 to myStringList.Count - 1 do
              DeleteFileUTF8(AttDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip');
            myStringList.Free;
            if fmMain.IsDirectoryEmpty(AttDir) = True then
              DeleteDirectory(AttDir, False);
          except
            MessageDlg(msg035, mtWarning, [mbOK], 0);
            Abort;
          end;
        // Delete images
        AttDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
        if DirectoryExistsUTF8(AttDir) = True then
        begin
          i := 0;
          try
            while FileExistsUTF8(AttDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg') = True do
            begin
              DeleteFileUTF8(AttDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
              Inc(i);
            end;
            if IsDirectoryEmpty(AttDir) = True then
              DeleteDirectory(AttDir, False);
          except
            MessageDlg(msg035, mtWarning, [mbOK], 0);
            Abort;
          end;
        end;
        sqSyncWrite.Delete;
        sqSyncWrite.ApplyUpdates;
        Inc(DelRecords);
      end
      else
      begin
        sqSyncWrite.Next;
      end;
      Application.ProcessMessages;
    end;
    sqSyncWrite.Free;
    // Delete Subjects in DelRec
    sqSyncWrite := TSqlite3Dataset.Create(Self);
    sqSyncWrite.FileName := WriteFile;
    sqSyncWrite.AutoIncrementKey := True;
    sqSyncWrite.TableName := 'Subjects';
    sqSyncWrite.PrimaryKey := 'IDSubjects';
    sqSyncWrite.Open;
    while not sqSyncWrite.EOF do
    begin
      if sqSyncRead.Locate('DelRecUID', sqSyncWrite.FieldByName('SubjectsUID').AsString, []) = True then
      begin
        sqSyncWrite.Delete;
        sqSyncWrite.ApplyUpdates;
        Inc(DelRecords);
      end
      else
      begin
        sqSyncWrite.Next;
      end;
      Application.ProcessMessages;
    end;
    Result := DelRecords;
  finally
    sqSyncRead.Free;
    sqSyncWrite.Free;
  end;
end;

function TfmMain.SyncUpdateSubjectsNotes(ReadFile, WriteFile: string): integer;
var
  ChnRecords, i: integer;
  sqSyncRead, sqSyncWrite, sqCheckIDSubjects: TSqlite3Dataset;
  AttReadDir, AttWriteDir: string;
  myStringList: TStringList;
begin
  if DirectoryExistsUTF8(ExtractFileNameWithoutExt(ReadFile)) = True then
    try
      if DirectoryExistsUTF8(ExtractFileNameWithoutExt(WriteFile)) = False then
        CreateDirUTF8(ExtractFileNameWithoutExt(WriteFile));
    except
      MessageDlg(msg030, mtWarning, [mbOK], 0);
    end;
  // Sync subjects
  ChnRecords := 0;
  try
    sqSyncRead := TSqlite3Dataset.Create(Self);
    sqSyncWrite := TSqlite3Dataset.Create(Self);
    sqSyncRead.FileName := ReadFile;
    sqSyncRead.TableName := 'Subjects';
    sqSyncRead.PrimaryKey := 'IDSubjects';
    sqSyncRead.AutoIncrementKey := True;
    sqSyncRead.Open;
    sqSyncWrite.FileName := WriteFile;
    sqSyncWrite.TableName := 'Subjects';
    sqSyncWrite.PrimaryKey := 'IDSubjects';
    sqSyncWrite.AutoIncrementKey := True;
    sqSyncWrite.Open;
    while not sqSyncRead.EOF do
    begin
      // Subject not present: add it
      if sqSyncWrite.Locate('SubjectsUID', sqSyncRead.FieldByName('SubjectsUID').AsString, []) = False then
      begin
        sqSyncWrite.Append;
        sqSyncWrite.FieldByName('SubjectsName').AsString :=
          sqSyncRead.FieldByName('SubjectsName').AsString;
        sqSyncWrite.FieldByName('SubjectsComments').AsString :=
          sqSyncRead.FieldByName('SubjectsComments').AsString;
        sqSyncWrite.FieldByName('SubjectsBackColor').AsString :=
          sqSyncRead.FieldByName('SubjectsBackColor').AsString;
        sqSyncWrite.FieldByName('SubjectsFontColor').AsString :=
          sqSyncRead.FieldByName('SubjectsFontColor').AsString;
        sqSyncWrite.FieldByName('SubjectsSort').AsInteger :=
          sqSyncWrite.FieldByName('IDSubjects').AsInteger;
        sqSyncWrite.FieldByName('SubjectsUID').AsString :=
          sqSyncRead.FieldByName('SubjectsUID').AsString;
        sqSyncWrite.FieldByName('SubjectsDTMod').AsDateTime :=
          sqSyncRead.FieldByName('SubjectsDTMod').AsDateTime;
        sqSyncWrite.Post;
        sqSyncWrite.ApplyUpdates;
        Inc(ChnRecords);
      end
      // Subject present: check date
      else if sqSyncRead.FieldByName('SubjectsDTMod').AsDateTime > sqSyncWrite.FieldByName('SubjectsDTMod').AsDateTime then
      begin
        sqSyncWrite.Edit;
        sqSyncWrite.FieldByName('SubjectsName').AsString :=
          sqSyncRead.FieldByName('SubjectsName').AsString;
        sqSyncWrite.FieldByName('SubjectsComments').AsString :=
          sqSyncRead.FieldByName('SubjectsComments').AsString;
        sqSyncWrite.FieldByName('SubjectsBackColor').AsString :=
          sqSyncRead.FieldByName('SubjectsBackColor').AsString;
        sqSyncWrite.FieldByName('SubjectsFontColor').AsString :=
          sqSyncRead.FieldByName('SubjectsFontColor').AsString;
        sqSyncWrite.FieldByName('SubjectsDTMod').AsDateTime :=
          sqSyncRead.FieldByName('SubjectsDTMod').AsDateTime;
        sqSyncWrite.Post;
        sqSyncWrite.ApplyUpdates;
        Inc(ChnRecords);
      end;
      sqSyncRead.Next;
      Application.ProcessMessages;
    end;
  finally
    sqSyncRead.Free;
    sqSyncWrite.Free;
  end;
  // Sync notes
  try
    sqSyncRead := TSqlite3Dataset.Create(Self);
    sqSyncWrite := TSqlite3Dataset.Create(Self);
    sqCheckIDSubjects := TSqlite3Dataset.Create(Self);
    sqSyncRead.FileName := ReadFile;
    sqSyncRead.TableName := 'Notes';
    sqSyncRead.PrimaryKey := 'IDNotes';
    sqSyncRead.SQL := 'Select Notes.IDNotes, Notes.ID_Subjects, Notes.NotesTitle, ' + 'Notes.NotesDate, Notes.NotesText, Notes.NotesTags, Notes.NotesBackColor, ' + 'Notes.NotesFontColor, Notes.NotesSort, Notes.NotesAttName, ' + 'Notes.NotesUID, Notes.NotesDTMod, Notes.NotesCheckPwd, ' + 'Notes.NotesActivities, Notes.NotesDateFormat, ' + 'Subjects.IDSubjects, Subjects.SubjectsUID from Subjects, Notes ' + 'where Notes.ID_Subjects = Subjects.IDSubjects';
    sqSyncRead.AutoIncrementKey := True;
    sqSyncRead.Open;
    sqSyncWrite.FileName := WriteFile;
    sqSyncWrite.TableName := 'Notes';
    sqSyncWrite.PrimaryKey := 'IDNotes';
    sqSyncWrite.AutoIncrementKey := True;
    sqSyncWrite.Open;
    sqCheckIDSubjects.FileName := WriteFile;
    sqCheckIDSubjects.FileName := sqSyncWrite.FileName;
    sqCheckIDSubjects.TableName := 'Subjects';
    sqCheckIDSubjects.PrimaryKey := 'IDSubjects';
    sqCheckIDSubjects.AutoIncrementKey := True;
    sqCheckIDSubjects.Open;
    while not sqSyncRead.EOF do
    begin
      // Note not present: add it
      if sqSyncWrite.Locate('NotesUID', sqSyncRead.FieldByName('NotesUID').AsString, []) = False then
      begin
        sqCheckIDSubjects.Locate('SubjectsUID',
          sqSyncRead.FieldByName('SubjectsUID').AsString, []);
        // Delete attachments in the writer directory:
        // they will be copied from reader directory
        if sqSyncWrite.FieldByName('NotesAttName').AsString <> '' then
        begin
          AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
          begin
            MessageDlg(msg030, mtWarning, [mbOK], 0);
            Abort;
          end;
          myStringList := TStringList.Create;
          myStringList.Text := sqSyncWrite.FieldByName('NotesAttName').AsString;
          for i := 0 to myStringList.Count - 1 do
            DeleteFileUTF8(AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip');
          myStringList.Free;
          if IsDirectoryEmpty(AttWriteDir) = True then
            DeleteDirectory(AttWriteDir, False);
        end;
        // Delete images in the writer directory:
        // they will be copied from reader directory
        AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
        if DirectoryExistsUTF8(AttWriteDir) = True then
        begin
          i := 0;
          try
            while FileExistsUTF8(AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg') = True do
            begin
              DeleteFileUTF8(AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
              Inc(i);
            end;
            if IsDirectoryEmpty(AttWriteDir) = True then
              DeleteDirectory(AttWriteDir, False);
          except
            MessageDlg(msg035, mtWarning, [mbOK], 0);
            Abort;
          end;
        end;
        sqSyncWrite.Append;
        sqSyncWrite.FieldByName('ID_Subjects').AsInteger :=
          sqCheckIDSubjects.FieldByName('IDSubjects').AsInteger;
        sqSyncWrite.FieldByName('NotesTitle').AsString :=
          sqSyncRead.FieldByName('NotesTitle').AsString;
        sqSyncWrite.FieldByName('NotesDate').AsDateTime :=
          sqSyncRead.FieldByName('NotesDate').AsDateTime;
        sqSyncWrite.FieldByName('NotesText').AsString :=
          sqSyncRead.FieldByName('NotesText').AsString;
        sqSyncWrite.FieldByName('NotesActivities').AsString :=
          sqSyncRead.FieldByName('NotesActivities').AsString;
        sqSyncWrite.FieldByName('NotesTags').AsString :=
          sqSyncRead.FieldByName('NotesTags').AsString;
        sqSyncWrite.FieldByName('NotesBackColor').AsString :=
          sqSyncRead.FieldByName('NotesBackColor').AsString;
        sqSyncWrite.FieldByName('NotesFontColor').AsString :=
          sqSyncRead.FieldByName('NotesFontColor').AsString;
        sqSyncWrite.FieldByName('NotesSort').AsInteger :=
          sqSyncWrite.FieldByName('IDNotes').AsInteger;
        sqSyncWrite.FieldByName('NotesAttName').AsString :=
          sqSyncRead.FieldByName('NotesAttName').AsString;
        sqSyncWrite.FieldByName('NotesUID').AsString :=
          sqSyncRead.FieldByName('NotesUID').AsString;
        sqSyncWrite.FieldByName('NotesDTMod').AsDateTime :=
          sqSyncRead.FieldByName('NotesDTMod').AsDateTime;
        sqSyncWrite.FieldByName('NotesDateFormat').AsString :=
          sqSyncRead.FieldByName('NotesDateFormat').AsString;
        sqSyncWrite.FieldByName('NotesCheckPwd').AsString :=
          sqSyncRead.FieldByName('NotesCheckPwd').AsString;
        sqSyncWrite.Post;
        sqSyncWrite.ApplyUpdates;
        // Copy attachments
        if sqSyncRead.FieldByName('NotesAttName').AsString <> '' then
        begin
          AttReadDir := ExtractFileNameWithoutExt(sqSyncRead.FileName);
          if DirectoryExistsUTF8(AttReadDir) = False then
          begin
            MessageDlg(msg030, mtWarning, [mbOK], 0);
            Abort;
          end;
          AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
            try
              CreateDirUTF8(AttWriteDir);
            except
              MessageDlg(msg029, mtWarning, [mbOK], 0);
              Screen.Cursor := crDefault;
              Abort;
            end;
          myStringList := TStringList.Create;
          myStringList.Text := sqSyncRead.FieldByName('NotesAttName').AsString;
          for i := 0 to myStringList.Count - 1 do
            if FileExistsUTF8(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip') then
              CopyFile(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip',
                AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip');
          myStringList.Free;
        end;
        // Copy images
        AttReadDir := ExtractFileNameWithoutExt(sqSyncRead.FileName);
        if DirectoryExistsUTF8(AttReadDir) = True then
        begin
          AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
            try
              CreateDirUTF8(AttWriteDir);
            except
              MessageDlg(msg029, mtWarning, [mbOK], 0);
              Screen.Cursor := crDefault;
              Abort;
            end;
          i := 0;
          try
            while FileExistsUTF8(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg') = True do
            begin
              CopyFile(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg',
                AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
              Inc(i);
            end;
          except
            MessageDlg(msg035, mtWarning, [mbOK], 0);
            Abort;
          end;
        end;
        Inc(ChnRecords);
      end
      // Note present: check date
      else if sqSyncRead.FieldByName('NotesDTMod').AsDateTime > sqSyncWrite.FieldByName('NotesDTMod').AsDateTime then
      begin
        // Delete attachments in the writer directory:
        // they will be copied from reader directory
        if sqSyncWrite.FieldByName('NotesAttName').AsString <> '' then
        begin
          AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
          begin
            MessageDlg(msg030, mtWarning, [mbOK], 0);
            Abort;
          end;
          myStringList := TStringList.Create;
          myStringList.Text := sqSyncWrite.FieldByName('NotesAttName').AsString;
          for i := 0 to myStringList.Count - 1 do
            DeleteFileUTF8(AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip');
          myStringList.Free;
          if IsDirectoryEmpty(AttWriteDir) = True then
            DeleteDirectory(AttWriteDir, False);
        end;
        // Delete images in the writer directory:
        // they will be copied from reader directory
        AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
        if DirectoryExistsUTF8(AttWriteDir) = True then
        begin
          i := 0;
          try
            while FileExistsUTF8(AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg') = True do
            begin
              DeleteFileUTF8(AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
              Inc(i);
            end;
            if IsDirectoryEmpty(AttWriteDir) = True then
              DeleteDirectory(AttWriteDir, False);
          except
            MessageDlg(msg035, mtWarning, [mbOK], 0);
            Abort;
          end;
        end;
        sqSyncWrite.Edit;
        sqSyncWrite.FieldByName('NotesTitle').AsString :=
          sqSyncRead.FieldByName('NotesTitle').AsString;
        sqSyncWrite.FieldByName('NotesDate').AsDateTime :=
          sqSyncRead.FieldByName('NotesDate').AsDateTime;
        sqSyncWrite.FieldByName('NotesText').AsString :=
          sqSyncRead.FieldByName('NotesText').AsString;
        sqSyncWrite.FieldByName('NotesActivities').AsString :=
          sqSyncRead.FieldByName('NotesActivities').AsString;
        sqSyncWrite.FieldByName('NotesTags').AsString :=
          sqSyncRead.FieldByName('NotesTags').AsString;
        sqSyncWrite.FieldByName('NotesBackColor').AsString :=
          sqSyncRead.FieldByName('NotesBackColor').AsString;
        sqSyncWrite.FieldByName('NotesFontColor').AsString :=
          sqSyncRead.FieldByName('NotesFontColor').AsString;
        sqSyncWrite.FieldByName('NotesAttName').AsString :=
          sqSyncRead.FieldByName('NotesAttName').AsString;
        sqSyncWrite.FieldByName('NotesDTMod').AsDateTime :=
          sqSyncRead.FieldByName('NotesDTMod').AsDateTime;
        sqSyncWrite.FieldByName('NotesCheckPwd').AsString :=
          sqSyncRead.FieldByName('NotesCheckPwd').AsString;
        sqSyncWrite.FieldByName('NotesDateFormat').AsString :=
          sqSyncRead.FieldByName('NotesDateFormat').AsString;
        sqSyncWrite.Post;
        sqSyncWrite.ApplyUpdates;
        // Copy attachments
        if sqSyncRead.FieldByName('NotesAttName').AsString <> '' then
        begin
          AttReadDir := ExtractFileNameWithoutExt(sqSyncRead.FileName);
          if DirectoryExistsUTF8(AttReadDir) = False then
          begin
            MessageDlg(msg030, mtWarning, [mbOK], 0);
            Abort;
          end;
          AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
            try
              CreateDirUTF8(AttWriteDir);
            except
              MessageDlg(msg029, mtWarning, [mbOK], 0);
              Screen.Cursor := crDefault;
              Abort;
            end;
          myStringList := TStringList.Create;
          myStringList.Text := sqSyncRead.FieldByName('NotesAttName').AsString;
          for i := 0 to myStringList.Count - 1 do
            if FileExistsUTF8(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip') then
              CopyFile(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip',
                AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[i]) + '.zip');
          myStringList.Free;
        end;
        // Copy images
        AttReadDir := ExtractFileNameWithoutExt(sqSyncRead.FileName);
        if DirectoryExistsUTF8(AttReadDir) = True then
        begin
          AttWriteDir := ExtractFileNameWithoutExt(sqSyncWrite.FileName);
          if DirectoryExistsUTF8(AttWriteDir) = False then
            try
              CreateDirUTF8(AttWriteDir);
            except
              MessageDlg(msg029, mtWarning, [mbOK], 0);
              Screen.Cursor := crDefault;
              Abort;
            end;
          i := 0;
          try
            while FileExistsUTF8(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg') = True do
            begin
              CopyFile(AttReadDir + DirectorySeparator + sqSyncRead.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg',
                AttWriteDir + DirectorySeparator + sqSyncWrite.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', i) + '.jpeg');
              Inc(i);
            end;
          except
            MessageDlg(msg035, mtWarning, [mbOK], 0);
            Abort;
          end;
        end;
        Inc(ChnRecords);
      end;
      sqSyncRead.Next;
      Application.ProcessMessages;
    end;
    Result := ChnRecords;
  finally
    sqSyncRead.Free;
    sqSyncWrite.Free;
    sqCheckIDSubjects.Free;
  end;
end;

// *****************************************************************************
// ****************************** TRANSLATION **********************************
// *****************************************************************************

procedure TfmMain.Translation;
var
  MyIni: TIniFile;
  myHomeDir: string;
begin
  // Set home directory and data directories
  myHomeDir := GetEnvironmentVariable('HOME') + '/.config';
  // Copy file if existing in installation directory;
  // Directory mynotex has been already created
  if FileExistsUTF8(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'translation-' + VersMyNt) = False then
    try
      if FileExistsUTF8(InstallDir + LowerCase(Copy(GetEnvironmentVariable('LANG'), 1, 2)) + '.lng') then
      begin
        CopyFile(InstallDir + LowerCase(Copy(GetEnvironmentVariable('LANG'), 1, 2) + '.lng'),
          myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'translation-' + VersMyNt);
      end;
    except;
    end;
  // Load translation file if existing
  if FileExistsUTF8(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'translation-' + VersMyNt) then
  begin
    MyIni := TIniFile.Create(myHomeDir + DirectorySeparator + 'mynotex' + DirectorySeparator + 'translation-' + VersMyNt);
    try
      // Messages dialogs
      msg001 := MyIni.ReadString('mynotex', 'msg001', '');
      msg002 := MyIni.ReadString('mynotex', 'msg002', '');
      msg003 := MyIni.ReadString('mynotex', 'msg003', '');
      msg004 := MyIni.ReadString('mynotex', 'msg004', '');
      msg005 := MyIni.ReadString('mynotex', 'msg005', '');
      msg006 := MyIni.ReadString('mynotex', 'msg006', '');
      msg007 := MyIni.ReadString('mynotex', 'msg007', '');
      msg008 := MyIni.ReadString('mynotex', 'msg008', '');
      msg009 := MyIni.ReadString('mynotex', 'msg009', '');
      msg010 := MyIni.ReadString('mynotex', 'msg010', '');
      msg011 := MyIni.ReadString('mynotex', 'msg011', '');
      msg012 := MyIni.ReadString('mynotex', 'msg012', '');
      msg013 := MyIni.ReadString('mynotex', 'msg013', '');
      msg014 := MyIni.ReadString('mynotex', 'msg014', '');
      msg015 := MyIni.ReadString('mynotex', 'msg015', '');
      msg016 := MyIni.ReadString('mynotex', 'msg016', '');
      msg017 := MyIni.ReadString('mynotex', 'msg017', '');
      msg018 := MyIni.ReadString('mynotex', 'msg018', '');
      msg019 := MyIni.ReadString('mynotex', 'msg019', '');
      msg020 := MyIni.ReadString('mynotex', 'msg020', '');
      msg021 := MyIni.ReadString('mynotex', 'msg021', '');
      msg022 := MyIni.ReadString('mynotex', 'msg022', '');
      msg023 := MyIni.ReadString('mynotex', 'msg023', '');
      msg024 := MyIni.ReadString('mynotex', 'msg024', '');
      msg025 := MyIni.ReadString('mynotex', 'msg025', '');
      // msg026 := MyIni.ReadString('mynotex','msg026','');
      msg027 := MyIni.ReadString('mynotex', 'msg027', '');
      msg028 := MyIni.ReadString('mynotex', 'msg028', '');
      msg029 := MyIni.ReadString('mynotex', 'msg029', '');
      msg030 := MyIni.ReadString('mynotex', 'msg030', '');
      msg031 := MyIni.ReadString('mynotex', 'msg031', '');
      msg032 := MyIni.ReadString('mynotex', 'msg032', '');
      msg033 := MyIni.ReadString('mynotex', 'msg033', '');
      msg034 := MyIni.ReadString('mynotex', 'msg034', '');
      msg035 := MyIni.ReadString('mynotex', 'msg035', '');
      msg036 := MyIni.ReadString('mynotex', 'msg036', '');
      msg037 := MyIni.ReadString('mynotex', 'msg037', '');
      msg038 := MyIni.ReadString('mynotex', 'msg038', '');
      msg039 := MyIni.ReadString('mynotex', 'msg039', '');
      msg040 := MyIni.ReadString('mynotex', 'msg040', '');
      msg041 := MyIni.ReadString('mynotex', 'msg041', '');
      msg042 := MyIni.ReadString('mynotex', 'msg042', '');
      msg043 := MyIni.ReadString('mynotex', 'msg043', '');
      msg044 := MyIni.ReadString('mynotex', 'msg044', '');
      msg045 := MyIni.ReadString('mynotex', 'msg045', '');
      msg046 := MyIni.ReadString('mynotex', 'msg046', '');
      msg047 := MyIni.ReadString('mynotex', 'msg047', '');
      msg048 := MyIni.ReadString('mynotex', 'msg048', '');
      msg049 := MyIni.ReadString('mynotex', 'msg049', '');
      msg050 := MyIni.ReadString('mynotex', 'msg050', '');
      msg051 := MyIni.ReadString('mynotex', 'msg051', '');
      msg052 := MyIni.ReadString('mynotex', 'msg052', '');
      msg053 := MyIni.ReadString('mynotex', 'msg053', '');
      msg054 := MyIni.ReadString('mynotex', 'msg054', '');
      msg055 := MyIni.ReadString('mynotex', 'msg055', '');
      msg056 := MyIni.ReadString('mynotex', 'msg056', '');
      msg057 := MyIni.ReadString('mynotex', 'msg057', '');
      msg058 := MyIni.ReadString('mynotex', 'msg058', '');
      msg059 := MyIni.ReadString('mynotex', 'msg059', '');
      msg060 := MyIni.ReadString('mynotex', 'msg060', '');
      msg061 := MyIni.ReadString('mynotex', 'msg061', '');
      msg062 := MyIni.ReadString('mynotex', 'msg062', '');
      msg063 := MyIni.ReadString('mynotex', 'msg063', '');
      msg064 := MyIni.ReadString('mynotex', 'msg064', '');
      msg065 := MyIni.ReadString('mynotex', 'msg065', '');
      msg066 := MyIni.ReadString('mynotex', 'msg066', '');
      // Not used from 1.4 version
      // msg067 := MyIni.ReadString('mynotex','msg067','');
      msg068 := MyIni.ReadString('mynotex', 'msg068', '');
      msg069 := MyIni.ReadString('mynotex', 'msg069', '');
      msg070 := MyIni.ReadString('mynotex', 'msg070', '');
      msg071 := MyIni.ReadString('mynotex', 'msg071', '');
      msg072 := MyIni.ReadString('mynotex', 'msg072', '');
      msg073 := MyIni.ReadString('mynotex', 'msg073', '');
      msg074 := MyIni.ReadString('mynotex', 'msg074', '');
      msg075 := MyIni.ReadString('mynotex', 'msg075', '');
      msg076 := MyIni.ReadString('mynotex', 'msg076', '');
      msg077 := MyIni.ReadString('mynotex', 'msg077', '');
      msg078 := MyIni.ReadString('mynotex', 'msg078', '');
      msg079 := MyIni.ReadString('mynotex', 'msg079', '');
      msg080 := MyIni.ReadString('mynotex', 'msg080', '');
      msg081 := MyIni.ReadString('mynotex', 'msg081', '');
      msg082 := MyIni.ReadString('mynotex', 'msg082', '');
      msg083 := MyIni.ReadString('mynotex', 'msg083', '');
      msg084 := MyIni.ReadString('mynotex', 'msg084', '');
      // Status bar messages
      sbr001 := MyIni.ReadString('mynotex', 'sbr001', '');
      sbr002 := MyIni.ReadString('mynotex', 'sbr002', '');
      sbr003 := MyIni.ReadString('mynotex', 'sbr003', '');
      sbr004 := MyIni.ReadString('mynotex', 'sbr004', '');
      sbr005 := MyIni.ReadString('mynotex', 'sbr005', '');
      sbr006 := MyIni.ReadString('mynotex', 'sbr006', '');
      sbr007 := MyIni.ReadString('mynotex', 'sbr007', '');
      sbr008 := MyIni.ReadString('mynotex', 'sbr008', '');
      sbr009 := MyIni.ReadString('mynotex', 'sbr009', '');
      sbr010 := MyIni.ReadString('mynotex', 'sbr010', '');
      sbr011 := MyIni.ReadString('mynotex', 'sbr011', '');
      // Various captions modified in the code
      cpt001 := MyIni.ReadString('mynotex', 'cpt001', '');
      cpt002 := MyIni.ReadString('mynotex', 'cpt002', '');
      cpt003 := MyIni.ReadString('mynotex', 'cpt003', '');
      cpt004 := MyIni.ReadString('mynotex', 'cpt004', '');
      cpt005 := MyIni.ReadString('mynotex', 'cpt005', '');
      cpt006 := MyIni.ReadString('mynotex', 'cpt006', '');
      cpt007 := MyIni.ReadString('mynotex', 'cpt007', '');
      cpt008 := MyIni.ReadString('mynotex', 'cpt008', '');
      cpt009 := MyIni.ReadString('mynotex', 'cpt009', '');
      cpt010 := MyIni.ReadString('mynotex', 'cpt010', '');
      cpt011 := MyIni.ReadString('mynotex', 'cpt011', '');
      cpt012 := MyIni.ReadString('mynotex', 'cpt012', '');
      cpt013 := MyIni.ReadString('mynotex', 'cpt013', '');
      cpt014 := MyIni.ReadString('mynotex', 'cpt014', '');
      cpt015 := MyIni.ReadString('mynotex', 'cpt015', '');
      cpt016 := MyIni.ReadString('mynotex', 'cpt016', '');
      cpt017 := MyIni.ReadString('mynotex', 'cpt017', '');
      cpt018 := MyIni.ReadString('mynotex', 'cpt018', '');
      cpt019 := MyIni.ReadString('mynotex', 'cpt019', '');
      cpt020 := MyIni.ReadString('mynotex', 'cpt020', '');
      cpt021 := MyIni.ReadString('mynotex', 'cpt021', '');
      cpt022 := MyIni.ReadString('mynotex', 'cpt022', '');
      cpt023 := MyIni.ReadString('mynotex', 'cpt023', '');
      cpt024 := MyIni.ReadString('mynotex', 'cpt024', '');
      cpt025 := MyIni.ReadString('mynotex', 'cpt025', '');
      cpt026 := MyIni.ReadString('mynotex', 'cpt026', '');
      cpt027 := MyIni.ReadString('mynotex', 'cpt027', '');
      cpt028 := MyIni.ReadString('mynotex', 'cpt028', '');
      cpt029 := MyIni.ReadString('mynotex', 'cpt029', '');
      cpt030 := MyIni.ReadString('mynotex', 'cpt030', '');
      cpt031 := MyIni.ReadString('mynotex', 'cpt031', '');
      cpt032 := MyIni.ReadString('mynotex', 'cpt032', '');
      // Labels not modified in English, but that cannot be changed
      // in the OnCreate of main form because fmCopyright is not still created
      cpt033 := MyIni.ReadString('mynotex', 'cpt033', '');
      cpt034 := MyIni.ReadString('mynotex', 'cpt034', '');
      cpt035 := MyIni.ReadString('mynotex', 'cpt035', '');
      cpt036 := MyIni.ReadString('mynotex', 'cpt036', '');
      cpt037 := MyIni.ReadString('mynotex', 'cpt037', '');
      cpt038 := MyIni.ReadString('mynotex', 'cpt038', '');
      cpt039 := MyIni.ReadString('mynotex', 'cpt039', '');
      cpt040 := MyIni.ReadString('mynotex', 'cpt040', '');
      cpt041 := MyIni.ReadString('mynotex', 'cpt041', '');
      cpt042 := MyIni.ReadString('mynotex', 'cpt042', '');
      cpt043 := MyIni.ReadString('mynotex', 'cpt043', '');
      cpt044 := MyIni.ReadString('mynotex', 'cpt044', '');
      cpt045 := MyIni.ReadString('mynotex', 'cpt045', '');
      cpt046 := MyIni.ReadString('mynotex', 'cpt046', '');
      cpt047 := MyIni.ReadString('mynotex', 'cpt047', '');
      cpt048 := MyIni.ReadString('mynotex', 'cpt048', '');
      cpt049 := MyIni.ReadString('mynotex', 'cpt049', '');
      cpt050 := MyIni.ReadString('mynotex', 'cpt050', '');
      cpt051 := MyIni.ReadString('mynotex', 'cpt051', '');
      cpt052 := MyIni.ReadString('mynotex', 'cpt052', '');
      cpt053 := MyIni.ReadString('mynotex', 'cpt053', '');
      cpt054 := MyIni.ReadString('mynotex', 'cpt054', '');
      cpt055 := MyIni.ReadString('mynotex', 'cpt055', '');
      cpt056 := MyIni.ReadString('mynotex', 'cpt056', '');
      cpt057 := MyIni.ReadString('mynotex', 'cpt057', '');
      cpt058 := MyIni.ReadString('mynotex', 'cpt058', '');
      cpt059 := MyIni.ReadString('mynotex', 'cpt059', '');
      cpt060 := MyIni.ReadString('mynotex', 'cpt060', '');
      cpt061 := MyIni.ReadString('mynotex', 'cpt061', '');
      cpt062 := MyIni.ReadString('mynotex', 'cpt062', '');
      cpt063 := MyIni.ReadString('mynotex', 'cpt063', '');
      cpt064 := MyIni.ReadString('mynotex', 'cpt064', '');
      cpt065 := MyIni.ReadString('mynotex', 'cpt065', '');
      cpt066 := MyIni.ReadString('mynotex', 'cpt066', '');
      cpt067 := MyIni.ReadString('mynotex', 'cpt067', '');
      cpt068 := MyIni.ReadString('mynotex', 'cpt068', '');
      cpt069 := MyIni.ReadString('mynotex', 'cpt069', '');
      cpt070 := MyIni.ReadString('mynotex', 'cpt070', '');
      cpt071 := MyIni.ReadString('mynotex', 'cpt071', '');
      cpt072 := MyIni.ReadString('mynotex', 'cpt072', '');
      cpt073 := MyIni.ReadString('mynotex', 'cpt073', '');
      cpt074 := MyIni.ReadString('mynotex', 'cpt074', '');
      cpt075 := MyIni.ReadString('mynotex', 'cpt075', '');
      cpt076 := MyIni.ReadString('mynotex', 'cpt076', '');
      cpt077 := MyIni.ReadString('mynotex', 'cpt077', '');
      cpt078 := MyIni.ReadString('mynotex', 'cpt078', '');
      cpt079 := MyIni.ReadString('mynotex', 'cpt079', '');
      cpt080 := MyIni.ReadString('mynotex', 'cpt080', '');
      cpt081 := MyIni.ReadString('mynotex', 'cpt081', '');
      cpt082 := MyIni.ReadString('mynotex', 'cpt082', '');
      cpt083 := MyIni.ReadString('mynotex', 'cpt083', '');
      cpt084 := MyIni.ReadString('mynotex', 'cpt084', '');
      cpt085 := MyIni.ReadString('mynotex', 'cpt085', '');
      cpt086 := MyIni.ReadString('mynotex', 'cpt086', '');
      cpt087 := MyIni.ReadString('mynotex', 'cpt087', '');
      cpt088 := MyIni.ReadString('mynotex', 'cpt088', '');
      cpt089 := MyIni.ReadString('mynotex', 'cpt089', '');
      cpt090 := MyIni.ReadString('mynotex', 'cpt090', '');
      cpt091 := MyIni.ReadString('mynotex', 'cpt091', '');
      cpt092 := MyIni.ReadString('mynotex', 'cpt092', '');
      cpt093 := MyIni.ReadString('mynotex', 'cpt093', '');
      cpt094 := MyIni.ReadString('mynotex', 'cpt094', '');
      cpt095 := MyIni.ReadString('mynotex', 'cpt095', '');
      cpt096 := MyIni.ReadString('mynotex', 'cpt096', '');
      cpt097 := MyIni.ReadString('mynotex', 'cpt097', '');
      cpt098 := MyIni.ReadString('mynotex', 'cpt098', '');
      cpt099 := MyIni.ReadString('mynotex', 'cpt099', '');
      // Components
      lbFound.Caption := MyIni.ReadString('mynotex', 'cpn001', '');
      cbFindKind.Items.Clear;
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn002', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn003', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn004', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn005', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn006', ''));
      // Added later
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn057', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn095', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn096', ''));
      cbFindKind.Items.Add(MyIni.ReadString('mynotex', 'cpn072', ''));
      grSubjects.Columns[0].Title.Caption := MyIni.ReadString('mynotex', 'cpn007', '');
      lbTitle.Caption := MyIni.ReadString('mynotex', 'cpn008', '');
      lbDate.Caption := MyIni.ReadString('mynotex', 'cpn009', '');
      grTitles.Cells[1, 0] := MyIni.ReadString('mynotex', 'cpn010', '');
      grTitles.Cells[2, 0] := MyIni.ReadString('mynotex', 'cpn011', '');
      bnFindFirst.Caption := MyIni.ReadString('mynotex', 'cpn012', '');
      bnFindNext.Caption := MyIni.ReadString('mynotex', 'cpn013', '');
      // Following item is unuseful because changed by code from ver. 1.0.6
      // odOpenDialog.Title := MyIni.ReadString('mynotex','cpn014','');
      OpenSaveDlgFilter := MyIni.ReadString('mynotex', 'cpn015', '');
      fdColorDialog.Title := MyIni.ReadString('mynotex', 'cpn016', '');
      fdColorFormatting.Title := MyIni.ReadString('mynotex', 'cpn016', '');
      fdFontSelDialog.Title := MyIni.ReadString('mynotex', 'cpn017', '');
      fdFontGenDialog.Title := MyIni.ReadString('mynotex', 'cpn017', '');
      // Added later
      lbTags.Caption := MyIni.ReadString('mynotex', 'cpn054', '');
      stAttachments := MyIni.ReadString('mynotex', 'cpn055', '');
      lbListAttach.Caption := MyIni.ReadString('mynotex', 'cpn055', '');
      // Menù items
      miFile.Caption := MyIni.ReadString('mynotex', 'cpn018', '');
      miFileNew.Caption := MyIni.ReadString('mynotex', 'cpn019', '');
      tbFileNew.Hint := miFile.Caption + ' - ' + miFileNew.Caption;
      miFileOpen.Caption := MyIni.ReadString('mynotex', 'cpn020', '');
      tbFileOpen.Hint := miFile.Caption + ' - ' + miFileOpen.Caption;
      miFileClose.Caption := MyIni.ReadString('mynotex', 'cpn021', '');
      miFileSave.Caption := MyIni.ReadString('mynotex', 'cpn022', '');
      tbFileSave.Hint := miFile.Caption + ' - ' + miFileSave.Caption;
      miFileUndo.Caption := MyIni.ReadString('mynotex', 'cpn023', '');
      miFileUpdate.Caption := MyIni.ReadString('mynotex', 'cpn024', '');
      miFileCopyAs.Caption := MyIni.ReadString('mynotex', 'cpn025', '');
      miFileImport.Caption := MyIni.ReadString('mynotex', 'cpn026', '');
      miFileExport.Caption := MyIni.ReadString('mynotex', 'cpn027', '');
      // miFileHTML and miFileConvert are below because added later
      miFileExit.Caption := MyIni.ReadString('mynotex', 'cpn028', '');
      miTrayExit.Caption := MyIni.ReadString('mynotex', 'cpn028', '');
      miSubject.Caption := MyIni.ReadString('mynotex', 'cpn029', '');
      miSubjectNew.Caption := MyIni.ReadString('mynotex', 'cpn030', '');
      tbSubjectNew.Hint := miSubject.Caption + ' - ' + miSubjectNew.Caption;
      miSubjectDelete.Caption := MyIni.ReadString('mynotex', 'cpn031', '');
      pmSubNew.Caption := MyIni.ReadString('mynotex', 'cpn030', '');
      pmSubDelete.Caption := MyIni.ReadString('mynotex', 'cpn031', '');
      miNotes.Caption := MyIni.ReadString('mynotex', 'cpn032', '');
      miNotesNew.Caption := MyIni.ReadString('mynotex', 'cpn033', '');
      tbNotesNew.Hint := miNotes.Caption + ' - ' + miNotesNew.Caption;
      miNotesDelete.Caption := MyIni.ReadString('mynotex', 'cpn034', '');
      miNotesUndo.Caption := MyIni.ReadString('mynotex', 'cpn023', '');
      pmNotesNew.Caption := MyIni.ReadString('mynotex', 'cpn033', '');
      pmNotesDelete.Caption := MyIni.ReadString('mynotex', 'cpn034', '');
      miNotesOrderBy.Caption := MyIni.ReadString('mynotex', 'cpn035', '');
      miOrderByDate.Caption := MyIni.ReadString('mynotex', 'cpn036', '');
      miOrderByTitle.Caption := MyIni.ReadString('mynotex', 'cpn037', '');
      miNotesFind.Caption := MyIni.ReadString('mynotex', 'cpn038', '');
      // miNotesMove is below because added later
      miNotesShowOnlyText.Caption := MyIni.ReadString('mynotex', 'cpn040', '');
      // Item deleted from version 1.3 of MyNotex
      // miOptionNotesFont.Caption := MyIni.ReadString('mynotex','cpn041','');
      // Item deleted from version 1.1 of MyNotex
      // miOptionNotesColorFont.Caption := MyIni.ReadString('mynotex','cpn042','');
      miTools.Caption := MyIni.ReadString('mynotex', 'cpn044', '');
      miToolsSyncDo.Caption := MyIni.ReadString('mynotex', 'cpn045', '');
      tbToolsSyncDo.Hint := miTools.Caption + ' - ' + miToolsSyncDo.Caption;
      // Item deleted from version 1.1 of MyNotex
      // miToolsSyncFolder.Caption := MyIni.ReadString('mynotex','cpn046','');
      miToolsCompact.Caption := MyIni.ReadString('mynotex', 'cpn047', '');
      miCopyright.Caption := MyIni.ReadString('mynotex', 'cpn048', '');
      miLicence.Caption := MyIni.ReadString('mynotex', 'cpn049', '');
      // Item deleted from version 1.1 of MyNotex
      // miToolsFormColor.Caption := MyIni.ReadString('mynotex','cpn050','');
      miFileHTML.Caption := MyIni.ReadString('mynotex', 'cpn051', '');
      miFileZim.Caption := MyIni.ReadString('mynotex', 'cpn151', '');
      miNotesMove.Caption := MyIni.ReadString('mynotex', 'cpn052', '');
      miSubjectComments.Caption := MyIni.ReadString('mynotex', 'cpn053', '');
      pmSubComments.Caption := MyIni.ReadString('mynotex', 'cpn053', '');
      miNotesInsert.Caption := MyIni.ReadString('mynotex', 'cpn056', '');
      miNotesAttach.Caption := MyIni.ReadString('mynotex', 'cpn055', '');
      miAttachNew.Caption := MyIni.ReadString('mynotex', 'cpn102', '');
      miAttachOpen.Caption := MyIni.ReadString('mynotex', 'cpn020', '');
      miAttachSaveAs.Caption := MyIni.ReadString('mynotex', 'cpn058', '');
      miAttachDelete.Caption := MyIni.ReadString('mynotex', 'cpn034', '');// Notes delete
      pmAttNew.Caption := MyIni.ReadString('mynotex', 'cpn102', '');
      pmAttOpen.Caption := MyIni.ReadString('mynotex', 'cpn020', '');
      pmAttSaveAs.Caption := MyIni.ReadString('mynotex', 'cpn058', '');
      pmAttDelete.Caption := MyIni.ReadString('mynotex', 'cpn034', '');
      pmChangeFontColor.Caption := MyIni.ReadString('mynotex', 'cpn059', '');
      pmChangeFontBackColor.Caption := MyIni.ReadString('mynotex', 'cpn059', '');
      pmTextCopyHtml.Caption := MyIni.ReadString('mynotex', 'cpn060', '');
      pmTextSelectAll.Caption := MyIni.ReadString('mynotex', 'cpn061', '');
      miToolsLanguage.Caption := MyIni.ReadString('mynotex', 'cpn062', '');
      tbKindFontChange.Hint := MyIni.ReadString('mynotex', 'cpn063', '');
      tbColorFontChange.Hint := MyIni.ReadString('mynotex', 'cpn064', '');
      tbFontBold.Hint := MyIni.ReadString('mynotex', 'cpn065', '');
      tbFontItalic.Hint := MyIni.ReadString('mynotex', 'cpn066', '');
      tbFontUnderline.Hint := MyIni.ReadString('mynotex', 'cpn067', '');
      tbFontStrike.Hint := MyIni.ReadString('mynotex', 'cpn068', '');
      tbFontRestore.Hint := MyIni.ReadString('mynotex', 'cpn069', '');
      pmChangeFontKind.Caption := MyIni.ReadString('mynotex', 'cpn070', '');
      tbOpenNote.Hint := MyIni.ReadString('mynotex', 'cpn071', '');
      miNotesEncDecrypt.Caption := MyIni.ReadString('mynotex', 'cpn073', '');
      miNotesSendToWp.Caption := MyIni.ReadString('mynotex', 'cpn074', '');
      lbPwdTop.Caption := MyIni.ReadString('mynotex', 'cpn075', '');
      lbPwdBottom.Caption := MyIni.ReadString('mynotex', 'cpn076', '');
      lbListTags.Caption := MyIni.ReadString('mynotex', 'cpn077', '');
      miToolsOptions.Caption := MyIni.ReadString('mynotex', 'cpn039', '');
      pmHeading1.Caption := MyIni.ReadString('mynotex', 'cpn079', '');
      pmHeading2.Caption := MyIni.ReadString('mynotex', 'cpn080', '');
      pmHeading3.Caption := MyIni.ReadString('mynotex', 'cpn081', '');
      pmRestoreFont.Caption := MyIni.ReadString('mynotex', 'cpn082', '');
      miNotesTags.Caption := MyIni.ReadString('mynotex', 'cpn083', '');
      miTagsRename.Caption := MyIni.ReadString('mynotex', 'cpn084', '');
      miTagsRemove.Caption := MyIni.ReadString('mynotex', 'cpn085', '');
      miFileConvert.Caption := MyIni.ReadString('mynotex', 'cpn086', '');
      tbBackColorFontChange.Hint := MyIni.ReadString('mynotex', 'cpn087', '');
      tbAlignLeft.Hint := MyIni.ReadString('mynotex', 'cpn088', '');
      tbAlignCenter.Hint := MyIni.ReadString('mynotex', 'cpn089', '');
      tbAlignRight.Hint := MyIni.ReadString('mynotex', 'cpn090', '');
      tbAlignFill.Hint := MyIni.ReadString('mynotex', 'cpn091', '');
      tbAlignIndent.Hint := MyIni.ReadString('mynotex', 'cpn092', '');
      // Item deleted from version 1.1 of MyNotex
      // miToolsAutoSync.Caption := MyIni.ReadString('mynotex','cpn093','');
      // miToolsFormTransp.Caption := MyIni.ReadString('mynotex','cpn094','');
      tbFind.Hint := MyIni.ReadString('mynotex', 'cpn097', '');
      pmTextCut.Caption := MyIni.ReadString('mynotex', 'cpn099', '');
      tbCut.Hint := MyIni.ReadString('mynotex', 'cpn099', '');
      pmTextCopy.Caption := MyIni.ReadString('mynotex', 'cpn100', '');
      tbCopy.Hint := MyIni.ReadString('mynotex', 'cpn100', '');
      tbCopyHtml.Hint := MyIni.ReadString('mynotex', 'cpn060', '');
      pmTextPaste.Caption := MyIni.ReadString('mynotex', 'cpn101', '');
      tbPaste.Hint := MyIni.ReadString('mynotex', 'cpn101', '');
      miNotesImages.Caption := MyIni.ReadString('mynotex', 'cpn103', '');
      miHelp.Caption := MyIni.ReadString('mynotex', 'cpn104', '');
      miNotesSendToBrowser.Caption := MyIni.ReadString('mynotex', 'cpn105', '');
      pmOpenBrowser.Caption := MyIni.ReadString('mynotex', 'cpn105', '');
      pmTextCopyLatex.Caption := MyIni.ReadString('mynotex', 'cpn106', '');
      pmTextSendAsEmail.Caption := MyIni.ReadString('mynotex', 'cpn107', '');
      miNotesPrint.Caption := MyIni.ReadString('mynotex', 'cpn108', '');
      miNotesShowActivities.Caption := MyIni.ReadString('mynotex', 'cpn133', '');
      miToolsEncryptGPG.Caption := MyIni.ReadString('mynotex', 'cpn109', '');
      miToolsDecryptGPG.Caption := MyIni.ReadString('mynotex', 'cpn110', '');
      tbActIndLeft.Hint := MyIni.ReadString('mynotex', 'cpn111', '');
      tbActIndRight.Hint := MyIni.ReadString('mynotex', 'cpn112', '');
      tbActMoveUp.Hint := MyIni.ReadString('mynotex', 'cpn113', '');
      tbActMoveDown.Hint := MyIni.ReadString('mynotex', 'cpn114', '');
      tbActInsertRow.Hint := MyIni.ReadString('mynotex', 'cpn115', '');
      tbActDeleteRow.Hint := MyIni.ReadString('mynotex', 'cpn116', '');
      tbActShowWbs.Hint := MyIni.ReadString('mynotex', 'cpn117', '');
      tbActCopyGroup.Hint := MyIni.ReadString('mynotex', 'cpn118', '');
      tbActPasteGroup.Hint := MyIni.ReadString('mynotex', 'cpn119', '');
      tbActCopyAll.Hint := MyIni.ReadString('mynotex', 'cpn120', '');
      tbActMoveFromText.Hint := MyIni.ReadString('mynotex', 'cpn135', '');
      miSubjectLook.Caption := MyIni.ReadString('mynotex', 'cpn136', '');
      pmSubLook.Caption := MyIni.ReadString('mynotex', 'cpn136', '');
      miSubjectOrder.Caption := MyIni.ReadString('mynotex', 'cpn035', '');
      miSubjectOrderTitle.Caption := MyIni.ReadString('mynotex', 'cpn037', '');
      miNotesLook.Caption := MyIni.ReadString('mynotex', 'cpn136', '');
      pmNotesLook.Caption := MyIni.ReadString('mynotex', 'cpn136', '');
      miNotesShowCal.Caption := MyIni.ReadString('mynotex', 'cpn137', '');
      tbActClearAll.Hint := MyIni.ReadString('mynotex', 'cpn138', '');
      miSubjectOrderCustom.Caption := MyIni.ReadString('mynotex', 'cpn139', '');
      miOrderCustom.Caption := MyIni.ReadString('mynotex', 'cpn139', '');
      miFilePrinterSetup.Caption := MyIni.ReadString('mynotex', 'cpn140', '');
      lbCalAct.Caption := MyIni.ReadString('mynotex', 'cpn141', '');
      tbActMoveBack.Hint := MyIni.ReadString('mynotex', 'cpn142', '');
      tbActMoveFor.Hint := MyIni.ReadString('mynotex', 'cpn143', '');
      miToolsAlarm.Caption := MyIni.ReadString('mynotex', 'cpn144', '');
      bnSubUp.Hint := MyIni.ReadString('mynotex', 'cpn113', '');
      bnNotesUp.Hint := MyIni.ReadString('mynotex', 'cpn113', '');
      bnSubDown.Hint := MyIni.ReadString('mynotex', 'cpn114', '');
      bnNotesDown.Hint := MyIni.ReadString('mynotex', 'cpn114', '');
      bnFind.Hint := MyIni.ReadString('mynotex', 'cpn149', '');
      bnFind2.Hint := MyIni.ReadString('mynotex', 'cpn150', '');
      miFileZim.Caption := MyIni.ReadString('mynotex', 'cpn151', '');
      // Hints in dbNavigator
      if MyIni.ReadString('mynotex', 'cpn145', '') <> '' then
      begin
        dbNavigator.Hints.Clear;
        dbNavigator.Hints.Add(MyIni.ReadString('mynotex', 'cpn145', ''));
        dbNavigator.Hints.Add(MyIni.ReadString('mynotex', 'cpn146', ''));
        dbNavigator.Hints.Add(MyIni.ReadString('mynotex', 'cpn147', ''));
        dbNavigator.Hints.Add(MyIni.ReadString('mynotex', 'cpn148', ''));
      end;
      // Label in activity grid
      lbID := MyIni.ReadString('mynotex', 'cpn122', '');
      lbWbs := MyIni.ReadString('mynotex', 'cpn123', '');
      lbState := MyIni.ReadString('mynotex', 'cpn124', '');
      lbActivity := MyIni.ReadString('mynotex', 'cpn125', '');
      lbStartDate := MyIni.ReadString('mynotex', 'cpn126', '');
      lbEndDate := MyIni.ReadString('mynotex', 'cpn127', '');
      lbDuration := MyIni.ReadString('mynotex', 'cpn128', '');
      lbResources := MyIni.ReadString('mynotex', 'cpn129', '');
      lbPriority := MyIni.ReadString('mynotex', 'cpn130', '');
      lbCompletion := MyIni.ReadString('mynotex', 'cpn131', '');
      lbCost := MyIni.ReadString('mynotex', 'cpn132', '');
      lbNotes := MyIni.ReadString('mynotex', 'cpn134', '');
      grActGrid.Repaint;
      // Months and days names
      with FDate do
      begin
        LongMonthNames[1] := MyIni.ReadString('mynotex', 'cmn001', '');
        LongMonthNames[2] := MyIni.ReadString('mynotex', 'cmn002', '');
        LongMonthNames[3] := MyIni.ReadString('mynotex', 'cmn003', '');
        LongMonthNames[4] := MyIni.ReadString('mynotex', 'cmn004', '');
        LongMonthNames[5] := MyIni.ReadString('mynotex', 'cmn005', '');
        LongMonthNames[6] := MyIni.ReadString('mynotex', 'cmn006', '');
        LongMonthNames[7] := MyIni.ReadString('mynotex', 'cmn007', '');
        LongMonthNames[8] := MyIni.ReadString('mynotex', 'cmn008', '');
        LongMonthNames[9] := MyIni.ReadString('mynotex', 'cmn009', '');
        LongMonthNames[10] := MyIni.ReadString('mynotex', 'cmn010', '');
        LongMonthNames[11] := MyIni.ReadString('mynotex', 'cmn011', '');
        LongMonthNames[12] := MyIni.ReadString('mynotex', 'cmn012', '');
        LongDayNames[1] := MyIni.ReadString('mynotex', 'cdn001', '');
        LongDayNames[2] := MyIni.ReadString('mynotex', 'cdn002', '');
        LongDayNames[3] := MyIni.ReadString('mynotex', 'cdn003', '');
        LongDayNames[4] := MyIni.ReadString('mynotex', 'cdn004', '');
        LongDayNames[5] := MyIni.ReadString('mynotex', 'cdn005', '');
        LongDayNames[6] := MyIni.ReadString('mynotex', 'cdn006', '');
        LongDayNames[7] := MyIni.ReadString('mynotex', 'cdn007', '');
      end;
      // Date notes format
      if MyIni.ReadString('mynotex', 'dtfmt', '') = 'DMY' then
      begin
        FDate.ShortDateFormat := 'dd-mm-yyyy';
        FDate.LongDateFormat := 'dddd d mmmm yyyy';
        dbDate.DateDisplayOrder := ddoDMY;
        dtCalAct.DateDisplayOrder := ddoDMY;
        DateOrder := 'DMY';
        DateMask := '  -  -    ';
      end
      else if MyIni.ReadString('mynotex', 'dtfmt', '') = 'YMD' then
      begin
        FDate.ShortDateFormat := 'yyyy-mm-dd';
        FDate.LongDateFormat := 'dddd mmmm d yyyy';
        dbDate.DateDisplayOrder := ddoYMD;
        dtCalAct.DateDisplayOrder := ddoYMD;
        DateOrder := 'YMD';
        DateMask := '    -  -  ';
      end
      else
      begin
        FDate.ShortDateFormat := 'mm-dd-yyyy';
        FDate.LongDateFormat := 'dddd mmmm d yyyy';
        dbDate.DateDisplayOrder := ddoMDY;
        dtCalAct.DateDisplayOrder := ddoMDY;
        DateOrder := 'MDY';
        DateMask := '  -  -    ';
      end;
      if MyIni.ReadString('mynotex', 'hour', '') = '24' then
      begin
        fl24Hour := True;
      end;
      // To refresh the date component if tables are opened;
      if sqNotes.Active = True then
        sqNotes.Refresh;
    finally
      MyIni.Free;
    end;
  end
  else
  begin
    // No translation file
    // Messages dialogs
    msg001 := 'Delete current subject and related notes?';
    msg002 := 'Create new subject?';
    msg003 := 'There are no subjects in the current file.';
    msg004 := 'Cancel changes?';
    msg005 := 'File in use cannot be selected.';
    msg006 := 'File for exportation is empty.';
    msg007 := 'Cannot open selected file; check it was created with MyNotex.';
    msg008 := 'Selected file is not available.';
    msg009 := 'Delete current note?';
    msg010 := 'Changes in current file:';
    msg011 := 'deleted elements:';
    msg012 := 'added or changed elements:';
    msg013 := 'Changes in external file:';
    msg014 := 'Synchronization was not successful; original file is available as';
    msg015 := 'Compact current file?';
    msg016 := 'Compact was successful; anyway, the original file is available as';
    msg017 := 'Compact was not successful; the original file is available as';
    msg018 := 'Operation not correct.';
    msg019 := 'Existing files cannot be overwritten.';
    msg020 := 'It was not possible to create the file.';
    msg021 := 'Search text not specified.';
    msg022 := 'Search text not found.';
    msg023 := 'No subjects selected.';
    msg024 := 'Operation successful.';
    msg025 := 'Cannot create HTML file.';
    // msg026 := 'Confirm file overwriting?';
    msg027 := 'It was not possible to move the note.';
    msg028 := 'It was not possible to load the file.';
    msg029 := 'It was not possible to create the attachment directory.';
    msg030 := 'Attachment directory is not available.';
    msg031 := 'No attachment is selected.';
    msg032 := 'It was not possible to open the file.';
    msg033 := 'This file is already attached to the current note:';
    msg034 := 'Delete selected attachment?';
    msg035 := 'It was not possible to delete the file.';
    msg036 := 'is a subject already present in the file in use.';
    msg037 := 'It was not possible to copy the attachments.';
    msg038 := 'Too many tags to create the list.';
    msg039 := 'The text of the note has been encrypted, so it is not available.';
    msg040 := 'The passwords do not match.';
    msg041 := 'Clear encryption of the text of the current note?';
    msg042 := 'The text must be visible to clear encryption.';
    msg043 := 'There is no other subject to move the note to.';
    msg044 := 'Rename a tag in the archive in use?';
    msg045 := 'Delete a tag in the archive in use?';
    msg046 := 'Modify tags.';
    msg047 := 'Insert the old name.';
    msg048 := 'Insert the new name.';
    msg049 := 'Confirm data conversion from';
    msg050 := 'Notes without subject';
    msg051 := 'Notes created:';
    msg052 := 'Notes modified:';
    msg053 := 'Indentation';
    msg054 := 'Insert a number between 0 and 250.';
    msg055 := 'Indentation not correct.';
    msg056 := 'Transparency';
    msg057 := 'Insert a number between 0 (full transparency) and 150 (no transparency).';
    msg058 := 'Level of transparency not correct.';
    msg059 := 'Date not correct.';
    msg060 := 'File to import is empty.';
    msg061 := 'Dates not correct.';
    msg062 := 'Print the current note?';
    msg063 := 'Existing encrypted file cannot be overwritten.';
    msg064 := 'Recipient';
    msg065 := 'Type the recipient email or ID.';
    msg066 := 'Password';
    // Not usd from the 1.4 version
    // msg067 := 'Type the password.';
    msg068 := 'It was not possible to encrypt the file.';
    msg069 := 'Existing decrypted file cannot be overwritten.';
    msg070 := 'It was not possible to decrypt the file.';
    msg071 := 'It is not possible to print a note with pictures within it.';
    msg072 := 'It was not possible to print the note.';
    msg073 := 'Delete all activities of the current note?';
    msg074 := 'Are you really sure to delete all activities?';
    msg075 := 'Start date cannot be later than end date.';
    msg076 := 'End date cannot be sooner than start date.';
    msg077 := 'Delete current activity and its possible subactivities?';
    msg078 := 'Move the activities in the note to the grid?';
    msg079 := 'Characters';
    msg080 := 'The date * is not in the right format.';
    msg081 := 'It''s';
    msg082 := 'left';
    msg083 := 'Export the current file as a Zim archive?';
    msg084 := 'It was not possibile to export the file as a Zim archive.';
    // Status bar messages
    sbr001 := 'Editing note.';
    sbr002 := 'No notes.';
    sbr003 := 'Note n.';
    sbr004 := 'of';
    sbr005 := 'Synchronization running...';
    sbr006 := 'Synchronization completed.';
    sbr007 := 'Synchronization not successful.';
    sbr008 := 'No file opened.';
    sbr009 := 'Last modified on';
    sbr010 := 'at';
    sbr011 := 'Set bookmark n.';
    // Various captions and labels modified in the code
    cpt001 := 'Create new file';
    cpt002 := 'Copy current file';
    cpt003 := 'Importation';
    cpt004 := 'Subjects to import';
    cpt005 := 'Import';
    cpt006 := 'Delete imported data from original file.';
    cpt007 := 'Exportation';
    cpt008 := 'Subjects to export';
    cpt009 := 'Export';
    cpt010 := 'Delete exported data from original file.';
    cpt011 := 'Search not active.';
    cpt012 := 'Subject';
    cpt013 := 'Note title';
    cpt014 := 'Note date';
    cpt015 := 'Occurrences:';
    cpt016 := 'Move note';
    cpt017 := 'Move the current note under the selected subject.';
    cpt018 := 'Cancel';
    cpt019 := 'Subject comments';
    cpt020 := 'Subject title';
    cpt021 := 'Comments';
    cpt022 := 'Open a file of MyNotex';
    cpt023 := 'File of MyNotex|*.mnt';
    cpt024 := 'Open a document';
    cpt025 := 'File of Writer|*.odt|Text file|*.*';
    cpt026 := 'Open a file';
    cpt027 := 'All files|*.*';
    cpt028 := 'Open a language file of MyNotex';
    cpt029 := 'Language file of MyNotex|*.lng';
    cpt030 := 'Encrypt';
    cpt031 := 'Insert password in the fields below';
    cpt032 := 'Show characters';
    cpt033 := 'Author and copyright:';
    cpt034 := 'Visit the web site of MyNotex.';
    cpt035 := 'File HTML|*.htm, *.html';
    cpt036 := 'Select and deselect all';
    cpt037 := 'Open picture';
    cpt038 := 'Resize picture';
    cpt039 := 'Picture size:';
    cpt040 := 'Visit the support forum of MyNotex.';
    cpt041 := 'Export dates';
    cpt042 := 'Open encrypted file';
    cpt043 := 'Encrypted file|*.pgp';
    cpt044 := 'Default font of the text:';
    cpt045 := 'pt.';
    cpt046 := 'Sync folder:';
    cpt047 := 'Forms color:';
    cpt048 := 'Activate tray icon';
    cpt049 := 'Open last file on start';
    cpt050 := 'Activate autosync';
    cpt051 := 'Change font';
    cpt052 := 'Change folder';
    cpt053 := 'Change color';
    cpt054 := 'Select a folder';
    cpt055 := 'Forms transparency';
    cpt056 := 'Look';
    cpt057 := 'Background color';
    cpt058 := 'Font color';
    cpt059 := 'Default color 1';
    cpt060 := 'Default color 2';
    cpt061 := 'Default color 3';
    cpt062 := 'Remove color';
    cpt063 := 'Default 1';
    cpt064 := 'Default 2';
    cpt065 := 'Default 3';
    cpt066 := 'Diary';
    cpt067 := 'Selected activities';
    cpt068 := 'Start date greater than or equal to';
    cpt069 := 'End date less than or equal to';
    cpt070 := 'Filter on resources';
    cpt071 := 'A month';
    cpt072 := 'Today';
    cpt073 := 'Export';
    cpt074 := 'No sync message except on error';
    cpt075 := 'No characters count';
    cpt076 := 'Titles font and background colors';
    cpt077 := 'Font color 1';
    cpt078 := 'Background color 1';
    cpt079 := 'Font color 2';
    cpt080 := 'Background color 2';
    cpt081 := 'Font color 3';
    cpt082 := 'Background color 3';
    cpt083 := 'Options';
    cpt084 := 'All';
    cpt085 := 'No resource';
    cpt086 := 'No date';
    cpt087 := 'Resources';
    cpt088 := 'Cost';
    cpt089 := 'Activities of all notes';
    cpt090 := 'Activities of the current note';
    cpt091 := 'The project begins on';
    cpt092 := 'and ends on';
    cpt093 := 'The project ends on';
    cpt094 := 'No dates';
    cpt095 := 'Disable autosave';
    cpt096 := 'Line space';
    cpt097 := 'Paragraph space';
    cpt098 := 'Set alarm';
    cpt099 := 'Activate';

    // sdSaveDialog filter
    OpenSaveDlgFilter := 'File of MyNotex';
  end;
end;


initialization
  {$I unit1.lrs}

finalization
  fmMain.ActivityGroup.Free;

end.
