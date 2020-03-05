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

unit Unit9;

{$mode objfpc}

interface

uses
  Classes, SysUtils, FileUtil, RTTICtrls, LResources, Forms, Controls,
  Graphics, Dialogs, Sqlite3DS, DB, BufDataset, LCLProc, StdCtrls,
  Grids, Calendar, ExtCtrls, types;

type

  { TfmCalendar }

  TfmCalendar = class(TForm)
    Bevel1: TBevel;
    bnMonthDiary: TButton;
    bvLine: TBevel;
    bfCal: TBufDataset;
    bnTodayDiary: TButton;
    bnClose: TButton;
    bnExportDiary: TButton;
    clCalEnd: TCalendar;
    clCalStart: TCalendar;
    dsCal: TDataSource;
    grCalGrid: TStringGrid;
    lbSelAct: TLabel;
    lbResources: TLabel;
    lbStartDate: TLabel;
    lbEndDate: TLabel;
    lbxResources: TListBox;
    pnCal: TPanel;
    rgCalRange: TRadioGroup;
    procedure bnCloseClick(Sender: TObject);
    procedure bnExportDiaryClick(Sender: TObject);
    procedure bnMonthDiaryClick(Sender: TObject);
    procedure bnTodayDiaryClick(Sender: TObject);
    procedure clCalEndChange(Sender: TObject);
    procedure clCalStartClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
    procedure grCalGridDrawCell(Sender: TObject; aCol, aRow: integer;
      aRect: TRect; aState: TGridDrawState);
    procedure grCalGridMouseWheelDown(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: boolean);
    procedure grCalGridMouseWheelUp(Sender: TObject; Shift: TShiftState;
      MousePos: TPoint; var Handled: boolean);
    procedure lbxResourcesClick(Sender: TObject);
    procedure rgCalRangeClick(Sender: TObject);
  private
    function CleanString(stInput: string): string;
    procedure CreateCal(flGetResList: boolean; flExp: boolean);
    function GetDate(stDate, stDtFormat: string): TDate;
    { private declarations }
  public
    stAll: string;
    stNoRes: string;
    stNoDate: string;
    stResources: string;
    stCost: string;
    { public declarations }
  end;

var
  fmCalendar: TfmCalendar;

implementation

uses Unit1;

{ TfmCalendar }

procedure TfmCalendar.FormCreate(Sender: TObject);
begin
  //Change form color
  fmCalendar.Color := fmMain.Color;
  with grCalGrid do
  begin
    Clear;
    FocusRectVisible := False;
    Columns.Add;
    Columns.Add;
    ColWidths[0] := 0;
    ColWidths[1] := 3000;
    Columns.Items[0].Visible := False;
    RowCount := 0;
  end;
  clCalStart.DateTime := IncMonth(Now, -12);
  clCalEnd.DateTime := IncMonth(Now, 1);
  // Do not change it, since another color could make invisible
  // the selected item when the component is not focused
  lbxResources.Color := clBtnFace;
  lbSelAct.Caption := fmMain.cpt067 + ': 0.';
end;

procedure TfmCalendar.FormShow(Sender: TObject);
begin
  // Create calendar
  CreateCal(True, False);
end;

procedure TfmCalendar.grCalGridDrawCell(Sender: TObject; aCol, aRow: integer;
  aRect: TRect; aState: TGridDrawState);
begin
  // Format output in the grid
  if grCalGrid.Cells[0, aRow] <> '' then
  begin
    grCalGrid.Canvas.Brush.Color := grCalGrid.FixedColor;
    grCalGrid.Canvas.Font.Style := [fsBold];
    if Utf8Copy(grCalGrid.Cells[0, aRow], 1, 1) = '-' then
      grCalGrid.Canvas.Font.Color := clRed;
    grCalGrid.Canvas.FillRect(aRect);
    grCalGrid.Canvas.TextOut(aRect.Left + 2,
      aRect.Top + 3, grCalGrid.Cells[aCol, aRow]);
  end
  else
  begin
    grCalGrid.Canvas.FillRect(aRect);
    grCalGrid.Canvas.TextOut(aRect.Left + 30,
      aRect.Top + 3, grCalGrid.Cells[aCol, aRow]);
  end;
end;

procedure TfmCalendar.grCalGridMouseWheelDown(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
begin
  // Scroll down faster
  grCalGrid.Row := grCalGrid.Row + grCalGrid.VisibleRowCount div 2;
end;

procedure TfmCalendar.grCalGridMouseWheelUp(Sender: TObject;
  Shift: TShiftState; MousePos: TPoint; var Handled: boolean);
begin
  // Scroll up faster
  grCalGrid.Row := grCalGrid.Row - grCalGrid.VisibleRowCount div 2;
end;

procedure TfmCalendar.lbxResourcesClick(Sender: TObject);
begin
  // Filter resources
  CreateCal(False, False);
end;

procedure TfmCalendar.rgCalRangeClick(Sender: TObject);
begin
  // Create calendar
  CreateCal(True, False);
end;

procedure TfmCalendar.bnTodayDiaryClick(Sender: TObject);
begin
  // Today
  clCalEnd.DateTime := Now;
  CreateCal(True, False);
end;

procedure TfmCalendar.bnMonthDiaryClick(Sender: TObject);
begin
  // One month
  clCalEnd.DateTime := IncMonth(Now, 1);
  CreateCal(True, False);
end;

procedure TfmCalendar.bnExportDiaryClick(Sender: TObject);
begin
  // Export data
  CreateCal(False, True);
end;

procedure TfmCalendar.clCalEndChange(Sender: TObject);
begin
  // Constraints on dates
  // Valid also as clCalStartChange
  if clCalEnd.DateTime < clCalStart.DateTime then
  begin
    clCalEnd.DateTime := clCalStart.DateTime;
  end;
end;

procedure TfmCalendar.clCalStartClick(Sender: TObject);
begin
  // Change calendars
  CreateCal(True, False);
end;

procedure TfmCalendar.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  // Clear grid
  grCalGrid.Clear;
  lbxResources.Clear;
end;

procedure TfmCalendar.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
    fmCalendar.ModalResult := mrCancel;
end;


procedure TfmCalendar.bnCloseClick(Sender: TObject);
begin
  // Close
  Close;
end;

procedure TfmCalendar.CreateCal(flGetResList: boolean; flExp: boolean);
var
  sqCal: TSqlite3Dataset;
  myGUID: TGUID;
  flCalendar: TextFile;
  slResList, slResList2: TStringList;
  myData: WideString;
  stFileName, stDtFormat: string;
  i, n, iTot, iSeq: integer;
begin
  // Create calendar
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;
  bfCal.Close;
  bfCal.FileName := GetTempDir + DirectorySeparator + '~cal';
  bfCal.CreateDataset;
  bfCal.Open;
  sqCal := TSqlite3Dataset.Create(Self);
  if flGetResList = True then
  begin
    lbxResources.Clear;
  end;
  // Get data
  with sqCal do
    try
      FileName := fmMain.sqNotes.FileName;
      TableName := 'Notes';
      PrimaryKey := 'IDNotes';
      if rgCalRange.ItemIndex = 0 then
        SQL := 'Select IDNotes, NotesTitle, NotesDateFormat, NotesActivities from Notes '
          + 'where NotesActivities <> ' + #39#39 + ' order by IDNotes'
      else
        SQL := 'Select IDNotes, NotesTitle, NotesDateFormat, NotesActivities from Notes '
          + 'where NotesActivities <> ' + #39#39 + ' and IDNotes = ' +
          fmMain.sqNotes.FieldByName('IDNotes').AsString + ' order by IDNotes';
      Open;
      slResList := TStringList.Create;
      while not EOF do
      begin
        stDtFormat := FieldByName('NotesDateFormat').AsString;
        myData := FieldByName('NotesActivities').AsString;
        while UTF8Length(myData) > 1 do
        begin
          bfCal.Append;
          bfCal.FieldByName('mdCalNoteID').AsInteger :=
            FieldByName('IDNotes').AsInteger;
          bfCal.FieldByName('mdCalNoteTitle').AsString :=
            FieldByName('NotesTitle').AsString;
          bfCal.FieldByName('mdCalCode').AsString :=
            UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1);
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          // Skip WBS
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          bfCal.FieldByName('mdCalState').AsString :=
            UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1);
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          bfCal.FieldByName('mdCalName').AsString :=
            UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1);
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalStartDate').AsDateTime :=
              GetDate(UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1), stDtFormat);
          if bfCal.FieldByName('mdCalStartDate').AsDateTime = 0 then
            bfCal.FieldByName('mdCalStartDate').Clear;
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalEndDate').AsDateTime :=
              GetDate(UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1), stDtFormat);
          if bfCal.FieldByName('mdCalEndDate').AsDateTime = 0 then
            bfCal.FieldByName('mdCalEndDate').Clear;
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          // Skip Duration
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalResources').AsString :=
              UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1)
          else
            bfCal.FieldByName('mdCalResources').AsString := stNoRes;
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalPriority').AsInteger :=
              StrToInt(UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1))
          else
            bfCal.FieldByName('mdCalPriority').AsInteger := 0;
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalCompletion').AsInteger :=
              StrToInt(UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1))
          else
            bfCal.FieldByName('mdCalCompletion').AsInteger := 0;
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalCost').AsString :=
              UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1)
          else
            bfCal.FieldByName('mdCalCost').AsString := '0';
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          if UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1) <> '' then
            bfCal.FieldByName('mdCalNotes').AsString :=
              UTF8Copy(myData, 1, UTF8Pos(#9, myData) - 1);
          myData := UTF8Copy(myData, UTF8Pos(#9, myData) + 1, UTF8Length(myData));
          myData := UTF8Copy(myData, UTF8Pos(LineEnding, myData) +
            1, UTF8Length(myData));
          bfCal.Post;
          // Remove record if the selected resource is not present
          if flGetResList = False then
          begin
            if lbxResources.ItemIndex > 0 then
            begin
              slResList.Clear;
              slResList.Text :=
                StringReplace(bfCal.FieldByName('mdCalResources').AsString,
                ', ', LineEnding, [rfReplaceAll]);
              if slResList.IndexOf(lbxResources.Items[lbxResources.ItemIndex]) < 0 then
              begin
                bfCal.Delete;
              end;
            end;
          end;
        end;
        Next;
      end;
    finally
      sqCal.Free;
      slResList.Free;
      Screen.Cursor := crDefault;
    end;
  Screen.Cursor := crHourGlass;
  // Write calendar
  if flExp = False then
    try
      slResList2 := TStringList.Create;
      grCalGrid.Clear;
      bfCal.First;
      i := -1;
      iTot := 0;
      while not bfCal.EOF do
      begin
        if ((not (UpperCase(bfCal.FieldByName('mdCalState').AsString) = 'DONE')) and
          ((bfCal.FieldByName('mdCalStartDate').AsDateTime >= clCalStart.DateTime) or
          (bfCal.FieldByName('mdCalStartDate').AsString = '')) and
          ((bfCal.FieldByName('mdCalEndDate').AsDateTime <= clCalEnd.DateTime) or
          (bfCal.FieldByName('mdCalEndDate').AsString = '')) and
          (Copy(bfCal.FieldByName('mdCalCode').AsString, 1, 1) = '*')) then
        begin
          if flGetResList = True then
          begin
            if bfCal.FieldByName('mdCalResources').AsString <> '' then
            begin
              slResList2.Clear;
              slResList2.Text :=
                StringReplace(bfCal.FieldByName('mdCalResources').AsString,
                ', ', LineEnding, [rfReplaceAll]);
              for n := 0 to slResList2.Count - 1 do
              begin
                if lbxResources.Items.IndexOf(slResList2[n]) < 0 then
                  lbxResources.Items.Add(slResList2[n]);
              end;
            end;
          end;
          Inc(iTot);
          Inc(i);
          grCalGrid.RowCount := i + 1;
          // Expired activity
          if bfCal.FieldByName('mdCalEndDate').AsDateTime < Now then
            grCalGrid.Cells[0, i] := '-'
          else
            grCalGrid.Cells[0, i] := '+';
          grCalGrid.Cells[0, i] :=
            grCalGrid.Cells[0, i] + bfCal.FieldByName('mdCalNoteID').AsString;
          grCalGrid.Cells[1, i] :=
            bfCal.FieldByName('mdCalName').AsString + ' [' +
            bfCal.FieldByName('mdCalNoteTitle').AsString + ']';
          Inc(i);
          grCalGrid.RowCount := i + 1;
          if ((bfCal.FieldByName('mdCalStartDate').AsDateTime > 0) and
            (bfCal.FieldByName('mdCalEndDate').AsDateTime > 0)) then
          begin
            grCalGrid.Cells[1, i] :=
              FormatDateTime(FDate.LongDateFormat,
              bfCal.FieldByName('mdCalStartDate').AsDateTime) + ' - ' +
              FormatDateTime(FDate.LongDateFormat,
              bfCal.FieldByName('mdCalEndDate').AsDateTime) + ' | ' +
              bfCal.FieldByName('mdCalState').AsString;
          end
          else if bfCal.FieldByName('mdCalStartDate').AsDateTime > 0 then
          begin
            grCalGrid.Cells[1, i] :=
              FormatDateTime(FDate.LongDateFormat,
              bfCal.FieldByName('mdCalStartDate').AsDateTime) + ' - ' +
              stNoDate + ' | ' + bfCal.FieldByName('mdCalState').AsString;
          end
          else if bfCal.FieldByName('mdCalEndDate').AsDateTime > 0 then
          begin
            grCalGrid.Cells[1, i] :=
              stNoDate + ' - ' + FormatDateTime(FDate.LongDateFormat,
              bfCal.FieldByName('mdCalEndDate').AsDateTime) +
              ' | ' + bfCal.FieldByName('mdCalState').AsString;
          end
          else
          begin
            grCalGrid.Cells[1, i] :=
              stNoDate + ' | ' + bfCal.FieldByName('mdCalState').AsString;
          end;
          Inc(i);
          grCalGrid.RowCount := i + 1;
          grCalGrid.Cells[1, i] :=
            bfCal.FieldByName('mdCalResources').AsString + ' | ' +
            bfCal.FieldByName('mdCalPriority').AsString + ' | ' +
            bfCal.FieldByName('mdCalCompletion').AsString + '% | ' +
            bfCal.FieldByName('mdCalCost').AsString + ' ' + CurrencyString;
        end;
        bfCal.Next;
      end;
      if flGetResList = True then
      begin
        lbxResources.Sorted := True;
        lbxResources.Sorted := False;
        if lbxResources.Items.Count > 0 then
        begin
          lbxResources.Items.Insert(0, stAll);
          lbxResources.ItemIndex := 0;
        end;
      end;
      lbSelAct.Caption := fmMain.cpt067 + ': ' + IntToStr(iTot) + '.';
      if iTot > 0 then
        bnExportDiary.Enabled := True
      else
        bnExportDiary.Enabled := False;
    finally
      bfCal.Close;
      slResList2.Free;
      Screen.Cursor := crDefault;
    end
  else
    try
      // Create a new ical file
      fmMain.sdSaveDialog.Title := fmMain.cpt001;
      fmMain.sdSaveDialog.Filter := 'File iCal|*.ics';
      fmMain.sdSaveDialog.DefaultExt := '.ics';
      fmMain.sdSaveDialog.FileName := '';
      if fmMain.sdSaveDialog.Execute = True then
      begin
        stFileName := fmMain.sdSaveDialog.FileName;
      end;
      AssignFile(flCalendar, stFileName);
      Rewrite(flCalendar);
      if stFileName = '' then
        Abort;
      bfCal.First;
      iSeq := 0;
      while not bfCal.EOF do
      begin
        if ((not (UpperCase(bfCal.FieldByName('mdCalState').AsString) = 'DONE')) and
          ((bfCal.FieldByName('mdCalStartDate').AsDateTime >= clCalStart.DateTime) or
          (bfCal.FieldByName('mdCalStartDate').AsString = '')) and
          ((bfCal.FieldByName('mdCalEndDate').AsDateTime <= clCalEnd.DateTime) or
          (bfCal.FieldByName('mdCalEndDate').AsString = ''))) then
        begin
          WriteLn(flCalendar, 'BEGIN:VCALENDAR');
          WriteLn(flCalendar, 'PRODID:MyNotex Activity List');
          WriteLn(flCalendar, 'VERSION:2.0');
          WriteLn(flCalendar, 'METHOD:PUBLISH');
          WriteLn(flCalendar, 'BEGIN:VTODO');
          CreateGUID(myGUID);
          WriteLn(flCalendar, 'UID:' + FormatDateTime('YYYYMMDD', Now) +
            'T' + FormatDateTime('HHMMSS', Now) + 'Z' +
            Copy(GUIDToString(myGUID), 2, Length(GUIDToString(myGUID)) - 2));
          WriteLn(flCalendar, 'DTSTAMP:' + FormatDateTime('YYYYMMDD', Now) +
            'T' + FormatDateTime('HHMMSS', Now) + 'Z');
          WriteLn(flCalendar, 'SUMMARY:' + CleanString(
            bfCal.FieldByName('mdCalName').AsString));
          if bfCal.FieldByName('mdCalNotes').AsString <> '' then
            WriteLn(flCalendar, 'DESCRIPTION:' + stResources + ': ' +
              CleanString(bfCal.FieldByName('mdCalResources').AsString) +
              '. ' + stCost + ': ' + CleanString(
              bfCal.FieldByName('mdCalCost').AsString) + ' ' +
              CurrencyString + '.' + CleanString(LineEnding +
              LineEnding + bfCal.FieldByName('mdCalNotes').AsString))
          else
            WriteLn(flCalendar, 'DESCRIPTION:' + stResources + ': ' +
              CleanString(bfCal.FieldByName('mdCalResources').AsString) +
              '. ' + stCost + ': ' + CleanString(
              bfCal.FieldByName('mdCalCost').AsString) + ' ' +
              CurrencyString + '.');
          if bfCal.FieldByName('mdCalEndDate').AsString <> '' then
            WriteLn(flCalendar, 'DUE;VALUE=DATE:' +
              FormatDateTime('YYYYMMDD', bfCal.FieldByName('mdCalEndDate').AsDateTime));
          if bfCal.FieldByName('mdCalStartDate').AsString <> '' then
            WriteLn(flCalendar, 'DTSTART;VALUE=DATE:' +
              FormatDateTime('YYYYMMDD', bfCal.FieldByName(
              'mdCalStartDate').AsDateTime));
          WriteLn(flCalendar, 'CLASS:PUBLIC');
          WriteLn(flCalendar, 'PERCENT-COMPLETE:' +
            bfCal.FieldByName('mdCalCompletion').AsString);
          if UpperCase(bfCal.FieldByName('mdCalState').AsString) = 'DONE' then
            WriteLn(flCalendar, 'STATUS:COMPLETED')
          else if UpperCase(bfCal.FieldByName('mdCalState').AsString) = 'STARTED' then
            WriteLn(flCalendar, 'STATUS:IN-PROCESS');
          // If status is to do, nothing must be added
          WriteLn(flCalendar, 'PRIORITY:' + bfCal.FieldByName(
            'mdCalPriority').AsString);
          WriteLn(flCalendar, 'SEQUENCE:' + IntToStr(iSeq));
          Inc(iSeq);
          WriteLn(flCalendar, 'CREATED:' + FormatDateTime('YYYYMMDD', Now) +
            'T' + FormatDateTime('HHMMSS', Now) + 'Z');
          WriteLn(flCalendar, 'LAST-MODIFIED:' + FormatDateTime('YYYYMMDD', Now) +
            'T' + FormatDateTime('HHMMSS', Now) + 'Z');
          WriteLn(flCalendar, 'END:VTODO');
          WriteLn(flCalendar, 'END:VCALENDAR');
        end;
        bfCal.Next;
      end;
    finally
      bfCal.Close;
      CloseFile(flCalendar);
      Screen.Cursor := crDefault;
    end;
end;

function TfmCalendar.GetDate(stDate, stDtFormat: string): TDate;
begin
  // Set the date in the right format
  try
    stDate := fmMain.ConvertDateFormat(stDate, stDtFormat, fmMain.DateOrder);
    // Set the date separator of the OS
    // necessary for the bfCal component
    if TryStrToDate(StringReplace(stDate, '-', DateSeparator, [rfReplaceAll]),
      Result, FDate.ShortDateFormat) = False then
      Result := 0;
    // This should never happen, anyway...
  except
    Result := 0;
    MessageDlg(StringReplace(fmMain.msg080, '*', stDate, []),
      mtWarning, [mbOK], 0);
  end;
end;

function TfmCalendar.CleanString(stInput: string): string;
begin
  // Escape the comma and CR in the string fields
  Result := StringReplace(stInput, ',', '\,', [rfReplaceAll]);
  Result := StringReplace(stInput, LineEnding, '\n', [rfReplaceAll]);
end;

initialization
  {$I unit9.lrs}

end.





