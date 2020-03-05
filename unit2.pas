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

unit Unit2;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Sqlite3DS, FileUtil, LResources, Forms, Controls,
  Graphics, Dialogs, CheckLst, StdCtrls, ComCtrls, LCLProc;

type

  { TfmImpExp }

  TfmImpExp = class(TForm)
    bnImpExp: TButton;
    bnClose: TButton;
    cbNoExpDate: TCheckBox;
    cbReadSubjects: TCheckListBox;
    cbDeleteData: TCheckBox;
    cbSelDeselAll: TCheckBox;
    lbSubjects: TLabel;
    pbProgressBar: TProgressBar;
    sqReadDelRec: TSqlite3Dataset;
    sqWriteDelRec: TSqlite3Dataset;
    sqReadSubjects: TSqlite3Dataset;
    sqReadNotes: TSqlite3Dataset;
    sqWriteSubjects: TSqlite3Dataset;
    sqWriteNotes: TSqlite3Dataset;
    procedure bnCloseClick(Sender: TObject);
    procedure cbSelDeselAllClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure ImportExport(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmImpExp: TfmImpExp;

implementation

// Main form
uses Unit1;

{ TfmImpExp }

procedure TfmImpExp.bnCloseClick(Sender: TObject);
begin
  // Quit
  Close;
end;

procedure TfmImpExp.cbSelDeselAllClick(Sender: TObject);
var
  i: integer;
begin
  // Select and deselect all the items
  for i := 0 to cbReadSubjects.Items.Count - 1 do
    cbReadSubjects.Checked[i] := cbSelDeselAll.Checked;
end;

procedure TfmImpExp.FormCreate(Sender: TObject);
begin
  //Change form color
  fmImpExp.Color := fmMain.Color;
end;

procedure TfmImpExp.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  // Close all tables and clear subjects list
  sqReadSubjects.Close;
  sqReadNotes.Close;
  sqWriteSubjects.Close;
  sqWriteNotes.Close;
  cbReadSubjects.Items.Clear;
  pbProgressBar.Position := 0;
end;

procedure TfmImpExp.ImportExport(Sender: TObject);
var
  i, n, IDPos, FnSize, HTMLFnSize, NumList, IniIndent, EndIndent, SpIndent, PosName: integer;
  SubNotesText, SubActText, AttReadDir, AttWriteDir, stSpcInd: string;
  myFile: TextFile;
  myStringList: TStringList;
  flNumList: boolean = False;
begin
  // Import and export procedure
  // Set progressbar values
  pbProgressBar.Max := 0;
  for i := 0 to cbReadSubjects.Items.Count - 1 do
  begin
    if cbReadSubjects.Checked[i] = True then
    begin
      pbProgressBar.Max := pbProgressBar.Max + 1;
    end;
  end;
  if pbProgressBar.Max = 0 then
  begin
    MessageDlg(fmMain.msg023, mtWarning, [mbOK], 0);
    Abort;
  end;
  Screen.Cursor := crHourGlass;
  // No HTML exportation
  if sqWriteSubjects.Active = True then
    try
      for i := 0 to cbReadSubjects.Items.Count - 1 do
      begin
        // The subject is checked
        if cbReadSubjects.Checked[i] = True then
        begin
          // Check if subject is already present
          sqWriteSubjects.Active := False;
          sqWriteSubjects.SQL :=
            'Select * from Subjects where SubjectsUID = "' + sqReadSubjects.FieldByName('SubjectsUID').AsString + '"';
          sqWriteSubjects.Active := True;
          if sqWriteSubjects.RecordCount > 0 then
          begin
            MessageDlg(sqReadSubjects.FieldByName('SubjectsName').AsString + ' ' + fmMain.msg036, mtInformation, [mbOK], 0);
          end
          else
          begin
            // Add subject
            sqWriteSubjects.Append;
            sqWriteSubjects.FieldByName('SubjectsName').AsString :=
              sqReadSubjects.FieldByName('SubjectsName').AsString;
            sqWriteSubjects.FieldByName('SubjectsComments').AsString :=
              sqReadSubjects.FieldByName('SubjectsComments').AsString;
            sqWriteSubjects.FieldByName('SubjectsBackColor').AsString :=
              sqReadSubjects.FieldByName('SubjectsBackColor').AsString;
            sqWriteSubjects.FieldByName('SubjectsFontColor').AsString :=
              sqReadSubjects.FieldByName('SubjectsFontColor').AsString;
            sqWriteSubjects.FieldByName('SubjectsSort').AsInteger :=
              sqWriteSubjects.FieldByName('IDSubjects').AsInteger;
            sqWriteSubjects.FieldByName('SubjectsUID').AsString :=
              sqReadSubjects.FieldByName('SubjectsUID').AsString;
            sqWriteSubjects.FieldByName('SubjectsDTMod').AsDateTime :=
              sqReadSubjects.FieldByName('SubjectsDTMod').AsDateTime;
            sqWriteSubjects.Post;
            sqWriteSubjects.ApplyUpdates;
            // A deleted subject in an import or export operation
            // with deletion of original records
            // could be imported again, and must not be deleted
            with sqWriteDelRec do
            begin
              Open;
              if Locate('DelRecUID', sqReadSubjects.FieldByName('SubjectsUID').AsString, []) = True then
              begin
                Delete;
                ApplyUpdates;
              end;
              Close;
            end;
            Application.ProcessMessages;
            pbProgressBar.Position := pbProgressBar.Position + 1;
            Application.ProcessMessages;
            // Add notes
            if fmMain.miOrderByTitle.Checked = True then
              sqReadNotes.SQL :=
                'Select * from Notes where ID_Subjects = ' + sqReadSubjects.FieldByName('IDSubjects').AsString + ' order by NotesTitle, IDNotes'
            else
              sqReadNotes.SQL :=
                'Select * from Notes where ID_Subjects = ' + sqReadSubjects.FieldByName('IDSubjects').AsString + ' order by NotesDate, IDNotes';
            sqReadNotes.Open;
            while not sqReadNotes.EOF do
            begin
              sqWriteNotes.Append;
              sqWriteNotes.FieldByName('ID_Subjects').AsInteger :=
                sqWriteSubjects.FieldByName('IDSubjects').AsInteger;
              sqWriteNotes.FieldByName('NotesTitle').AsString :=
                sqReadNotes.FieldByName('NotesTitle').AsString;
              sqWriteNotes.FieldByName('NotesDate').AsDateTime :=
                sqReadNotes.FieldByName('NotesDate').AsDateTime;
              sqWriteNotes.FieldByName('NotesText').AsString :=
                sqReadNotes.FieldByName('NotesText').AsString;
              sqWriteNotes.FieldByName('NotesActivities').AsString :=
                sqReadNotes.FieldByName('NotesActivities').AsString;
              sqWriteNotes.FieldByName('NotesTags').AsString :=
                sqReadNotes.FieldByName('NotesTags').AsString;
              sqWriteNotes.FieldByName('NotesBackColor').AsString :=
                sqReadNotes.FieldByName('NotesBackColor').AsString;
              sqWriteNotes.FieldByName('NotesFontColor').AsString :=
                sqReadNotes.FieldByName('NotesFontColor').AsString;
              sqWriteNotes.FieldByName('NotesSort').AsInteger :=
                sqWriteNotes.FieldByName('IDNotes').AsInteger;
              sqWriteNotes.FieldByName('NotesAttName').AsString :=
                sqReadNotes.FieldByName('NotesAttName').AsString;
              sqWriteNotes.FieldByName('NotesUID').AsString :=
                sqReadNotes.FieldByName('NotesUID').AsString;
              sqWriteNotes.FieldByName('NotesDTMod').AsDateTime :=
                sqReadNotes.FieldByName('NotesDTMod').AsDateTime;
              sqWriteNotes.FieldByName('NotesDateFormat').AsString :=
                sqReadNotes.FieldByName('NotesDateFormat').AsString;
              sqWriteNotes.FieldByName('NotesCheckPwd').AsString :=
                sqReadNotes.FieldByName('NotesCheckPwd').AsString;
              sqWriteNotes.Post;
              sqWriteNotes.ApplyUpdates;
              // A deleted note in an import or export operation
              // with deletion of original records
              // could be imported again, and must not be deleted
              with sqWriteDelRec do
              begin
                Open;
                if Locate('DelRecUID', sqReadNotes.FieldByName('NotesUID').AsString, []) = True then
                begin
                  Delete;
                  ApplyUpdates;
                end;
                Close;
              end;
              // Copy attachments
              if sqReadNotes.FieldByName('NotesAttName').AsString <> '' then
              begin
                AttReadDir := ExtractFileNameWithoutExt(sqReadSubjects.FileName);
                if DirectoryExistsUTF8(AttReadDir) = False then
                begin
                  MessageDlg(fmMain.msg030, mtWarning, [mbOK], 0);
                  Abort;
                end;
                AttWriteDir := ExtractFileNameWithoutExt(sqWriteSubjects.FileName);
                if DirectoryExistsUTF8(AttWriteDir) = False then
                  try
                    CreateDir(AttWriteDir);
                  except
                    MessageDlg(fmMain.msg029, mtWarning, [mbOK], 0);
                    Screen.Cursor := crDefault;
                    Abort;
                  end;
                myStringList := TStringList.Create;
                myStringList.Text := sqReadNotes.FieldByName('NotesAttName').AsString;
                for n := 0 to myStringList.Count - 1 do
                  CopyFile(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[n]) + '.zip',
                    AttWriteDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[n]) + '.zip');
                myStringList.Free;
              end;
              // Copy pictures
              AttReadDir := ExtractFileNameWithoutExt(sqReadSubjects.FileName);
              if FileExistsUTF8(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img0000.jpeg') then
              begin
                AttWriteDir := ExtractFileNameWithoutExt(sqWriteSubjects.FileName);
                if DirectoryExistsUTF8(AttWriteDir) = False then
                  try
                    CreateDir(AttWriteDir);
                  except
                    MessageDlg(fmMain.msg029, mtWarning, [mbOK], 0);
                    Screen.Cursor := crDefault;
                    Abort;
                  end;
                n := 0;
                while FileExistsUTF8(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg') do
                begin
                  CopyFile(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg',
                    AttWriteDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg');
                  Inc(n);
                end;
              end;
              sqReadNotes.Next;
              Application.ProcessMessages;
            end;
            sqReadNotes.Close;
          end;
        end;
        sqReadSubjects.Next;
        Application.ProcessMessages;
      end;
      // Delete imported/exported data from original file
      if cbDeleteData.Checked = True then
      begin
        pbProgressBar.Position := 0;
        sqReadSubjects.First;
        for i := 0 to cbReadSubjects.Items.Count - 1 do
        begin
          // The subject is checked
          if cbReadSubjects.Checked[i] = True then
          begin
            // Delete notes
            sqReadNotes.SQL :=
              'Select * from Notes where ID_Subjects = ' + sqReadSubjects.FieldByName('IDSubjects').AsString;
            sqReadNotes.Open;
            while not sqReadNotes.EOF do
            begin
              // Delete attachment
              if sqReadNotes.FieldByName('NotesAttName').AsString <> '' then
                try
                  AttReadDir := ExtractFileNameWithoutExt(sqReadSubjects.FileName);
                  if DirectoryExistsUTF8(AttReadDir) = False then
                  begin
                    MessageDlg(fmMain.msg030, mtWarning, [mbOK], 0);
                    Abort;
                  end;
                  myStringList := TStringList.Create;
                  myStringList.Text := sqReadNotes.FieldByName('NotesAttName').AsString;
                  for n := 0 to myStringList.Count - 1 do
                    DeleteFileUTF8(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-' + ExtractFileNameOnly(myStringList[n]) + '.zip');
                  myStringList.Free;
                  if fmMain.IsDirectoryEmpty(AttReadDir) = True then
                    DeleteDirectory(AttReadDir, False);
                except
                  MessageDlg(fmMain.msg035, mtWarning, [mbOK], 0);
                  Abort;
                end;
              // Delete images
              n := 0;
              try
                AttReadDir := ExtractFileNameWithoutExt(sqReadSubjects.FileName);
                while FileExistsUTF8(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg') = True do
                begin
                  DeleteFileUTF8(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg');
                  Inc(n);
                end;
                if fmMain.IsDirectoryEmpty(AttReadDir) = True then
                  DeleteDirectory(AttReadDir, False);
              except
                MessageDlg(fmMain.msg035, mtWarning, [mbOK], 0);
                Abort;
              end;
              with sqReadDelRec do
              begin
                // DelRec must be opened and closed each time to read the modification
                // made by a sync operation by another possible session of mynotex.
                Open;
                Append;
                sqReadDelRec.FieldByName('DelRecUID').AsString :=
                  sqReadNotes.FieldByName('NotesUID').AsString;
                sqReadDelRec.FieldByName('DelRecDTMod').AsDateTime := Now;
                Post;
                ApplyUpdates;
                Close;
              end;
              sqReadNotes.Delete;
              sqReadNotes.ApplyUpdates;
              Application.ProcessMessages;
            end;
            sqReadNotes.Close;
            // Delete subject
            IDPos := sqReadSubjects.RecNo;
            with sqReadDelRec do
            begin
              Open;
              Append;
              sqReadDelRec.FieldByName('DelRecUID').AsString :=
                sqReadSubjects.FieldByName('SubjectsUID').AsString;
              sqReadDelRec.FieldByName('DelRecDTMod').AsDateTime := Now;
              Post;
              ApplyUpdates;
              Close;
            end;
            sqReadSubjects.Delete;
            sqReadSubjects.ApplyUpdates;
            // Sometime after Delete sqReadNotes moves to the first record; so...
            if IDPos <= sqReadSubjects.RecordCount then
            begin
              sqReadSubjects.RecNo := IDPos;
            end;
          end
          else
          begin
            sqReadSubjects.Next;
          end;
          pbProgressBar.Position := pbProgressBar.Position + 1;
          Application.ProcessMessages;
        end;
      end;
      //Update data
      fmMain.sqSubjects.Close;
      fmMain.sqSubjects.Open;
      // sqNotes is closed and reopened in the sqSubjects.AfterScroll event
      // only if there are subjects; so I have to update sqNotes
      // because if all the subjects are deleted the sqNotes dataset is not updated
      fmMain.sqNotes.Close;
      fmMain.dbText.Clear;
      fmMain.sqNotes.Open;
      fmMain.CreateTagsList;
      //    MessageDlg(fmMain.msg024, mtInformation, [mbOK], 0);
      Screen.Cursor := crDefault;
      Close;
    except
      MessageDlg(fmMain.msg018, mtWarning, [mbOK], 0);
      Screen.Cursor := crDefault;
      Close;
    end

  // HTML exportation
  else
    try
      i := 0;
      AssignFile(myFile, fmMain.sdSaveDialog.FileName);
      ReWrite(myFile);
      WriteLn(myFile, '<HTML>');
      WriteLn(myFile, '   <HEAD>');
      WriteLn(myFile, '   <meta http-equiv="Content-Type" ' + 'content="text/html; charset=UTF-8">');
      WriteLn(myFile, '      <TITLE>' + ExtractFileNameOnly(fmMain.sqSubjects.FileName) + '</TITLE>');
      WriteLn(myFile, '   </HEAD>');
      WriteLn(myFile, '   <BODY>');
      while not sqReadSubjects.EOF do
      begin
        // The subject is checked
        if cbReadSubjects.Checked[i] = True then
        begin
          // Add subject
          SubNotesText := sqReadSubjects.FieldByName('SubjectsName').AsString;
          SubNotesText := StringReplace(SubNotesText, '&', '&amp;', [rfReplaceAll]);
          SubNotesText := StringReplace(SubNotesText, '<', '&lt;', [rfReplaceAll]);
          SubNotesText := StringReplace(SubNotesText, '>', '&gt;', [rfReplaceAll]);
          WriteLn(myFile, '<H1>' + SubNotesText + '</H1>');
          pbProgressBar.Position := pbProgressBar.Position + 1;
          Application.ProcessMessages;
          // Add notes
          if fmMain.miOrderByTitle.Checked = True then
            sqReadNotes.SQL :=
              'Select * from Notes where ID_Subjects = ' + sqReadSubjects.FieldByName('IDSubjects').AsString + ' order by NotesTitle, IDNotes'
          else
            sqReadNotes.SQL :=
              'Select * from Notes where ID_Subjects = ' + sqReadSubjects.FieldByName('IDSubjects').AsString + ' order by NotesDate, IDNotes';
          sqReadNotes.Open;
          while not sqReadNotes.EOF do
          begin
            // Add title
            SubNotesText := sqReadNotes.FieldByName('NotesTitle').AsString;
            SubNotesText := StringReplace(SubNotesText, '&', '&amp;', [rfReplaceAll]);
            SubNotesText := StringReplace(SubNotesText, '<', '&lt;', [rfReplaceAll]);
            SubNotesText := StringReplace(SubNotesText, '>', '&gt;', [rfReplaceAll]);
            WriteLn(myFile, '<H2>' + SubNotesText + '</H2>');
            // Add date
            if cbNoExpDate.Checked = True then
              WriteLn(myFile, '<H3>' + sqReadNotes.FieldByName('NotesDate').AsString + '</H3>');
            // Add text
            // Text is not encrypted
            if sqReadNotes.FieldByName('NotesCheckPwd').AsString = '' then
            begin
              SubNotesText := sqReadNotes.FieldByName('NotesText').AsString;
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
                SubNotesText :=
                  StringReplace(SubNotesText, 'size="' + IntToStr(FnSize) + '" face="', 'size="' + IntToStr(HTMLFnSize) + '" face="', [rfIgnoreCase, rfReplaceAll]);
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
                SubNotesText :=
                  StringReplace(SubNotesText, #3, '</li><li>', [rfReplaceAll]);
                if flNumList = True then
                begin
                  SubNotesText :=
                    StringReplace(SubNotesText, '<blockquote>', '<p align=left><ol><li>', []);
                  SubNotesText :=
                    StringReplace(SubNotesText, '</blockquote>', '</li></ol></p>', []);
                end
                else
                begin
                  SubNotesText :=
                    StringReplace(SubNotesText, '<blockquote>', '<p align=left><ul><li>', []);
                  SubNotesText :=
                    StringReplace(SubNotesText, '</blockquote>', '</li></ul></p>', []);
                end;
              end;
              // A list has been processed
              if IniIndent > 0 then
              begin
                // Delete numbers, letters and bullets
                for NumList := 1 to 100 do
                begin
                  SubNotesText :=
                    StringReplace(SubNotesText, IntToStr(NumList) + '.' + #9, '', [rfReplaceAll]);
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
                SubNotesText :=
                  StringReplace(SubNotesText, '•' + #9, '', [rfReplaceAll]);
                // Remove alignment tags within a list to avoid empty rows
                // Alignment is always left
                SubNotesText :=
                  StringReplace(SubNotesText, '<li><p align=left>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li><p align=center>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li><p align=right>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li><p align=justify>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li></p><p align=left>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li></p><p align=center>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li></p><p align=right>', '<li>', [rfReplaceAll]);
                SubNotesText :=
                  StringReplace(SubNotesText, '<li></p><p align=justify>', '<li>', [rfReplaceAll]);
              end;
              SubNotesText := StringReplace(SubNotesText, LineEnding, '<p>', [rfReplaceAll]);
            end
            else
            begin
              SubNotesText := fmMain.msg039;
            end;
            SubNotesText := StringReplace(SubNotesText, #5, '&lt;', [rfReplaceAll]);
            SubNotesText := StringReplace(SubNotesText, #6, '&gt;', [rfReplaceAll]);
            WriteLn(myFile, SubNotesText);
            // Copy pictures
            n := 0;
            AttReadDir := ExtractFileNameWithoutExt(sqReadSubjects.FileName);
            AttWriteDir := ExtractFileDir(fmMain.sdSaveDialog.FileName);
            while FileExistsUTF8(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg') do
            begin
              CopyFile(AttReadDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg',
                AttWriteDir + DirectorySeparator + sqReadNotes.FieldByName('NotesUID').AsString + '-img' + FormatFloat('0000', n) + '.jpeg');
              Inc(n);
            end;
            // Save activities
            if sqReadNotes.FieldByName('NotesActivities').AsString <> '' then
            begin
              WriteLn(myFile, '<hr width="30%" size=1 color="grey">');
              WriteLn(myFile, '<p><table>');
              WriteLn(myFile, '<tr>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbWbs + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbActivity + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbState + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbStartDate + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbEndDate + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbDuration + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbResources + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbPriority + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbCompletion + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbCost + '</font></th>');
              WriteLn(myFile, '  <th><font size="1">' + fmMain.lbNotes + '</font></th>');
              WriteLn(myFile, '</tr>');
              SubActText := sqReadNotes.FieldByName('NotesActivities').AsString;
              WriteLn(myFile, '<tr>');
              // Remove code column of the first row
              SubActText := UTF8Copy(SubActText, UTF8Pos(#9, SubActText) + 1, UTF8Length(SubActText));
              PosName := 1;
              stSpcInd := '';
              while UTF8Length(SubActText) > 0 do
              begin
                // The field is activity name, so is indented
                if PosName = 2 then
                  WriteLn(myFile, '<td><font size="1">' + stSpcInd + Copy(SubActText, 1, Pos(#9, SubActText) - 1) + '</font></td>')
                // The field is not activity name
                else
                  WriteLn(myFile, '<td><font size="1">' + Copy(SubActText, 1, Pos(#9, SubActText) - 1) + '</font></td>');
                SubActText := UTF8Copy(SubActText, UTF8Pos(#9, SubActText) + 1, UTF8Length(SubActText));
                Inc(PosName);
                if UTF8Copy(SubActText, 1, 1) = LineEnding then
                begin
                  SubActText := UTF8Copy(SubActText, 2, UTF8Length(SubActText));
                  // Get the code and remove code column
                  if fmMain.CheckNumbers(UTF8Copy(SubActText, 2, UTF8Pos(#9, SubActText) - 2)) = True then
                  begin
                    SpIndent :=
                      StrToInt(UTF8Copy(SubActText, 2, UTF8Pos(#9, SubActText) - 2));
                    stSpcInd := '';
                    for i := 1 to SpIndent do
                      stSpcInd := stSpcInd + '&nbsp;';
                  end;
                  SubActText := UTF8Copy(SubActText, UTF8Pos(#9, SubActText) + 1, UTF8Length(SubActText));
                  WriteLn(myFile, '</tr>' + LineEnding + '<tr>');
                  PosName := 0;
                end;
              end;
              WriteLn(myFile, '</tr>');
              WriteLn(myFile, '</table>');
            end;
            WriteLn(myFile, '<hr width="100%" size=1 color="black">');
            sqReadNotes.Next;
            Application.ProcessMessages;
          end;
          sqReadNotes.Close;
        end;
        Inc(i);
        sqReadSubjects.Next;
        Application.ProcessMessages;
      end;
      WriteLn(myFile, '   </BODY>');
      WriteLn(myFile, '</HTML>');
      CloseFile(myFile);
      Screen.Cursor := crDefault;
      //    MessageDlg(fmMain.msg024, mtInformation, [mbOK], 0);
      Close;
    except;
      Screen.Cursor := crDefault;
      MessageDlg(fmMain.msg025, mtWarning, [mbOK], 0);
      CloseFile(myFile);
      Close;
    end;
  sqReadSubjects.Close;
  sqReadNotes.Close;
  sqReadDelRec.Close;
  sqWriteSubjects.Close;
  sqWriteNotes.Close;
  sqWriteDelRec.Close;
end;

initialization
  {$I unit2.lrs}

end.
