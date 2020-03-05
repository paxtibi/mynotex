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

unit Unit8;

{$mode objfpc}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ColorBox;

type

  { TfmLook }

  TfmLook = class(TForm)
    bnColOK: TButton;
    bnColDef1: TButton;
    bnColDef2: TButton;
    bnColDef3: TButton;
    bnNoCol: TButton;
    bnColCancel: TButton;
    cbColBack: TColorBox;
    cbColFont: TColorBox;
    lbColLorem1: TLabel;
    lbColBack: TLabel;
    lbColFont: TLabel;
    lbColLorem2: TLabel;
    lbColLorem3: TLabel;
    procedure bnColCancelClick(Sender: TObject);
    procedure bnColDef1Click(Sender: TObject);
    procedure bnColDef2Click(Sender: TObject);
    procedure bnColDef3Click(Sender: TObject);
    procedure bnColOKClick(Sender: TObject);
    procedure bnNoColClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
  private
    { private declarations }
  public
    flSubjectForm: boolean;
    { public declarations }
  end;

var
  fmLook: TfmLook;

implementation

uses Unit1;

{ TfmLook }

procedure TfmLook.FormCreate(Sender: TObject);
begin
  //Change form color
  fmLook.Color := fmMain.Color;
end;

procedure TfmLook.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
    Close;
end;

procedure TfmLook.FormShow(Sender: TObject);
begin
  // Load colors
  if flSubjectForm = True then
  begin
    with fmMain.sqSubjects do
    begin
      if FieldByName('SubjectsFontColor').AsString = '' then
        cbColFont.Selected := clBlack
      else
        cbColFont.Selected := StringToColor(FieldByName('SubjectsFontColor').AsString);
      if FieldByName('SubjectsBackColor').AsString = '' then
        cbColBack.Selected := clWhite
      else
        cbColBack.Selected := StringToColor(FieldByName('SubjectsBackColor').AsString);
    end;
  end
  else
  begin
    with fmMain.sqNotes do
    begin
      if FieldByName('NotesFontColor').AsString = '' then
        cbColFont.Selected := clBlack
      else
        cbColFont.Selected := StringToColor(FieldByName('NotesFontColor').AsString);
      if FieldByName('NotesBackColor').AsString = '' then
        cbColBack.Selected := clWhite
      else
        cbColBack.Selected := StringToColor(FieldByName('NotesBackColor').AsString);
    end;
  end;
  lbColLorem1.Font.Color := StringToColor(fmMain.stColFont1);
  lbColLorem2.Font.Color := StringToColor(fmMain.stColFont2);
  lbColLorem3.Font.Color := StringToColor(fmMain.stColFont3);
  lbColLorem1.Color := StringToColor(fmMain.stColBack1);
  lbColLorem2.Color := StringToColor(fmMain.stColBack2);
  lbColLorem3.Color := StringToColor(fmMain.stColBack3);
end;

procedure TfmLook.bnColDef1Click(Sender: TObject);
begin
  // Default color 1
  cbColBack.Selected := StringToColor(fmMain.stColBack1);
  cbColFont.Selected := StringToColor(fmMain.stColFont1);
end;

procedure TfmLook.bnColDef2Click(Sender: TObject);
begin
  // Default color 2
  cbColBack.Selected := StringToColor(fmMain.stColBack2);
  cbColFont.Selected := StringToColor(fmMain.stColFont2);
end;

procedure TfmLook.bnColDef3Click(Sender: TObject);
begin
  // Default color 3
  cbColBack.Selected := StringToColor(fmMain.stColBack3);
  cbColFont.Selected := StringToColor(fmMain.stColFont3);
end;

procedure TfmLook.bnNoColClick(Sender: TObject);
begin
  // No color
  cbColBack.Selected := clWhite;
  cbColFont.Selected := clBlack;
end;

procedure TfmLook.bnColOKClick(Sender: TObject);
begin
  // Save data
  if flSubjectForm = True then
  begin
    with fmMain.sqSubjects do
    begin
      Edit;
      if ((cbColBack.Colors[cbColBack.ItemIndex] = clWhite) or (cbColBack.Colors[cbColBack.ItemIndex] = clDefault)) then
        FieldByName('SubjectsBackColor').AsString := ''
      else
        FieldByName('SubjectsBackColor').AsString :=
          ColorToString(cbColBack.Colors[cbColBack.ItemIndex]);
      if ((cbColFont.Colors[cbColFont.ItemIndex] = clBlack) or (cbColFont.Colors[cbColFont.ItemIndex] = clDefault)) then
        FieldByName('SubjectsFontColor').AsString := ''
      else
        FieldByName('SubjectsFontColor').AsString :=
          ColorToString(cbColFont.Colors[cbColFont.ItemIndex]);
      Post;
      ApplyUpdates;
    end;
  end
  else
  begin
    with fmMain.sqNotes do
    begin
      Edit;
      if ((cbColBack.Colors[cbColBack.ItemIndex] = clWhite) or (cbColBack.Colors[cbColBack.ItemIndex] = clDefault)) then
        FieldByName('NotesBackColor').AsString := ''
      else
        FieldByName('NotesBackColor').AsString :=
          ColorToString(cbColBack.Colors[cbColBack.ItemIndex]);
      if ((cbColFont.Colors[cbColFont.ItemIndex] = clBlack) or (cbColFont.Colors[cbColFont.ItemIndex] = clDefault)) then
        FieldByName('NotesFontColor').AsString := ''
      else
        FieldByName('NotesFontColor').AsString :=
          ColorToString(cbColFont.Colors[cbColFont.ItemIndex]);
      Post;
      ApplyUpdates;
    end;
  end;
  Close;
end;

procedure TfmLook.bnColCancelClick(Sender: TObject);
begin
  // Close
  Close;
end;

initialization
  {$I unit8.lrs}

end.


