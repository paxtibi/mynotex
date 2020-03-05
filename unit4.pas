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

unit Unit4;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  DBCtrls, StdCtrls, DB;

type

  { TfmCommentsSubjects }

  TfmCommentsSubjects = class(TForm)
    bnSubCommOK: TButton;
    bnSubCommCancel: TButton;
    dbSubName: TDBEdit;
    dbSubComments: TDBMemo;
    lbSubName: TLabel;
    lbComments: TLabel;
    procedure bnSubCommCancelClick(Sender: TObject);
    procedure bnSubCommOKClick(Sender: TObject);
    procedure dbSubCommentsKeyDown(Sender: TObject; var Key: word;
      Shift: TShiftState);
    procedure FormActivate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmCommentsSubjects: TfmCommentsSubjects;

implementation

uses Unit1;

{ TfmCommentsSubjects }

procedure TfmCommentsSubjects.FormCreate(Sender: TObject);
begin
  //Change form color
  fmCommentsSubjects.Color := fmMain.Color;
end;

procedure TfmCommentsSubjects.bnSubCommOKClick(Sender: TObject);
begin
  // Save data and Close
  if fmMain.dsSubjects.State in [dsEdit, dsInsert] then
  begin
    fmMain.sqSubjects.Post;
    fmMain.sqSubjects.ApplyUpdates;
  end;
  Close;
end;

procedure TfmCommentsSubjects.bnSubCommCancelClick(Sender: TObject);
begin
  // Close
  Close;
end;

procedure TfmCommentsSubjects.dbSubCommentsKeyDown(Sender: TObject;
  var Key: word; Shift: TShiftState);
begin
  // Put in edit the Memo (sometimes it doesn't)
  if fmMain.dsSubjects.State in [dsBrowse] then
    fmMain.sqSubjects.Edit;
end;

procedure TfmCommentsSubjects.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
    Close;
end;

procedure TfmCommentsSubjects.FormActivate(Sender: TObject);
begin
  // Set focus on Memo
  dbSubName.SetFocus;
end;

procedure TfmCommentsSubjects.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  // Undo
  if fmMain.dsSubjects.State in [dsEdit, dsInsert] then
  begin
    fmMain.sqSubjects.Cancel;
  end;
end;

initialization
  {$I unit4.lrs}

end.


