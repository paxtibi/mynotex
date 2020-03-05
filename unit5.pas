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

unit Unit5;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, DCPsha1;

type

  { TfmEncryption }

  TfmEncryption = class(TForm)
    bnClose: TButton;
    bnEncrypt: TButton;
    cbShowChar: TCheckBox;
    edPwd1: TEdit;
    edPwd2: TEdit;
    lbEncrypt: TLabel;
    procedure bnCloseClick(Sender: TObject);
    procedure bnEncryptClick(Sender: TObject);
    procedure cbShowCharChange(Sender: TObject);
    procedure edPwd1KeyPress(Sender: TObject; var Key: char);
    procedure edPwd2KeyPress(Sender: TObject; var Key: char);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmEncryption: TfmEncryption;

implementation

uses Unit1;

{ TfmEncryption }

procedure TfmEncryption.FormCreate(Sender: TObject);
begin
  //Change form color
  fmEncryption.Color := fmMain.Color;
end;

procedure TfmEncryption.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
    Close;
end;

procedure TfmEncryption.FormShow(Sender: TObject);
begin
  // Set the focus on pwd field
  edPwd1.SetFocus;
end;

procedure TfmEncryption.cbShowCharChange(Sender: TObject);
begin
  // Toggle show char
  if cbShowChar.Checked = True then
  begin
    edPwd1.PasswordChar := #0;
    edPwd2.PasswordChar := #0;
  end
  else
  begin
    edPwd1.PasswordChar := '*';
    edPwd2.PasswordChar := '*';
  end;
end;

procedure TfmEncryption.edPwd1KeyPress(Sender: TObject; var Key: char);
begin
  // Change field with Return
  if Key = #13 then
    edPwd2.SetFocus;
end;

procedure TfmEncryption.edPwd2KeyPress(Sender: TObject; var Key: char);
begin
  // Execute encryption on Return
  if Key = #13 then
    bnEncryptClick(nil);
end;

procedure TfmEncryption.bnCloseClick(Sender: TObject);
begin
  // Close the form
  Close;
end;

procedure TfmEncryption.bnEncryptClick(Sender: TObject);
begin
  // Encrypt
  if edPwd1.Text <> edPwd2.Text then
  begin
    MessageDlg(fmMain.msg040, mtWarning, [mbOK], 0);
    edPwd1.Text := '';
    edPwd2.Text := '';
    edPwd1.SetFocus;
  end
  else
    try
      Screen.Cursor := crHourGlass;
      Application.ProcessMessages;
      fmMain.sqNotes.Edit;
      fmMain.dcAES.InitStr(edPwd1.Text, TDCP_sha1);
      fmMain.sqNotes.FieldByName('NotesText').AsString :=
        fmMain.dcAES.EncryptString(fmMain.sqNotes.FieldByName('NotesText').AsString);
      fmMain.dcAES.InitStr(edPwd1.Text, TDCP_sha1);
      fmMain.sqNotes.FieldByName('NotesCheckPwd').AsString :=
        fmMain.dcAES.EncryptString(fmMain.PwdCheckString);
      fmMain.sqNotes.Post;
      fmMain.sqNotes.ApplyUpdates;
    finally
      Screen.Cursor := crDefault;
      Close;
    end;
end;

initialization
  {$I unit5.lrs}

end.
