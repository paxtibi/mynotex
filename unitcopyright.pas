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

unit UnitCopyright;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls, LCLIntf;

type

  { TfmCopyright }

  TfmCopyright = class(TForm)
    imImagecopyright: TImage;
    lbButton: TLabel;
    lbCopyrightName1: TLabel;
    lbCopyrightSite: TLabel;
    lbCopyrightAuthor2: TLabel;
    lbCopyrightName: TLabel;
    lbCopyrightAuthor1: TLabel;
    lbCopyrightForum: TLabel;
    meCopyrightText: TMemo;
    tmAlarmForm: TTimer;
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
    procedure lbButtonClick(Sender: TObject);
    procedure lbButtonMouseEnter(Sender: TObject);
    procedure lbButtonMouseLeave(Sender: TObject);
    procedure lbCopyrightForumClick(Sender: TObject);
    procedure lbCopyrightSiteClick(Sender: TObject);
    procedure tmAlarmFormTimer(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmCopyright: TfmCopyright;

implementation

{ TfmCopyright }

procedure TfmCopyright.lbCopyrightForumClick(Sender: TObject);
begin
  // Open MyNotex forum
  OpenURL('http://groups.google.com/forum/#!forum/mynotex');
end;

procedure TfmCopyright.lbButtonMouseEnter(Sender: TObject);
begin
  // Button color on enter
  lbButton.Color := $000000D4;
  lbButton.Font.Color := $0026B1F2;
end;

procedure TfmCopyright.lbButtonMouseLeave(Sender: TObject);
begin
  // Button color on exit
  lbButton.Color := $0026B1F2;
  lbButton.Font.Color := $000000D4;
end;

procedure TfmCopyright.lbCopyrightSiteClick(Sender: TObject);
begin
  // Open MyNotex site
  OpenURL('http://sites.google.com/site/mynotex');
end;

procedure TfmCopyright.tmAlarmFormTimer(Sender: TObject);
begin
  // Show the form fading
  if fmCopyright.AlphaBlendValue < 255 then
  begin
    fmCopyright.AlphaBlendValue := fmCopyright.AlphaBlendValue + 1;
  end
  else
  begin
    tmAlarmForm.Enabled := False;
  end;
end;

procedure TfmCopyright.lbButtonClick(Sender: TObject);
begin
  // Close the form
  Close;
end;

procedure TfmCopyright.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
    Close;
end;

procedure TfmCopyright.FormShow(Sender: TObject);
begin
  // Start fading
  fmCopyright.AlphaBlendValue := 0;
  tmAlarmForm.Enabled := True;
end;

initialization
  {$I unitcopyright.lrs}

end.

