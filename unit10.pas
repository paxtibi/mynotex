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
unit Unit10;

{$mode objfpc}

interface

uses
  Classes, SysUtils, FileUtil, ZVDateTimePicker, LResources, Forms, Controls,
  Graphics, Dialogs, StdCtrls, ExtCtrls;

type

  { TfmSetAlarm }

  TfmSetAlarm = class(TForm)
    bnOKAlarm: TButton;
    cbAlarm: TCheckBox;
    imAlarm: TImage;
    tpAlarm: TZVDateTimePicker;
    procedure cbAlarmChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmSetAlarm: TfmSetAlarm;

implementation

uses Unit1;

{ TfmSetAlarm }

procedure TfmSetAlarm.FormCreate(Sender: TObject);
begin
  //Change form color
  fmSetAlarm.Color := fmMain.Color;
end;

procedure TfmSetAlarm.cbAlarmChange(Sender: TObject);
begin
  if cbAlarm.Checked = True then
  begin
    fmMain.tmAlarm.Enabled := True;
  end
  else
  begin
    fmMain.tmAlarm.Enabled := False;
    fmMain.sbStatusBar.Panels[1].Text := '';
  end;
end;

procedure TfmSetAlarm.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
  begin
    cbAlarm.Checked := False;
    Close;
  end
  else if key = #13 then
  begin
    cbAlarm.Checked := True;
    Close;
  end;
end;

procedure TfmSetAlarm.FormShow(Sender: TObject);
var
  Fds: TFormatSettings;
begin
  // Set hour and select time field
  Fds := DefaultFormatSettings;
  Fds.LongTimeFormat := 'hh:nn:00';
  if fmMain.fl24Hour = True then
  begin
    tpAlarm.TimeFormat := tf24;
  end;
  if cbAlarm.Checked = False then
  begin
    tpAlarm.Time := StrToTime(TimeToStr(Time, Fds));
  end;
  tpAlarm.SetFocus;
end;

initialization
  {$I unit10.lrs}

end.
