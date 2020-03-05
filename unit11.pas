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
unit Unit11;

{$mode objfpc}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls;

type

  { TfmShowAlarm }

  TfmShowAlarm = class(TForm)
    imClock: TImage;
    lbButton: TLabel;
    lbMessage: TLabel;
    spBox: TShape;
    tmAlarmForm: TTimer;
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure lbButtonClick(Sender: TObject);
    procedure lbButtonMouseEnter(Sender: TObject);
    procedure lbButtonMouseLeave(Sender: TObject);
    procedure tmAlarmFormTimer(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmShowAlarm: TfmShowAlarm;
  iShowButton: integer = 0;

implementation

{ TfmShowAlarm }

procedure TfmShowAlarm.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  // Deactivate the timer
  tmAlarmForm.Enabled := False;
  fmShowAlarm.AlphaBlendValue := 255;
  iShowButton := 0;
  lbButton.Visible := False;
end;

procedure TfmShowAlarm.FormShow(Sender: TObject);
begin
  // Activate the timer
  tmAlarmForm.Enabled := True;
end;

procedure TfmShowAlarm.lbButtonClick(Sender: TObject);
begin
  // Close the form
  Close;
end;

procedure TfmShowAlarm.lbButtonMouseEnter(Sender: TObject);
begin
  // Button color on enter
  lbButton.Color := clMaroon;
end;

procedure TfmShowAlarm.lbButtonMouseLeave(Sender: TObject);
begin
  // Button color on exit
  lbButton.Color := $000000D4;
end;

procedure TfmShowAlarm.tmAlarmFormTimer(Sender: TObject);
begin
  // Show the form fading
  if fmShowAlarm.AlphaBlendValue > 0 then
  begin
    fmShowAlarm.AlphaBlendValue := fmShowAlarm.AlphaBlendValue - 1;
  end
  else
  begin
    fmShowAlarm.AlphaBlendValue := 255;
    Inc(iShowButton);
    if iShowButton = 2 then
    begin
      lbButton.Visible := True;
    end;
  end;
  fmShowAlarm.BringToFront;
end;

initialization
  {$I unit11.lrs}

end.


