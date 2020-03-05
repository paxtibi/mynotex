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

unit Unit6;

{$mode objfpc}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls;

type

  { TfmResizeImage }

  TfmResizeImage = class(TForm)
    bnSubCommCancel: TButton;
    bnSubCommOK: TButton;
    lbSize: TLabel;
    rgResizeImage: TRadioGroup;
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure FormShow(Sender: TObject);
  private
    procedure SetLabelValues;
    { private declarations }
  public
    { public declarations }
    procedure SetImageHeigth(ImgHeigth: integer);
    procedure SetImageWidth(ImgWidth: integer);
  end;

var
  fmResizeImage: TfmResizeImage;
  ImageWidth, ImageHeigth: integer;
  LabelCaption: string = 'Picture size:';

implementation

// Main form
uses Unit1;

{ TfmResizeImage }

procedure TfmResizeImage.FormShow(Sender: TObject);
begin
  // Change label size values
  // Valid also for OnClick event of TRadioGroup
  SetLabelValues;
end;

procedure TfmResizeImage.FormCreate(Sender: TObject);
begin
  //Change form color
  fmResizeImage.Color := fmMain.Color;
end;

procedure TfmResizeImage.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
  begin
    ModalResult := mrCancel;
  end;
end;

procedure TfmResizeImage.SetImageWidth(ImgWidth: integer);
begin
  //Set image width for label
  ImageWidth := ImgWidth;
end;

procedure TfmResizeImage.SetImageHeigth(ImgHeigth: integer);
begin
  //Set image height for label
  ImageHeigth := ImgHeigth;
end;

procedure TfmResizeImage.SetLabelValues;
var
  LevResize: single;
begin
  // Set label size value
  case rgResizeImage.ItemIndex of
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
  end;
  lbSize.Caption := LabelCaption + ' ' + IntToStr(Trunc(ImageWidth * LevResize)) +
    'x' + IntToStr(Trunc(ImageHeigth * LevResize));
end;

initialization
  {$I unit6.lrs}

end.

