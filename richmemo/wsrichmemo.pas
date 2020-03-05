{
 wsrichmemo.pas 
 
 Author: Dmitry 'skalogryz' Boyarintsev 

 *****************************************************************************
 *                                                                           *
 *  This file is part of the Lazarus Component Library (LCL)                 *
 *                                                                           *
 *  See the file COPYING.modifiedLGPL.txt, included in this distribution,    *
 *  for details about the copyright.                                         *
 *                                                                           *
 *  This program is distributed in the hope that it will be useful,          *
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of           *
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.                     *
 *                                                                           *
 *****************************************************************************
}

unit WSRichMemo; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, 

  Graphics, Controls, StdCtrls,
  
  WSStdCtrls;  
  
type

  TChangedFontItems = set of (fiName, fiSize, fiColor, fiBackcolor, fiItalic,
    fiBold, fiUnderline, fiStrike, fiAlignment, fiIndented); // Added by Massimo Nardello

  TRichAlignment = set of (trLeft, trRight, trCenter, trJustified); // Added by Massimo Nardello

  TIntFontParams = record
     Name      : String;
     Size      : Integer;
     Color     : TColor;
     BackColor : TColor; // Added by Massimo Nardello
     Style     : TFontStyles;
     Alignment : TRichAlignment; // Added by Massimo Nardello
     Indented  : Integer; // Added by Massimo Nardello
     Changed   : TChangedFontItems; // Added by Massimo Nardello
   end;


  { TWSCustomRichMemo }

  TWSCustomRichMemo = class(TWSCustomMemo)
  private
  published

    //Note: RichMemo cannot use LCL TCustomEdit copy/paste/cut operations
    //      because there's no support for (system native) RICHTEXT clipboard format
    //      that's why Clipboard operations are moved to widgetset level
    class procedure CutToClipboard(const AWinControl: TWinControl); virtual;
    class procedure CopyToClipboard(const AWinControl: TWinControl); virtual;
    class procedure PasteFromClipboard(const AWinControl: TWinControl); virtual;
    class function GetStyleRange(const AWinControl: TWinControl; TextStart: Integer; var RangeStart, RangeLen: Integer): Boolean; virtual;
    class function GetTextAttributes(const AWinControl: TWinControl; TextStart: Integer;
      var Params: TIntFontParams): Boolean; virtual;
    class procedure SetDefaultFont(const AWinControl: TWinControl;
      FontName: String; FontSize: Integer); virtual;
    class procedure SetTextAttributes(const AWinControl: TWinControl; TextStart, TextLen: Integer;
      const Params: TIntFontParams); virtual;
    class procedure InDelText(const AWinControl: TWinControl; const TextUTF8: String; DstStart, DstLen: Integer); virtual; 
    class procedure SetHideSelection(const ACustomEdit: TCustomEdit; AHideSelection: Boolean); override;
    class function LoadRichText(const AWinControl: TWinControl; Source: TStream): Boolean; virtual;
    class function SaveRichText(const AWinControl: TWinControl; Dest: TStream): Boolean; virtual;
    class function LoadImageFromFile(const AWinControl: TWinControl;
      TextStart: Integer; FileName: String; ImgScale: Single):
      Boolean; virtual;
    class function SaveImageToFile(const AWinControl: TWinControl;
      TextStart: Integer; FileName: String): Boolean; virtual;
    class function GetImagePosInText(const AWinControl: TWinControl): TStringList; virtual;
    class function ListNumber(const AWinControl: TWinControl;
      TextStart: Integer; Indent: Integer; Bullet: Char): Boolean; virtual;
    class function SetToDo(const AWinControl: TWinControl; CurrLine: Integer;
      Done: Boolean): Boolean; virtual;
    class function InsertText(const AWinControl: TWinControl; Start: Integer;
      TextToInsert: String): Boolean; virtual;
    class function GetWordParagraphStartEnd(const AWinControl: TWinControl;
      CurrPos: Integer; Words: Boolean; Beginning: Boolean): Integer;  virtual;
    class function ClearAll(const AWinControl: TWinControl): Boolean; virtual;
    class function MoveParUpDown(const AWinControl: TWinControl;
      CurrLine: Integer; MoveUp: Boolean): Boolean; virtual;
    class function CountChar(const AWinControl: TWinControl): integer; virtual;
    class function GetUndo(const AWinControl: TWinControl; Start: integer): boolean; virtual;
    class function SetUndo(const AWinControl: TWinControl; Start: integer): boolean; virtual;
    class function ClearUndo(const AWinControl: TWinControl): boolean; virtual;
    class function SetLineSpace(const AWinControl: TWinControl; Space: integer): boolean; virtual;
    class function SetParagraphSpace(const AWinControl: TWinControl; Space: integer): boolean; virtual;
    class function CleanParagraph(const AWinControl: TWinControl; Start: Integer): boolean; virtual;
    class function SetCursorMiddleScreen(const AWinControl: TWinControl; CurrLine: integer): boolean; virtual;

  end;
  TWSCustomRichMemoClass = class of TWSCustomRichMemo;

  
function WSRegisterCustomRichMemo: Boolean; external name 'WSRegisterCustomRichMemo';

implementation

{ TWSCustomRichMemo }

class procedure TWSCustomRichMemo.CutToClipboard(const AWinControl: TWinControl);
begin

end;

class procedure TWSCustomRichMemo.CopyToClipboard(const AWinControl: TWinControl);
begin

end;

class procedure TWSCustomRichMemo.PasteFromClipboard(const AWinControl: TWinControl);
begin

end;

class function TWSCustomRichMemo.GetStyleRange(const AWinControl: TWinControl;
  TextStart: Integer; var RangeStart, RangeLen: Integer): Boolean;
begin
  RangeStart :=-1;
  RangeLen := -1;
  Result := false;
end;

class function TWSCustomRichMemo.GetTextAttributes(const AWinControl: TWinControl; 
  TextStart: Integer; var Params: TIntFontParams): Boolean;
begin
  Result := false;
end;

class procedure TWSCustomRichMemo.SetTextAttributes(const AWinControl: TWinControl; 
  TextStart, TextLen: Integer;  
  {Mask: TTextStyleMask;} const Params: TIntFontParams);
begin
end;

class procedure TWSCustomRichMemo.SetDefaultFont(const AWinControl: TWinControl;
   FontName: String; FontSize: Integer);
begin
  // Added by Massimo Nardello
end;

class procedure TWSCustomRichMemo.InDelText(const AWinControl: TWinControl; const TextUTF8: String; DstStart, DstLen: Integer); 
begin

end;

class procedure TWSCustomRichMemo.SetHideSelection(const ACustomEdit: TCustomEdit; AHideSelection: Boolean); 
begin

end;

class function TWSCustomRichMemo.LoadRichText(const AWinControl: TWinControl; Source: TStream): Boolean;
begin
  Result := false;
end;

class function TWSCustomRichMemo.SaveRichText(const AWinControl: TWinControl; Dest: TStream): Boolean;
begin
  Result := false;
end;

class function TWSCustomRichMemo.LoadImageFromFile(const AWinControl: TWinControl;
   TextStart: Integer; FileName: String; ImgScale: Single): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.SaveImageToFile(const AWinControl: TWinControl;
  TextStart: Integer; FileName: String): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.GetImagePosInText(const AWinControl: TWinControl): TStringList;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.ListNumber(const AWinControl: TWinControl;
  TextStart: Integer; Indent: Integer; Bullet: Char): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.SetToDo(const AWinControl: TWinControl;
    CurrLine: Integer; Done: Boolean): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.InsertText(const AWinControl: TWinControl; Start: Integer;
  TextToInsert: String): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.GetWordParagraphStartEnd(const AWinControl: TWinControl;
  CurrPos: Integer; Words: Boolean; Beginning: Boolean): Integer;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.ClearAll(const AWinControl: TWinControl): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.MoveParUpDown(const AWinControl: TWinControl;
  CurrLine: Integer; MoveUp: Boolean): Boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.CountChar(const AWinControl: TWinControl): integer;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.GetUndo(const AWinControl: TWinControl; Start: integer): boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.SetUndo(const AWinControl: TWinControl; Start: integer): boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.ClearUndo(const AWinControl: TWinControl): boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.SetLineSpace(const AWinControl: TWinControl; Space: integer): boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.SetParagraphSpace(const AWinControl: TWinControl; Space: integer): boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.CleanParagraph(const AWinControl: TWinControl; Start: integer): boolean;
begin
  // Added by Massimo Nardello
end;

class function TWSCustomRichMemo.SetCursorMiddleScreen(const AWinControl: TWinControl; CurrLine: integer): boolean;
begin
  // Added by Massimo Nardello
end;

end.

