{
 richmemo.pp
 
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

unit RichMemo; 

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Graphics, StdCtrls, 
  WSRichMemo; 

type

  TFontParams = TIntFontParams;

//  TChangedFontItems = TIntChangedFontItems; // Added by Massimo Nardello
//  TRichAlignment = TIntRichAlignment; // Added by Massimo Nardello

  {TIntFontParams = record // declared at WSRichMemo
     Name      : String;
     Size      : Integer;
     Color     : TColor;
     BackColor : TColor; // Added by Massimo
     Style     : TFontStyles;
     Alignment : TIntRichAlignment; // Added by Massimo Nardello
     Indented  : Boolean; // Added by Massimo Nardello
     Changed   : TIntChangedFontItems; // Added by Massimo
   end; }

  TTextModifyMask  = set of (tmm_Color, tmm_Name, tmm_Size, tmm_Styles);

  { TCustomRichMemo }

  TCustomRichMemo = class(TCustomMemo)
  private
    fHideSelection  : Boolean;
  protected
    class procedure WSRegisterClass; override;
    procedure CreateWnd; override;    
    procedure UpdateRichMemo; virtual;
    procedure SetHideSelection(AValue: Boolean);
    function GetContStyleLength(TextStart: Integer): Integer;
    procedure SetSelText(const SelTextUTF8: string); override;
  public
    procedure CopyToClipboard; override;
    procedure CutToClipboard; override;
    procedure PasteFromClipboard; override;
    procedure SetTextAttributes(TextStart, TextLen: Integer; const TextParams: TFontParams); virtual;
    function GetTextAttributes(TextStart: Integer; var TextParams: TFontParams): Boolean; virtual;
    procedure SetDefaultFont(FontName: String; FontSize: Integer); virtual;
    function GetStyleRange(CharOfs: Integer; var RangeStart, RangeLen: Integer): Boolean; virtual;
    // Removed by Massimo Nardello
    {procedure SetTextAttributes(TextStart, TextLen: Integer; AFont: TFont);}
    procedure SetRangeColor(TextStart, TextLength: Integer; FontColor: TColor);
    procedure SetRangeParams(TextStart, TextLength: Integer; ModifyMask: TTextModifyMask;
      const FontName: String; FontSize: Integer; FontColor: TColor;
      AddFontStyle, RemoveFontStyle: TFontStyles);
    function LoadImageFromFile(TextStart: Integer; FileName: String;
      ImgScale: Single): Boolean; virtual;
    function SaveImageToFile(TextStart: Integer; FileName: String): Boolean; virtual;
    function GetImagePosInText: TStringList; virtual;
    function ListNumber(TextStart: Integer; Indent: Integer; Bullet: Char): Boolean; virtual;
    function SetToDo(CurrLine: Integer; Done: Boolean): Boolean; virtual;
    function InsertText(Start: Integer; TextToInsert: String): Boolean; virtual;
    function GetWordParagraphStartEnd(CurrPos: Integer; Words: Boolean; Beginning: Boolean): Integer; virtual;
    function ClearAll: Boolean; virtual;
    function LoadRichText(Source: TStream): Boolean; virtual;
    function SaveRichText(Dest: TStream): Boolean; virtual;
    function MoveParUpDown(CurrLine: Integer; MoveUp: Boolean): Boolean; virtual;
    function CountChar: integer;  virtual;
    function SetCursorMiddleScreen(CurrLine: integer): boolean; virtual;
    function GetUndo(Start: integer): boolean; virtual;
    function SetUndo(Start: integer): boolean; virtual;
    function ClearUndo: boolean; virtual;
    function SetLineSpace(Space: integer): boolean; virtual;
    function SetParagraphSpace(Space: integer): boolean; virtual;
    function CleanParagraph(Start: integer): boolean; virtual;
    property HideSelection : Boolean read fHideSelection write SetHideSelection;
  end;
  
  TRichMemo = class(TCustomRichMemo)
  published
    property Align;
    property Alignment;
    property Anchors;
    property BidiMode;
    property BorderSpacing;
    property BorderStyle;
    property Color;
    property Constraints;
    property DragCursor;
    property DragKind;
    property DragMode;
    property Enabled;
    property Font;
    property HideSelection;
    property Lines;
    property MaxLength;
    property OnChange;
    property OnClick;
    property OnContextPopup;
    property OnDblClick;
    property OnDragDrop;
    property OnDragOver;
    property OnEditingDone;
    property OnEndDrag;
    property OnEnter;
    property OnExit;
    property OnKeyDown;
    property OnKeyPress;
    property OnKeyUp;
    property OnMouseDown;
    property OnMouseEnter;
    property OnMouseLeave;
    property OnMouseMove;
    property OnMouseUp;
    property OnMouseWheel;
    property OnMouseWheelDown;
    property OnMouseWheelUp;
    property OnStartDrag;
    property OnUTF8KeyPress;
    property ParentBidiMode;
    property ParentColor;
    property ParentFont;
    property PopupMenu;
    property ParentShowHint;
    property ReadOnly;
    property ScrollBars;
    property ShowHint;
    property TabOrder;
    property TabStop;
    property Visible;
    property WantReturns;
    property WantTabs;
    property WordWrap;
  end;
  
function GetFontParams(styles: TFontStyles): TFontParams; overload;
function GetFontParams(color: TColor; styles: TFontStyles): TFontParams; overload;
function GetFontParams(const Name: String; color: TColor; styles: TFontStyles): TFontParams; overload;
function GetFontParams(const Name: String; Size: Integer; color: TColor;
  bcolor: TColor; styles: TFontStyles; alignment: TRichAlignment; indented: Boolean;
  changed: TChangedFontItems): TFontParams; overload;

var
  RTFLoadStream : function (AMemo: TCustomRichMemo; Source: TStream): Boolean = nil;
  RTFSaveStream : function (AMemo: TCustomRichMemo; Dest: TStream): Boolean = nil;

implementation

function GetFontParams(styles: TFontStyles): TFontParams; overload;
begin 
  Result := GetFontParams('', 0, 0, 0, styles, [], False, []); // Modified by Massimo Nardello
end;

function GetFontParams(color: TColor; styles: TFontStyles): TFontParams; overload;
begin
  Result := GetFontParams('', 0, color, 0, styles, [], False, []); // Modified by Massimo Nardello
end;

function GetFontParams(const Name: String; color: TColor; styles:
  TFontStyles): TFontParams; overload;
begin
  Result := GetFontParams(Name, 0, color, 0, styles, [], False, []); // Modified by Massimo Nardello
end;

function GetFontParams(const Name: String; Size: Integer; color: TColor;
      bcolor: TColor; styles: TFontStyles; alignment: TRichAlignment;
      indented: Boolean; changed: TChangedFontItems): TFontParams; overload;
begin
  Result.Name := Name;
  Result.Size := Size;
  Result.Color := color;
  Result.BackColor := bcolor; // Added by Massimo Nardello
  Result.Style := styles;
  Result.Alignment := alignment; // Added by Massimo Nardello
  Result.Indented := 0; // Added by Massimo Nardello
  Result.Changed := changed;  // Added by Massimo Nardello
end;

{ TCustomRichMemo }

procedure TCustomRichMemo.SetHideSelection(AValue: Boolean);
begin
  if HandleAllocated then 
    TWSCustomRichMemoClass(WidgetSetClass).SetHideSelection(Self, AValue);
  fHideSelection := AValue;   
end;

class procedure TCustomRichMemo.WSRegisterClass;  
begin
  inherited;
  WSRegisterCustomRichMemo;
end;

procedure TCustomRichMemo.CreateWnd;  
begin
  inherited CreateWnd;  
  UpdateRichMemo;
end;

procedure TCustomRichMemo.UpdateRichMemo; 
begin
  if not HandleAllocated then Exit;
  TWSCustomRichMemoClass(WidgetSetClass).SetHideSelection(Self, fHideSelection);
end;

// Removed by Massimo Nardello
{procedure TCustomRichMemo.SetTextAttributes(TextStart, TextLen: Integer;
  AFont: TFont); 
var
  params  : TFontParams;
begin
  params.Name := AFont.Name;
  params.Color := AFont.Color;
  params.Size := AFont.Size;
  params.Style := AFont.Style;
  SetTextAttributes(TextStart, TextLen, {TextStyleAll,} params);
end;}

procedure TCustomRichMemo.SetTextAttributes(TextStart, TextLen: Integer;  
  {SetMask: TTextStyleMask;} const TextParams: TFontParams);
begin
  if HandleAllocated then  
    TWSCustomRichMemoClass(WidgetSetClass).SetTextAttributes(Self, TextStart,
      TextLen, {SetMask,} TextParams);
end;

procedure TCustomRichMemo.SetDefaultFont(FontName: String; FontSize: Integer);
begin
  // Added by Massimo Nardello
  if HandleAllocated then
    TWSCustomRichMemoClass(WidgetSetClass).SetDefaultFont(Self, FontName, FontSize);
end;

function TCustomRichMemo.GetTextAttributes(TextStart: Integer; var TextParams: TFontParams): Boolean;
begin
  if HandleAllocated then  
    Result := TWSCustomRichMemoClass(WidgetSetClass).GetTextAttributes(Self, TextStart, TextParams)
  else
    Result := false;
end;

function TCustomRichMemo.GetStyleRange(CharOfs: Integer; var RangeStart,
  RangeLen: Integer): Boolean;
begin
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).GetStyleRange(Self, CharOfs, RangeStart, RangeLen);
    if Result and (RangeLen = 0) then RangeLen := 1;
  end else begin
    RangeStart := -1;
    RangeLen := -1;
    Result := false;
  end;
end;

function TCustomRichMemo.GetContStyleLength(TextStart: Integer): Integer;
var
  ofs, len  : Integer;
begin
  if GetStyleRange(TextStart, ofs, len) then Result := len - (TextStart-ofs)
  else Result := 1;
  if Result = 0 then Result := 1;
end;

procedure TCustomRichMemo.SetSelText(const SelTextUTF8: string);  
var
  st  : Integer;
begin
  Lines.BeginUpdate;
  try
    st := SelStart;
    if HandleAllocated then  
      TWSCustomRichMemoClass(WidgetSetClass).InDelText(Self, SelTextUTF8, SelStart, SelLength);
    SelStart := st;
    SelLength := length(UTF8Decode(SelTextUTF8));
  finally
    Lines.EndUpdate;
  end;
end;

procedure TCustomRichMemo.CopyToClipboard;
begin
  if HandleAllocated then  
    TWSCustomRichMemoClass(WidgetSetClass).CopyToClipboard(Self);
end;

procedure TCustomRichMemo.CutToClipboard;
begin
  if HandleAllocated then  
    TWSCustomRichMemoClass(WidgetSetClass).CutToClipboard(Self);
end;

procedure TCustomRichMemo.PasteFromClipboard;
begin
  if HandleAllocated then  
    TWSCustomRichMemoClass(WidgetSetClass).PasteFromClipboard(Self);
end;

procedure TCustomRichMemo.SetRangeColor(TextStart, TextLength: Integer; FontColor: TColor);
begin
  SetRangeParams(TextStart, TextLength, [tmm_Color], '', 0, FontColor, [], []);
end;

procedure TCustomRichMemo.SetRangeParams(TextStart, TextLength: Integer; ModifyMask: TTextModifyMask;
      const FontName: String; FontSize: Integer; FontColor: TColor; AddFontStyle, RemoveFontStyle: TFontStyles);
var
  i : Integer;
  j : Integer;
  l : Integer;
  p : TFontParams;
begin
  if (ModifyMask = []) or (TextLength = 0) then Exit;

  i := TextStart;
  j := TextStart + TextLength;
  while i < j do begin
    GetTextAttributes(i, p);

    if tmm_Name in ModifyMask then p.Name := FontName;
    if tmm_Color in ModifyMask then p.Color := FontColor;
    if tmm_Size in ModifyMask then p.Size := FontSize;
    if tmm_Styles in ModifyMask then p.Style := p.Style + AddFontStyle - RemoveFontStyle;

    l := GetContStyleLength(i);
    if i + l > j then l := j - i;
    if l = 0 then l := 1;
    SetTextAttributes(i, l, p);
    inc(i, l);
  end;
end;


function TCustomRichMemo.LoadRichText(Source: TStream): Boolean;
begin
  if Assigned(Source) and HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).LoadRichText(Self, Source);
    if not Result and Assigned(RTFLoadStream) then begin
      Self.Lines.BeginUpdate;
      Self.Lines.Clear;
      Result:=RTFLoadStream(Self, Source);
      Self.Lines.EndUpdate;
    end;
  end else
    Result := false;
end;

function TCustomRichMemo.SaveRichText(Dest: TStream): Boolean;
begin
  if Assigned(Dest) and HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SaveRichText(Self, Dest);
    if not Result and Assigned(RTFSaveStream) then
      Result:=RTFSaveStream(Self, Dest);
  end else
    Result := false;
end;

function TCustomRichMemo.LoadImageFromFile(TextStart: Integer;
  FileName: String; ImgScale: Single): Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).
    LoadImageFromFile(Self, TextStart, FileName, ImgScale);
  end else
    Result := false;
end;

function TCustomRichMemo.SaveImageToFile(TextStart: Integer; FileName: String): Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SaveImageToFile(Self, TextStart, FileName);
  end else
    Result := false;
end;

function TCustomRichMemo.GetImagePosInText: TStringList;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).GetImagePosInText(Self);
  end else
    Result := nil;
end;

function TCustomRichMemo.ListNumber(TextStart: Integer; Indent: Integer; Bullet: Char): Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).ListNumber(Self, TextStart, Indent, Bullet);
  end else
    Result := false;
end;

function TCustomRichMemo.SetToDo(CurrLine: Integer; Done: Boolean): Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SetToDo(Self, CurrLine, Done);
  end else
    Result := false;
end;

function TCustomRichMemo.InsertText(Start: Integer; TextToInsert: String): Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).InsertText(Self, Start, TextToInsert);
  end else
    Result := false;
end;

function TCustomRichMemo.ClearAll: Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).ClearAll(Self);
  end else
    Result := False;
end;

function TCustomRichMemo.GetWordParagraphStartEnd(CurrPos: Integer; Words: Boolean; Beginning: Boolean): Integer;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).GetWordParagraphStartEnd(Self, CurrPos, Words, Beginning);
  end else
    Result := 0;
end;

function TCustomRichMemo.MoveParUpDown(CurrLine: Integer; MoveUp: Boolean): Boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).MoveParUpDown(Self, CurrLine, MoveUp);
  end else
   Result := False;
end;

function TCustomRichMemo.CountChar: integer;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).CountChar(Self);
  end else
   Result := 0;
end;

function TCustomRichMemo.GetUndo(Start: integer): boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).GetUndo(Self, Start);
  end else
   Result := False;
end;

function TCustomRichMemo.SetUndo(Start: integer): boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SetUndo(Self, Start);
  end else
   Result := False;
end;

function TCustomRichMemo.ClearUndo: boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).ClearUndo(Self);
  end else
   Result := False;
end;

function TCustomRichMemo.SetCursorMiddleScreen(CurrLine: integer): boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SetCursorMiddleScreen(Self, CurrLine);
  end else
   Result := False;
end;

function TCustomRichMemo.SetLineSpace(Space: integer): boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SetLineSpace(Self, Space);
  end else
   Result := False;
end;

function TCustomRichMemo.SetParagraphSpace(Space: integer): boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).SetParagraphSpace(Self, Space);
  end else
   Result := False;
end;

function TCustomRichMemo.CleanParagraph(Start: Integer): boolean;
begin
  // Added by Massimo Nardello
  if HandleAllocated then begin
    Result := TWSCustomRichMemoClass(WidgetSetClass).CleanParagraph(Self, Start);
  end else
   Result := False;
end;


end.

