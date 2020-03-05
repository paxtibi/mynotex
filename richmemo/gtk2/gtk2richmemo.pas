{
 gtk2richmemo.pas

 Author: Dmitry 'skalogryz' Boyarintsev
 Modified by Massimo Nardello

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

unit Gtk2RichMemo;

{$mode objfpc}{$H+}

interface

uses
  // Bindings
  gtk2, glib2, gdk2, pango,
  // FCL
  Classes, SysUtils,
  // LCL
  LCLType, Controls, Graphics,
  // Gtk2 widget
  Gtk2Def,
  GTK2WinApiWindow, Gtk2Globals, Gtk2Proc, InterfaceBase,
  Gtk2WSControls, gdk2pixbuf,
  // RichMemo
  WSRichMemo, Forms;

{ TGtk2WSCustomRichMemo }
type
  TGtk2WSCustomRichMemo = class(TWSCustomRichMemo)
  private
  protected
    class procedure SetCallbacks(const AGtkWidget: PGtkWidget;
      const AWidgetInfo: PWidgetInfo);
  published
    class function CreateHandle(const AWinControl: TWinControl;
      const AParams: TCreateParams): TLCLIntfHandle; override;
    class procedure SetDefaultFont(const AWinControl: TWinControl;
      FontName: string; FontSize: integer); override;
    class procedure SetTextAttributes(const AWinControl: TWinControl;
      TextStart, TextLen: integer; const Params: TIntFontParams); override;
    class function GetTextAttributes(const AWinControl: TWinControl;
      TextStart: integer; var Params: TIntFontParams): boolean; override;
    class procedure CutToClipboard(const AWinControl: TWinControl); override;
    class procedure PasteFromClipboard(const AWinControl: TWinControl); override;
    class procedure CopyToClipboard(const AWinControl: TWinControl); override;
    class function LoadImageFromFile(const AWinControl: TWinControl;
      TextStart: integer; FileName: string; ImgScale: single): boolean; override;
    class function SaveImageToFile(const AWinControl: TWinControl;
      TextStart: integer; FileName: string): boolean; override;
    class function GetImagePosInText(const AWinControl: TWinControl): TStringList;
      override;
    class function ListNumber(const AWinControl: TWinControl;
      TextStart: integer; Indent: integer; Bullet: char): boolean; override;
    class function SetToDo(const AWinControl: TWinControl; CurrLine: integer;
      Done: boolean): boolean; override;
    class function InsertText(const AWinControl: TWinControl;
      Start: integer; TextToInsert: string): boolean; override;
    class function GetWordParagraphStartEnd(const AWinControl: TWinControl;
      CurrPos: integer; Words: boolean; Beginning: boolean): integer; override;
    class function ClearAll(const AWinControl: TWinControl): boolean; override;
    class function MoveParUpDown(const AWinControl: TWinControl;
      CurrLine: integer; MoveUp: boolean): boolean; override;
    class function CountChar(const AWinControl: TWinControl): integer; override;
    class function SetCursorMiddleScreen(const AWinControl: TWinControl; CurrLine: integer): boolean; override;
    class function GetUndo(const AWinControl: TWinControl; Start: integer): boolean; override;
    class function SetUndo(const AWinControl: TWinControl; Start: integer): boolean; override;
    class function ClearUndo(const AWinControl: TWinControl): boolean; override;
    class function SetLineSpace(const AWinControl: TWinControl; Space: integer): boolean; override;
    class function SetParagraphSpace(const AWinControl: TWinControl; Space: integer): boolean; override;
    class function CleanParagraph(const AWinControl: TWinControl; Start: Integer): boolean; override;
  end;

var
  IndValue: integer = 50;
  LeftMargin: Integer = 15;

implementation

function gtktextattr_underline(const a: TGtkTextAppearance): boolean;
begin
  Result := ((a.flag0 and bm_TGtkTextAppearance_underline) shr
    bp_TGtkTextAppearance_underline) > 0;
end;

function gtktextattr_strikethrough(const a: TGtkTextAppearance): boolean;
begin
  Result := ((a.flag0 and bm_TGtkTextAppearance_strikethrough) shr
    bp_TGtkTextAppearance_strikethrough) > 0;
end;

function GtkTextAttrToFontParams(const textAttr: TGtkTextAttributes;
  var FontParams: TIntFontParams): boolean;
var
  w: integer;
  st: TPangoStyle;
  pf: PPangoFontDescription;
  pj: longint; // Added by Massimo Nardello
  pi: longint; // Added by Massimo Nardello
begin
  FontParams.Style := [];
  FontParams.Name := '';
  FontParams.Size := 0;
  FontParams.Color := 0;
  FontParams.BackColor := clWhite; // Added by Massimo Nardello
  FontParams.Alignment := []; // Added by Massimo Nardello
  FontParams.Indented := LeftMargin; // Added by Massimo Nardello
  FontParams.Changed := []; // Added by Massimo Nardello

  // Added by Massimo Nardello...
  pj := textAttr.justification;
  if pj = 0 then
    FontParams.Alignment := [trLeft]
  else if pj = 1 then
    FontParams.Alignment := [trRight]
  else if pj = 2 then
    FontParams.Alignment := [trCenter]
  else if pj = 3 then
    FontParams.Alignment := [trJustified];
  pi := textAttr.left_margin;
  FontParams.Indented := pi;
  // ... up to here

  pf := textAttr.font;
  Result := Assigned(pf);
  if not Result then
    Exit;
  if Assigned(pf) then
  begin
    FontParams.Name := pango_font_description_get_family(pf);
    FontParams.Size := pango_font_description_get_size(pf);
    // Modified by Massimo Nardello...
    if pango_font_description_get_size_is_absolute(pf) then
      // Font size is set by the size property of the TRichMemo component
      FontParams.Size := FontParams.Size * 72 div Screen.PixelsPerInch;
    FontParams.Size := Round(FontParams.Size / PANGO_SCALE);
    // ... up to here
    w := pango_font_description_get_weight(pf);
    if w > PANGO_WEIGHT_NORMAL then
      Include(FontParams.Style, fsBold);

    st := pango_font_description_get_style(pf);
    if st and PANGO_STYLE_ITALIC > 0 then
      Include(FontParams.Style, fsItalic);
  end;
  FontParams.Color := TGDKColorToTColor(textAttr.appearance.fg_color);
  FontParams.BackColor := TGDKColorToTColor(textAttr.appearance.bg_color);
  // Added by Massimo Nardello
  if gtktextattr_underline(textAttr.appearance) then
    Include(FontParams.Style, fsUnderline);
  if gtktextattr_strikethrough(textAttr.appearance) then
    Include(FontParams.Style, fsStrikeOut);
end;

class procedure TGtk2WSCustomRichMemo.SetCallbacks(const AGtkWidget: PGtkWidget;
  const AWidgetInfo: PWidgetInfo);
begin
  TGtk2WSWinControl.SetCallbacks(PGtkObject(AGtkWidget),
    TComponent(AWidgetInfo^.LCLObject));
end;

class function TGtk2WSCustomRichMemo.CreateHandle(const AWinControl: TWinControl;
  const AParams: TCreateParams): TLCLIntfHandle;
var
  Widget, TempWidget: PGtkWidget;
  WidgetInfo: PWidgetInfo;
  tabArray: PPangoTabArray;
begin
  Widget := gtk_scrolled_window_new(nil, nil);
  Result := TLCLIntfHandle(PtrUInt(Widget));
  if Result = 0 then
    Exit;
  {$IFDEF DebugLCLComponents}
  DebugGtkWidgets.MarkCreated(Widget, dbgsName(AWinControl));
  {$ENDIF}
  WidgetInfo := CreateWidgetInfo(Pointer(Result), AWinControl, AParams);
  TempWidget := gtk_text_view_new();
  gtk_container_add(PGtkContainer(Widget), TempWidget);
  GTK_WIDGET_UNSET_FLAGS(PGtkScrolledWindow(Widget)^.hscrollbar, GTK_CAN_FOCUS);
  GTK_WIDGET_UNSET_FLAGS(PGtkScrolledWindow(Widget)^.vscrollbar, GTK_CAN_FOCUS);
  gtk_scrolled_window_set_policy(PGtkScrolledWindow(Widget),
    GTK_POLICY_AUTOMATIC,
    GTK_POLICY_AUTOMATIC);
  // add border for memo
  gtk_scrolled_window_set_shadow_type(PGtkScrolledWindow(Widget),
    BorderStyleShadowMap[TCustomControl(AWinControl).BorderStyle]);
  SetMainWidget(Widget, TempWidget);
  GetWidgetInfo(Widget, True)^.CoreWidget := TempWidget;
  gtk_text_view_set_editable(PGtkTextView(TempWidget), True);
  gtk_text_view_set_wrap_mode(PGtkTextView(TempWidget), GTK_WRAP_WORD);

  gtk_text_view_set_accepts_tab(PGtkTextView(TempWidget), True);
  // Added by Massimo Nardello...
  gtk_text_view_set_left_margin (PGtkTextView(TempWidget), LeftMargin);
  gtk_text_view_set_right_margin (PGtkTextView(TempWidget), LeftMargin);
  tabArray := pango_tab_array_new_with_positions(7, True, PANGO_TAB_LEFT,
    50, PANGO_TAB_LEFT, 100, PANGO_TAB_LEFT, 150, PANGO_TAB_LEFT,
    200, PANGO_TAB_LEFT, 250, PANGO_TAB_LEFT, 300, PANGO_TAB_LEFT, 350);
  gtk_text_view_set_tabs(PGtkTextView(TempWidget), tabArray);
  // ..up to here
  gtk_widget_show_all(Widget);
  Set_RC_Name(AWinControl, Widget);
  SetCallbacks(Widget, WidgetInfo);
end;

class procedure TGtk2WSCustomRichMemo.SetDefaultFont(const AWinControl: TWinControl;
  FontName: string; FontSize: integer);
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  font_desc: TPangoFontDescription;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  font_desc :=
    pango_font_description_from_string(PChar(FontName + ' ' + IntToStr(FontSize)));
  gtk_widget_modify_font(TextWidget, font_desc);
  pango_font_description_free(font_desc);
end;

class procedure TGtk2WSCustomRichMemo.CutToClipboard(const AWinControl: TWinControl);
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  clipboard: PGtkClipboard;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  clipboard := gtk_clipboard_get(GDK_SELECTION_CLIPBOARD);
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_cut_clipboard(buffer, clipboard, True);
end;

class procedure TGtk2WSCustomRichMemo.CopyToClipboard(const AWinControl: TWinControl);
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  clipboard: PGtkClipboard;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  clipboard := gtk_clipboard_get(GDK_SELECTION_CLIPBOARD);
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_copy_clipboard(buffer, clipboard);
end;

class procedure TGtk2WSCustomRichMemo.PasteFromClipboard(const AWinControl: TWinControl);
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  clipboard: PGtkClipboard;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  clipboard := gtk_clipboard_get(GDK_SELECTION_CLIPBOARD);
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_paste_clipboard(buffer, clipboard, NULL, True);
end;

class procedure TGtk2WSCustomRichMemo.SetTextAttributes(const AWinControl: TWinControl;
  TextStart, TextLen: integer; const Params: TIntFontParams);
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  tag: Pointer;
  istart: TGtkTextIter;
  iend: TGtkTextIter;
  gcolor: TGdkColor;
  bcolor: TGdkColor;
  nm: string;
const
  PangoUnderline: array [boolean] of integer =
    (PANGO_UNDERLINE_NONE, PANGO_UNDERLINE_SINGLE);
  PangoBold: array [boolean] of integer = (PANGO_WEIGHT_NORMAL, PANGO_WEIGHT_BOLD);
  PangoItalic: array [boolean] of integer = (PANGO_STYLE_NORMAL, PANGO_STYLE_ITALIC);
begin
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  if not Assigned(buffer) then
    Exit;
  gcolor := TColortoTGDKColor(Params.Color);
  bcolor := TColortoTGDKColor(Params.BackColor);
  nm := Params.Name;
  if nm = '' then
    nm := #0;
  // Code modified by Massimo Nardello
  // Font name is modified
  gtk_text_buffer_get_iter_at_offset(buffer, @istart, TextStart);
  gtk_text_buffer_get_iter_at_offset(buffer, @iend, TextStart + TextLen);
  if fiName in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'family-set',
      [gboolean(gTRUE), 'family', @nm[1], nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font size is modified
  if fiSize in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'size-set',
      [gboolean(gTRUE), 'size-points', double(Params.Size), nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font color is modified
  if fiColor in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'foreground-gdk',
      [@gcolor, nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font back color is modified
  if fiBackcolor in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'background-gdk',
      [@bcolor, nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font style underline is activated or deactivated
  if fiUnderline in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'underline-set',
      [gboolean(gTRUE), 'underline', PangoUnderline[fsUnderline in
      Params.Style], nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font style bold is activated or deactivated
  if fiBold in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'weight-set',
      [gboolean(gTRUE), 'weight', PangoBold[fsBold in Params.Style], nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font style italic is activated or deactivated
  if fiItalic in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'style-set',
      [gboolean(gTRUE), 'style', PangoItalic[fsItalic in Params.Style], nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Font style strike is activated or deactivated
  if fiStrike in Params.Changed then
  begin
    tag := gtk_text_buffer_create_tag(buffer, nil, 'strikethrough-set',
      [gboolean(gTRUE), 'strikethrough', gboolean(fsStrikeOut in
      Params.Style), nil]);
    gtk_text_buffer_apply_tag(buffer, tag, @istart, @iend);
  end;
  // Alignment
  if fiAlignment in Params.Changed then
  begin
    // Remove possibile previous paragraph formatting
    gtk_text_buffer_remove_tag_by_name(buffer, 'center', @istart, @iend);
    gtk_text_buffer_remove_tag_by_name(buffer, 'left', @istart, @iend);
    gtk_text_buffer_remove_tag_by_name(buffer, 'right', @istart, @iend);
    gtk_text_buffer_remove_tag_by_name(buffer, 'justified', @istart, @iend);
    if Params.Alignment = [trCenter] then
    begin
      gtk_text_buffer_create_tag(buffer, 'center', 'justification',
        [GTK_JUSTIFY_CENTER, nil]);
      gtk_text_buffer_apply_tag_by_name(buffer, 'center', @istart, @iend);
    end
    else if Params.Alignment = [trLeft] then
    begin
      gtk_text_buffer_create_tag(buffer, 'left', 'justification',
        [GTK_JUSTIFY_LEFT, nil]);
      gtk_text_buffer_apply_tag_by_name(buffer, 'left', @istart, @iend);
    end
    else if Params.Alignment = [trRight] then
    begin
      gtk_text_buffer_create_tag(buffer, 'right', 'justification',
        [GTK_JUSTIFY_RIGHT, nil]);
      gtk_text_buffer_apply_tag_by_name(buffer, 'right', @istart, @iend);
    end
    else if Params.Alignment = [trJustified] then
    begin
      gtk_text_buffer_create_tag(buffer, 'justified', 'justification',
        [GTK_JUSTIFY_FILL, nil]);
      gtk_text_buffer_apply_tag_by_name(buffer, 'justified', @istart, @iend);
    end;
  end;
  if fiIndented in Params.Changed then
  begin
    //  Indent the text
    //  Alignment perfect with a Bullet and a tab in Sans 12
    if Params.Indented > LeftMargin then
    begin
      gtk_text_buffer_create_tag(buffer, 'wide_margins',
        'left_margin', [IndValue, 'indent', -Params.Indented, nil]);
      gtk_text_buffer_apply_tag_by_name(buffer, 'wide_margins', @istart, @iend);
    end
    else
    begin
      gtk_text_buffer_remove_tag_by_name(buffer, 'wide_margins', @istart, @iend);
    end;
  end;
end;


class function TGtk2WSCustomRichMemo.GetTextAttributes(const AWinControl: TWinControl;
  TextStart: integer; var Params: TIntFontParams): boolean;
var
  Widget: PGtkWidget;
  TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  iter: TGtkTextIter;
  attr: PGtkTextAttributes;
begin
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));

  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  if not Assigned(buffer) then
    Exit;
  attr := gtk_text_view_get_default_attributes(PGtkTextView(TextWidget));
  Result := Assigned(attr);
  if not Assigned(attr) then
    Exit;
  gtk_text_buffer_get_iter_at_offset(buffer, @iter, TextStart);
  Result := gtk_text_iter_get_attributes(@iter, attr);
  GtkTextAttrToFontParams(attr^, Params);
  gtk_text_attributes_unref(attr);
end;

class function TGtk2WSCustomRichMemo.GetImagePosInText(
  const AWinControl: TWinControl): TStringList;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart: TGtkTextIter;
  iend: TGtkTextIter;
  myText: WideString;
  i: integer;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  Result := TStringList.Create; // It *must* be destroyed in the application code!
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_start_iter(buffer, @istart);
  gtk_text_buffer_get_end_iter(buffer, @iend);
  myText := gtk_text_iter_get_slice(@istart, @iend);
  for i := 1 to Length(myText) do
    if Ord(myText[i]) = 65532 then
      Result.Add(IntToStr(i));
end;

class function TGtk2WSCustomRichMemo.LoadImageFromFile(const AWinControl: TWinControl;
  TextStart: integer; FileName: string; ImgScale: single): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart: TGtkTextIter;
  pixbuf: PGdkPixbuf;
  error: PPGError = NULL; //Initialization is necessary
  ImgHeight, ImgWidth: integer;
begin
  // Added by Massimo Nardello
  Result := True;
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_iter_at_offset(buffer, @istart, TextStart);
  pixbuf := gdk_pixbuf_new_from_file(PChar(FileName), error);
  ImgWidth := gdk_pixbuf_get_width(pixbuf);
  ImgHeight := gdk_pixbuf_get_height(pixbuf);
  try
    pixbuf := gdk_pixbuf_new_from_file_at_size(PChar(FileName),
      Trunc(ImgWidth * ImgScale), Trunc(ImgHeight * ImgScale), error);
    gtk_text_buffer_insert_pixbuf(buffer, @istart, pixbuf);
  except;
    Result := False;
  end;
end;

class function TGtk2WSCustomRichMemo.SaveImageToFile(const AWinControl: TWinControl;
  TextStart: integer; FileName: string): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart: TGtkTextIter;
  error: PPGError = NULL; //Initialization is necessary
  pixbuf: TGdkPixBufBuffer;
begin
  // Added by Massimo Nardello
  Result := True;
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_iter_at_offset(buffer, @istart, TextStart);
  pixbuf := gtk_text_iter_get_pixbuf(@istart);
  try
    if Assigned(pixbuf) then
      gdk_pixbuf_savev(pixbuf, PChar(FileName), 'jpeg', nil, nil, error);
  except
    Result := False;
  end;
end;

class function TGtk2WSCustomRichMemo.ListNumber(const AWinControl: TWinControl;
  TextStart: integer; Indent: integer; Bullet: char): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  attr: PGtkTextAttributes;
  iStart, iBeginning, iEndNum: TGtkTextIter;
  i, iNumProgr: integer;
  IsList, IsTab: boolean;
  stNum: string;
begin
  // Added by Massimo Nardello
  Result := True;
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  if not Assigned(buffer) then
    Exit;
  attr := gtk_text_view_get_default_attributes(PGtkTextView(TextWidget));
  if not Assigned(attr) then
    Exit;
  gtk_text_buffer_get_iter_at_offset(buffer, @iStart, TextStart);
  // Set bullet kind
  if Bullet = '1' then
    iNumProgr := 1
  else if Bullet <> '*' then
    iNumProgr := Ord(Bullet);
  // Go to the first character after two CR
  IsList := True;
  while IsList = True do
  begin
    gtk_text_iter_backward_char(@iStart);
    if gtk_text_iter_is_start(@iStart) = True then
      IsList := False
    else if gtk_text_iter_get_char(@iStart) = Ord(LineEnding) then
    begin
      gtk_text_iter_backward_char(@iStart);
      if ((gtk_text_iter_is_start(@iStart) = True) or
        (gtk_text_iter_get_char(@iStart) = Ord(LineEnding))) then
      begin
        IsList := False;
        gtk_text_iter_forward_char(@iStart);
        gtk_text_iter_forward_char(@iStart);
      end;
    end;
  end;
  iBeginning := iStart;
  // Go down
  IsList := True;
  while IsList = True do
  begin
    if gtk_text_iter_is_end(@iStart) = True then
    begin
      IsList := False;
      Exit;
    end
    else if ((gtk_text_iter_equal(@iStart, @iBeginning)) or
      (gtk_text_iter_get_char(@iStart) = Ord(LineEnding))) then
    begin
      if gtk_text_iter_get_char(@iStart) = Ord(LineEnding) then
      begin
        gtk_text_iter_forward_char(@iStart);
        if ((gtk_text_iter_get_char(@iStart) = Ord(LineEnding)) or
          (gtk_text_iter_is_end(@iStart) = True)) then
        begin
          IsList := False;
          Exit;
        end;
      end;
      iEndNum := iStart;
      isTab := False;
      for i := 0 to 6 do
      begin
        gtk_text_iter_forward_char(@iEndNum);
        if gtk_text_iter_get_char(@iEndNum) = Ord(#9) then
          IsTab := True;
        if ((gtk_text_iter_get_char(@iEndNum) = Ord(#9)) or
          (gtk_text_iter_get_char(@iEndNum) = Ord(LineEnding)) or
          (gtk_text_iter_is_end(@iEndNum) = True)) then
          Break;
      end;
      if IsTab = True then
      begin
        gtk_text_iter_forward_char(@iEndNum);
        gtk_text_buffer_delete(buffer, @iStart, @iEndNum);
      end;
      if Bullet <> 'X' then begin
        if Bullet = '*' then
          gtk_text_buffer_insert(buffer, @iStart, PChar('•' + #9), -1)
        else
        begin
          if Bullet = '1' then
            stNum := IntToStr(iNumProgr)
          else if ((Bullet = 'A') or (Bullet = 'a')) then
          begin
            if ((iNumProgr >= 91) and (iNumProgr <= 96)) then
              iNumProgr := 90
            else if iNumProgr >= 123 then
              iNumProgr := 122;
            stNum := char(iNumProgr);
          end;
          gtk_text_buffer_insert(buffer, @iStart,
            PChar(stNum + '.' + #9), -1);
          iNumProgr := iNumProgr + 1;
        end;
      end;
      while ((gtk_text_iter_starts_line(@iStart) = False) and
          (gtk_text_iter_is_start(@iStart) = False)) do
        gtk_text_iter_backward_char(@iStart);
      iEndNum := iStart;
      while ((gtk_text_iter_ends_line(@iEndNum) = False) and
          (gtk_text_iter_is_end(@iEndNum) = False)) do
        gtk_text_iter_forward_char(@iEndNum);
      if gtk_text_iter_get_slice(@iStart, @iEndNum) = '' then
      begin
        gtk_text_buffer_insert(buffer, @iStart, PChar(' '), -1);
        gtk_text_iter_backward_char(@iStart);
      end;
      if Bullet = 'X' then
        gtk_text_buffer_remove_tag_by_name(buffer, 
          'wide_margins', @iStart, @iEndNum)
      else begin
        gtk_text_buffer_create_tag(buffer, 'wide_margins',
          'left_margin', [IndValue, 'indent', -Indent, nil]);
        gtk_text_buffer_apply_tag_by_name(buffer, 
          'wide_margins', @iStart, @iEndNum);
      end;
    end;
    gtk_text_iter_forward_char(@iStart);
  end;
  gtk_text_attributes_unref(attr);
end;

class function TGtk2WSCustomRichMemo.SetToDo(const AWinControl: TWinControl;
  CurrLine: integer; Done: boolean): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart, iend: TGtkTextIter;
begin
  // Added by Massimo Nardello
  Result := True;
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_iter_at_line(buffer, @istart, CurrLine);
  iend := istart;
  gtk_text_iter_forward_char(@iend);
  if Done = True then
  begin
    if gtk_text_iter_get_slice(@istart, @iend) <> '' then
    begin
      if gtk_text_iter_get_slice(@istart, @iend) = '' then
        gtk_text_buffer_delete(buffer, @istart, @iend);
      gtk_text_buffer_insert(buffer, @istart, PChar(''), -1);
      iend := istart;
      gtk_text_iter_forward_char(@iend);
      if gtk_text_iter_get_slice(@istart, @iend) <> #9 then
        gtk_text_buffer_insert(buffer, @istart, PChar(#9), -1);
    end;
  end
  else
  begin
    if gtk_text_iter_get_slice(@istart, @iend) <> '' then
    begin
      if gtk_text_iter_get_slice(@istart, @iend) = '' then
        gtk_text_buffer_delete(buffer, @istart, @iend);
      gtk_text_buffer_insert(buffer, @istart, PChar(''), -1);
      iend := istart;
      gtk_text_iter_forward_char(@iend);
      if gtk_text_iter_get_slice(@istart, @iend) <> #9 then
        gtk_text_buffer_insert(buffer, @istart, PChar(#9), -1);
    end;
  end;
end;

class function TGtk2WSCustomRichMemo.InsertText(const AWinControl: TWinControl;
  Start: integer; TextToInsert: string): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart: TGtkTextIter;
begin
  // Added by Massimo Nardello
  Result := True;
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_iter_at_offset(buffer, @istart, Start);
  gtk_text_buffer_insert(buffer, @istart, PChar(TextToInsert), -1);
end;

class function TGtk2WSCustomRichMemo.GetWordParagraphStartEnd(
  const AWinControl: TWinControl; CurrPos: integer; Words: boolean;
  Beginning: boolean): integer;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  ipos: TGtkTextIter;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_iter_at_offset(buffer, @ipos, CurrPos);
  // Beginning of word
  if ((Words = True) and (Beginning = True)) then
  begin
    while ((gtk_text_iter_starts_word(@iPos) = False) and
        (gtk_text_iter_is_start(@iPos) = False) and
        (gtk_text_iter_starts_line(@iPos) = False)) do
      gtk_text_iter_backward_char(@iPos);
  end;
  // End of word
  if ((Words = True) and (Beginning = False)) then
  begin
    while ((gtk_text_iter_ends_word(@iPos) = False) and
        (gtk_text_iter_is_end(@iPos) = False) and
        (gtk_text_iter_ends_line(@iPos) = False)) do
      gtk_text_iter_forward_char(@iPos);
  end;
  // Beginning of paragraph
  if ((Words = False) and (Beginning = True)) then
  begin
    while ((gtk_text_iter_starts_line(@iPos) = False) and
        (gtk_text_iter_is_start(@iPos) = False)) do
      gtk_text_iter_backward_char(@iPos);
  end;
  // End of paragraph
  if ((Words = False) and (Beginning = False)) then
  begin
    while ((gtk_text_iter_ends_line(@iPos) = False) and
        (gtk_text_iter_is_end(@iPos) = False)) do
      gtk_text_iter_forward_char(@iPos);
  end;
  Result := gtk_text_iter_get_offset(@iPos);
end;

class function TGtk2WSCustomRichMemo.ClearAll(const AWinControl: TWinControl): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart, iend: TGtkTextIter;
begin
  // Added by Massimo Nardello
  // Useful to clear all, also the pictures
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_start_iter(buffer, @istart);
  gtk_text_buffer_get_end_iter(buffer, @iend);
  gtk_text_buffer_delete(buffer, @istart, @iend);
end;

class function TGtk2WSCustomRichMemo.MoveParUpDown(const AWinControl: TWinControl;
  CurrLine: integer; MoveUp: boolean): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart, iend, iDest: TGtkTextIter;
  mypos: PGtkTextMark;
begin
  // Added by Massimo Nardello
  // Move up and down paragraphs
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));

  gtk_text_buffer_get_iter_at_line(buffer, @iStart, CurrLine);
  if ((gtk_text_iter_is_start(@iStart) = True) and (MoveUp = True)) then
    Exit
  else if ((gtk_text_buffer_get_line_count(buffer) =
    gtk_text_iter_get_line(@iStart)) and (MoveUp = False)) then
    Exit;
  iEnd := iStart;
  while ((gtk_text_iter_ends_line(@iEnd) = False) and
      (gtk_text_iter_is_end(@iEnd) = False)) do
    gtk_text_iter_forward_char(@iEnd);
  gtk_text_iter_forward_char(@iEnd);

  if MoveUp = True then
  begin
    iDest := iStart;
    gtk_text_iter_backward_line(@iDest)
  end
  else
  begin
    iDest := iEnd;
    gtk_text_iter_forward_line(@iDest);
  end;
  gtk_text_buffer_insert_range(buffer, @iDest, @iStart, @iEnd);
  // Select again the text to be deleted
  if MoveUp = True then
  begin
    CurrLine := CurrLine + 1;
  end;
  gtk_text_buffer_get_iter_at_line(buffer, @iStart, CurrLine);
  if ((gtk_text_iter_is_start(@iStart) = True) and (MoveUp = True)) then
    Exit
  else if ((gtk_text_buffer_get_line_count(buffer) =
    gtk_text_iter_get_line(@iStart)) and (MoveUp = False)) then
    Exit;
  iEnd := iStart;
  while ((gtk_text_iter_ends_line(@iEnd) = False) and
      (gtk_text_iter_is_end(@iEnd) = False)) do
    gtk_text_iter_forward_char(@iEnd);
  gtk_text_iter_forward_char(@iEnd);
  gtk_text_buffer_delete(buffer, @iStart, @iEnd);
  
  gtk_text_buffer_get_iter_at_line(buffer, @iDest, CurrLine);
  mypos := gtk_text_buffer_create_mark(buffer, PChar('mypos'), @iDest, False);
  gtk_text_view_scroll_to_mark(PGtkTextView(TextWidget), mypos, 1, True, 0.0, 0.5);
  gtk_text_buffer_delete_mark(buffer, mypos);
end;

class function TGtk2WSCustomRichMemo.CountChar(const AWinControl: TWinControl): integer;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
begin
  // Added by Massimo Nardello
  // Get number of characters
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  Result := gtk_text_buffer_get_char_count(buffer);
end;

class function TGtk2WSCustomRichMemo.SetCursorMiddleScreen(
  const AWinControl: TWinControl; CurrLine: integer): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  iDest: TGtkTextIter;
  mypos: PGtkTextMark;
begin
  // Added by Massimo Nardello
  // Set the current line at the middle of the screen
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_get_iter_at_line(buffer, @iDest, CurrLine);
  mypos := gtk_text_buffer_create_mark(buffer, PChar('mypos'), @iDest, True);
  gtk_text_view_scroll_to_mark(PGtkTextView(TextWidget), mypos, 0, True, 0, 0.5);
  gtk_text_buffer_delete_mark(buffer, mypos);
end;

class function TGtk2WSCustomRichMemo.SetUndo(
  const AWinControl: TWinControl; Start: integer): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart, iend: TGtkTextIter;
  clipboard: PGtkClipboard;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  clipboard := gtk_clipboard_get(GDK_SELECTION_SECONDARY);
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_begin_user_action(buffer);
  gtk_text_buffer_get_start_iter(buffer, @istart);
  gtk_text_buffer_get_end_iter(buffer, @iend);
  gtk_text_buffer_delete(buffer, @istart, @iend);
  gtk_text_buffer_paste_clipboard(buffer, clipboard, NULL, True);
  gtk_text_buffer_get_iter_at_offset(buffer, @istart, Start);  
  gtk_text_buffer_select_range(buffer, @istart, @istart);
  gtk_text_buffer_end_user_action(buffer);
end;

class function TGtk2WSCustomRichMemo.GetUndo(
  const AWinControl: TWinControl; Start: integer): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  istart, iend: TGtkTextIter;
  clipboard: PGtkClipboard;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  clipboard := gtk_clipboard_get(GDK_SELECTION_SECONDARY);
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  gtk_text_buffer_begin_user_action(buffer);
  gtk_text_buffer_get_start_iter(buffer, @istart);
  gtk_text_buffer_get_end_iter(buffer, @iend);
  gtk_text_buffer_select_range(buffer, @istart, @iend);
  gtk_text_buffer_copy_clipboard(buffer, clipboard);
  gtk_text_buffer_get_iter_at_offset(buffer, @istart, Start);  
  gtk_text_buffer_select_range(buffer, @istart, @istart);
  gtk_text_buffer_end_user_action(buffer);
end;

class function TGtk2WSCustomRichMemo.ClearUndo(
  const AWinControl: TWinControl): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  clipboard: PGtkClipboard;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  clipboard := gtk_clipboard_get(GDK_SELECTION_SECONDARY);
  gtk_clipboard_clear(clipboard);
  gtk_clipboard_set_text(clipboard, '', 0);
end;

class function TGtk2WSCustomRichMemo.SetLineSpace(
  const AWinControl: TWinControl; Space: Integer): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  gtk_text_view_set_pixels_inside_wrap(PGtkTextView(TextWidget), Space);
end;

class function TGtk2WSCustomRichMemo.SetParagraphSpace(
  const AWinControl: TWinControl; Space: Integer): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
begin
  // Added by Massimo Nardello
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  gtk_text_view_set_pixels_below_lines(PGtkTextView(TextWidget), Space);
end;

class function TGtk2WSCustomRichMemo.CleanParagraph(
  const AWinControl: TWinControl; Start: Integer): boolean;
var
  Widget, TextWidget: PGtkWidget;
  list: PGList;
  buffer: PGtkTextBuffer;
  iStart, iEnd: TGtkTextIter;
  IsList: boolean;
begin
  // Added by Massimo Nardello
  Result := True;
  Widget := PGtkWidget(PtrUInt(AWinControl.Handle));
  list := gtk_container_get_children(PGtkContainer(Widget));
  if not Assigned(list) then
    Exit;
  TextWidget := PGtkWidget(list^.Data);
  if not Assigned(TextWidget) then
    Exit;
  buffer := gtk_text_view_get_buffer(PGtkTextView(TextWidget));
  if not Assigned(buffer) then
    Exit;
  gtk_text_buffer_get_iter_at_offset(buffer, @iStart, Start);
  // Go to the first character after two CR
  IsList := True;
  while IsList = True do
  begin
    gtk_text_iter_backward_char(@iStart);
    if gtk_text_iter_is_start(@iStart) = True then
      IsList := False
    else if gtk_text_iter_get_char(@iStart) = Ord(LineEnding) then
    begin
      gtk_text_iter_backward_char(@iStart);
      if ((gtk_text_iter_is_start(@iStart) = True) or
        (gtk_text_iter_get_char(@iStart) = Ord(LineEnding))) then
      begin
        IsList := False;
        gtk_text_iter_forward_char(@iStart);
        gtk_text_iter_forward_char(@iStart);
      end;
    end;
  end;
  // Go down
  IsList := True;
  while IsList = True do
  begin
    if gtk_text_iter_is_end(@iStart) = True then
    begin
      IsList := False;
      Exit;
    end
    else if gtk_text_iter_get_char(@iStart) = Ord(LineEnding) then
    begin
      gtk_text_iter_forward_char(@iStart);
      if ((gtk_text_iter_get_char(@iStart) = Ord(LineEnding)) or
        (gtk_text_iter_is_end(@iStart) = True)) then
      begin
        IsList := False;
        Exit;
      end
      else
      begin
        iEnd := iStart;
        gtk_text_iter_backward_char(@iStart);
        gtk_text_buffer_delete(buffer, @iStart, @iEnd);
        gtk_text_buffer_insert(buffer, @iStart, PChar(' '), -1);
      end;
    end;
    gtk_text_iter_forward_char(@iStart);
  end;
end;

end.
