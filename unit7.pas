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

unit Unit7;

{$mode objfpc}

interface

uses
  Classes, SysUtils, FileUtil, LResources, Forms, Controls, Graphics, Dialogs,
  StdCtrls, ComCtrls, ExtCtrls;

type

  { TfmOptions }

  TfmOptions = class(TForm)
    bnOptFont1: TButton;
    bnOptBack1: TButton;
    bnOptFont2: TButton;
    bnOptBack2: TButton;
    bnOptFont3: TButton;
    bnOptBack3: TButton;
    bnOptionsFormClDef: TButton;
    bvFrame: TBevel;
    bnOptionsFont: TButton;
    bnOptionsSyncDir: TButton;
    bnOptionsFormColor: TButton;
    bnOptionsOK: TButton;
    cbOptionsAutosync: TCheckBox;
    cbOptionsNoMsg: TCheckBox;
    cbOptionsNoChar: TCheckBox;
    cbOptionsOpenLastFile: TCheckBox;
    cbOptionsActivateTray: TCheckBox;
    cbOptionsNoAutosave: TCheckBox;
    edOptParSpace: TEdit;
    edOptLineSpace: TEdit;
    lbOptionsPar: TLabel;
    lbOptionDef1: TLabel;
    lbOptionDef2: TLabel;
    lbOptionDef3: TLabel;
    lbOptionsColor: TLabel;
    lbOptionsFont: TLabel;
    lbOptionsLine: TLabel;
    lbOptionsSyncDir: TLabel;
    lbOptionsFormColor: TLabel;
    lbOptionsFormTrans: TLabel;
    tbOptionsFormTrans: TTrackBar;
    udOptParSpace: TUpDown;
    udOptLineSpace: TUpDown;
    procedure bnOptBack1Click(Sender: TObject);
    procedure bnOptBack2Click(Sender: TObject);
    procedure bnOptBack3Click(Sender: TObject);
    procedure bnOptFont1Click(Sender: TObject);
    procedure bnOptFont2Click(Sender: TObject);
    procedure bnOptFont3Click(Sender: TObject);
    procedure bnOptionsFontClick(Sender: TObject);
    procedure bnOptionsFormClDefClick(Sender: TObject);
    procedure bnOptionsFormColorClick(Sender: TObject);
    procedure bnOptionsOKClick(Sender: TObject);
    procedure bnOptionsSyncDirClick(Sender: TObject);
    procedure cbOptionsActivateTrayChange(Sender: TObject);
    procedure cbOptionsAutosyncChange(Sender: TObject);
    procedure cbOptionsNoAutosaveChange(Sender: TObject);
    procedure cbOptionsNoCharChange(Sender: TObject);
    procedure cbOptionsNoMsgChange(Sender: TObject);
    procedure cbOptionsOpenLastFileChange(Sender: TObject);
    procedure edOptLineSpaceChange(Sender: TObject);
    procedure edOptParSpaceChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: char);
    procedure tbOptionsFormTransChange(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  fmOptions: TfmOptions;

implementation

uses Unit1, Unit2, Unit3, Unit4, Unit5, Unit6, Unit8, Unit9, Unit10, UnitCopyright;

{ TfmOptions }

procedure TfmOptions.FormCreate(Sender: TObject);
begin
  // Set starting color
  fmOptions.Color := fmMain.Color;
  lbOptionDef1.Font.Color := StringToColor(fmMain.stColFont1);
  lbOptionDef1.Color := StringToColor(fmMain.stColBack1);
  lbOptionDef2.Font.Color := StringToColor(fmMain.stColFont2);
  lbOptionDef2.Color := StringToColor(fmMain.stColBack2);
  lbOptionDef3.Font.Color := StringToColor(fmMain.stColFont3);
  lbOptionDef3.Color := StringToColor(fmMain.stColBack3);
end;

procedure TfmOptions.bnOptionsFontClick(Sender: TObject);
begin
  // Change font notes
  with fmMain do
  begin
    fdFontGenDialog.Font.Name := DefFontName;
    fdFontGenDialog.Font.Size := DefFontSize;
    fdFontGenDialog.Font.Style := [];
    if fdFontGenDialog.Execute then
    begin
      // Bold, italics, etc. not allowed for default font
      DefFontName := fdFontGenDialog.Font.Name;
      DefFontSize := fdFontGenDialog.Font.Size;
      dbText.SetDefaultFont(DefFontName, DefFontSize + ZoomFontSize);
      lbOptionsFont.Caption :=
        cpt044 + ' ' + DefFontName + ', ' + IntToStr(DefFontSize) + ' ' + cpt045;
    end;
  end;
end;

procedure TfmOptions.edOptParSpaceChange(Sender: TObject);
begin
  // Set paragraph space
  fmMain.ParagraphSpace := StrToInt(edOptParSpace.Text);
  fmMain.dbText.SetParagraphSpace(fmMain.ParagraphSpace);
end;

procedure TfmOptions.edOptLineSpaceChange(Sender: TObject);
begin
  // Set line space
  fmMain.LineSpace := StrToInt(edOptLineSpace.Text);
  fmMain.dbText.SetLineSpace(fmMain.LineSpace);
  if StrToInt(edOptParSpace.Text) < StrToInt(edOptLineSpace.Text) then
  begin
    edOptParSpace.Text := edOptLineSpace.Text;
    edOptParSpaceChange(nil);
  end;
end;

procedure TfmOptions.bnOptionsFormColorClick(Sender: TObject);
begin
  // Select the forms color
  fmMain.fdColorDialog.Color := fmMain.Color;
  if fmMain.fdColorDialog.Execute then
  begin
    fmMain.Color := fmMain.fdColorDialog.Color;
    if fmMain.fdColorDialog.Color <> clDefault then
    begin
      fmMain.grSubjects.FixedColor := fmMain.Color;
      fmMain.grTitles.FixedColor := fmMain.Color;
      fmMain.grActGrid.FixedColor := fmMain.Color;
      fmMain.grFilter.FixedColor := fmMain.Color;
      fmCalendar.grCalGrid.FixedColor := fmMain.Color;
    end
    else
    begin
      fmMain.grSubjects.FixedColor := clBtnFace;
      fmMain.grTitles.FixedColor := clBtnFace;
      fmMain.grActGrid.FixedColor := clBtnFace;
      fmMain.grFilter.FixedColor := clBtnFace;
      fmCalendar.grCalGrid.FixedColor := clBtnFace;
    end;
    fmMain.dbSubComm.Color := fmMain.fdColorDialog.Color;
    fmMain.meSearchCond.Color := fmMain.fdColorDialog.Color;
    fmMain.lbTagsNames.Color := fmMain.fdColorDialog.Color;
    fmImpExp.Color := fmMain.fdColorDialog.Color;
    fmMoveNote.Color := fmMain.fdColorDialog.Color;
    fmCommentsSubjects.Color := fmMain.fdColorDialog.Color;
    fmEncryption.Color := fmMain.fdColorDialog.Color;
    fmResizeImage.Color := fmMain.fdColorDialog.Color;
    fmLook.Color := fmMain.fdColorDialog.Color;
    fmCalendar.Color := fmMain.fdColorDialog.Color;
    fmSetAlarm.Color := fmMain.fdColorDialog.Color;
    fmOptions.Color := fmMain.fdColorDialog.Color;
    fmOptions.lbOptionsFormColor.Caption :=
      fmMain.cpt047 + ' ' + ColorToString(fmMain.fdColorDialog.Color);
  end;
end;

procedure TfmOptions.bnOptionsFormClDefClick(Sender: TObject);
begin
  // Set default color
  fmMain.Color := clForm;
  fmMain.grSubjects.FixedColor := clBtnFace;
  fmMain.grTitles.FixedColor := clBtnFace;
  fmMain.grActGrid.FixedColor := clBtnFace;
  fmMain.grFilter.FixedColor := clBtnFace;
  fmCalendar.grCalGrid.FixedColor := clBtnFace;
  fmMain.dbSubComm.Color := clForm;
  fmMain.meSearchCond.Color := clForm;
  fmMain.lbTagsNames.Color := clForm;
  fmImpExp.Color := clDefault;
  fmMoveNote.Color := clDefault;
  fmCommentsSubjects.Color := clDefault;
  fmEncryption.Color := clDefault;
  fmResizeImage.Color := clDefault;
  fmLook.Color := clDefault;
  fmCalendar.Color := clDefault;
  fmSetAlarm.Color := clDefault;
  fmOptions.Color := clDefault;
  fmOptions.lbOptionsFormColor.Caption :=
    fmMain.cpt047 + ' clDefault';
end;

procedure TfmOptions.bnOptionsSyncDirClick(Sender: TObject);
begin
  // Change the directory of sync
  with fmMain do
  begin
    if sdSelDirDialog.Execute then
    begin
      SyncFolder := sdSelDirDialog.FileName;
      lbOptionsSyncDir.Caption := fmMain.cpt046 + ' ' + SyncFolder;
      if fmMain.sqSubjects.Active = True then
      begin
        if ((FileExistsUTF8(SyncFolder + DirectorySeparator +
          ExtractFileName(sqNotes.FileName))) and
          (SyncFolder <> ExtractFileDir(sqNotes.FileName))) then
        begin
          miToolsSyncDo.Enabled := True;
          tbToolsSyncDo.Enabled := True;
        end
        else
        begin
          miToolsSyncDo.Enabled := False;
          tbToolsSyncDo.Enabled := False;
        end;
      end;
    end;
  end;
end;

procedure TfmOptions.tbOptionsFormTransChange(Sender: TObject);
begin
  // Set transparency
  fmMain.TranspForm := tbOptionsFormTrans.Position;
  fmMain.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmImpExp.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmMoveNote.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmCommentsSubjects.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmEncryption.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmResizeImage.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmLook.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmCalendar.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmSetAlarm.AlphaBlendValue := tbOptionsFormTrans.Position;
  fmOptions.AlphaBlendValue := tbOptionsFormTrans.Position;
end;

procedure TfmOptions.cbOptionsActivateTrayChange(Sender: TObject);
begin
  // Show and hide the tray icon
  if cbOptionsActivateTray.Checked = True then
  begin
    fmMain.tiTrayIcon.Visible := True;
    fmMain.ShowInTaskBar := stNever;
    fmMain.flTrayIcon := True;
  end
  else
  begin
    fmMain.tiTrayIcon.Visible := False;
    fmMain.ShowInTaskBar := stDefault;
    fmMain.flTrayIcon := False;
  end;
end;

procedure TfmOptions.cbOptionsAutosyncChange(Sender: TObject);
begin
  // Update autosync state
  fmMain.flAutosync := cbOptionsAutosync.Checked;
end;

procedure TfmOptions.cbOptionsNoMsgChange(Sender: TObject);
begin
  // Update no sync message
  fmMain.flNoSyncMsg := cbOptionsNoMsg.Checked;
end;

procedure TfmOptions.cbOptionsNoCharChange(Sender: TObject);
begin
  // Update no char count
  fmMain.flNoCharCount := cbOptionsNoChar.Checked;
  if fmMain.sqNotes.RecordCount > 0 then
  begin
    fmMain.SetCharCount(fmMain.flNoCharCount);
  end;
end;

procedure TfmOptions.cbOptionsNoAutosaveChange(Sender: TObject);
begin
  // Update no autosave
  fmMain.flNoAutosave := cbOptionsNoAutosave.Checked;
end;

procedure TfmOptions.cbOptionsOpenLastFileChange(Sender: TObject);
begin
  // Update open last file state
  fmMain.flOpenLastFile := cbOptionsOpenLastFile.Checked;
end;

procedure TfmOptions.bnOptFont1Click(Sender: TObject);
begin
  // Set font1 color
  fmMain.fdColorDialog.Color := StringToColor(fmMain.stColFont1);
  if fmMain.fdColorDialog.Execute then
    fmMain.stColFont1 := ColorToString(fmMain.fdColorDialog.Color);
  lbOptionDef1.Font.Color := StringToColor(fmMain.stColFont1);
end;

procedure TfmOptions.bnOptFont2Click(Sender: TObject);
begin
  // Set font2 color
  fmMain.fdColorDialog.Color := StringToColor(fmMain.stColFont2);
  if fmMain.fdColorDialog.Execute then
    fmMain.stColFont2 := ColorToString(fmMain.fdColorDialog.Color);
  lbOptionDef2.Font.Color := StringToColor(fmMain.stColFont2);
end;

procedure TfmOptions.bnOptFont3Click(Sender: TObject);
begin
  // Set font3 color
  fmMain.fdColorDialog.Color := StringToColor(fmMain.stColFont3);
  if fmMain.fdColorDialog.Execute then
    fmMain.stColFont3 := ColorToString(fmMain.fdColorDialog.Color);
  lbOptionDef3.Font.Color := StringToColor(fmMain.stColFont3);
end;

procedure TfmOptions.bnOptBack1Click(Sender: TObject);
begin
  // Set back1 color
  fmMain.fdColorDialog.Color := StringToColor(fmMain.stColBack1);
  if fmMain.fdColorDialog.Execute then
    fmMain.stColBack1 := ColorToString(fmMain.fdColorDialog.Color);
  lbOptionDef1.Color := StringToColor(fmMain.stColBack1);
end;

procedure TfmOptions.bnOptBack2Click(Sender: TObject);
begin
  // Set back2 color
  fmMain.fdColorDialog.Color := StringToColor(fmMain.stColBack2);
  if fmMain.fdColorDialog.Execute then
    fmMain.stColBack2 := ColorToString(fmMain.fdColorDialog.Color);
  lbOptionDef2.Color := StringToColor(fmMain.stColBack2);
end;

procedure TfmOptions.bnOptBack3Click(Sender: TObject);
begin
  // Set back3 color
  fmMain.fdColorDialog.Color := StringToColor(fmMain.stColBack3);
  if fmMain.fdColorDialog.Execute then
    fmMain.stColBack3 := ColorToString(fmMain.fdColorDialog.Color);
  lbOptionDef3.Color := StringToColor(fmMain.stColBack3);
end;

procedure TfmOptions.bnOptionsOKClick(Sender: TObject);
begin
  // Close the form
  Close;
end;

procedure TfmOptions.FormKeyPress(Sender: TObject; var Key: char);
begin
  // Close on ESC
  if Key = #27 then
    Close;
end;

initialization
  {$I unit7.lrs}

end.
