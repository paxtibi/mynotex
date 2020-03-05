// ***********************************************************************
// ***********************************************************************
// MyNotex 1.4
// Author and copyright: Massimo Nardello, Modena (Italy) 2010-2017.
// Free software released under GPL licence vers. 3.
//
// In this software is used TDBZVDateTimePicker component
// (http://wiki.freepascal.org/ZVDateTimeControls_Package#TDBZVDateTimePicker),
// a modified version of RichMemo (http://wiki.lazarus.freepascal.org/RichMemo)
// and Dcpcrypt (http://wiki.lazarus.freepascal.org/DCPcrypt).
// The source code of the modified version of RichMemo is available in
// the website of MyNotex.
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version. You can read the version 3
// of the Licence in http://www.gnu.org/licenses/gpl-3.0.txt
// or in the file Licence.txt included in the files of the
// source code of this software.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
// GNU General Public License for more details.
// ***********************************************************************
// ***********************************************************************

program mynotex;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, Unit1, Unit2, Unit3, Unit4, UnitCopyright, sqlite3laz, ZVDateTimeCtrls,
  richmemopackage, printer4lazarus, Unit5,
  Unit6, Unit7, Unit8, Unit9, Unit10, Unit11
  { you can add units after this };

{$R *.res}

begin
  Application.Title:='MyNotex';
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TfmImpExp, fmImpExp);
  Application.CreateForm(TfmMoveNote, fmMoveNote);
  Application.CreateForm(TfmCopyright, fmCopyright);
  Application.CreateForm(TfmCommentsSubjects, fmCommentsSubjects);
  Application.CreateForm(TfmEncryption, fmEncryption);
  Application.CreateForm(TfmResizeImage, fmResizeImage);
  Application.CreateForm(TfmOptions, fmOptions);
  Application.CreateForm(TfmLook, fmLook);
  Application.CreateForm(TfmCalendar, fmCalendar);
  Application.CreateForm(TfmSetAlarm, fmSetAlarm);
  Application.CreateForm(TfmShowAlarm, fmShowAlarm);
  Application.Run;
end.

