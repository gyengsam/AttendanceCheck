//---------------------------------------------------------------------------

// This software is Copyright (c) 2015 Embarcadero Technologies, Inc.
// You may only use this software if you are an authorized licensee
// of an Embarcadero developer tools product.
// This software is considered a Redistributable as defined under
// the software license agreement that comes with the Embarcadero Products
// and is subject to that software license agreement.

//---------------------------------------------------------------------------

program AttendanceCheck;

uses
  Vcl.Forms,
  uSplitView in 'uSplitView.pas' {SplitViewForm},
  Vcl.Themes,
  Vcl.Styles,
  uDisplayMain in 'uDisplayMain.pas' {fmDisplayMain},
  uCommonFunction in 'uCommonFunction.pas',
  uSetUp in 'uSetUp.pas' {fmSetUp},
  uReport in 'uReport.pas' {fmReport},
  uDataModule in 'uDataModule.pas' {DM: TDataModule},
  uPassword in 'uPassword.pas' {fmPassword},
  uLog in 'uLog.pas' {fmLog};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'TSplitView Demo';
  Application.CreateForm(TSplitViewForm, SplitViewForm);
  Application.CreateForm(TfmDisplayMain, fmDisplayMain);
  Application.CreateForm(TfmSetUp, fmSetUp);
  Application.CreateForm(TfmReport, fmReport);
  Application.CreateForm(TDM, DM);
  Application.CreateForm(TfmPassword, fmPassword);
  Application.CreateForm(TfmLog, fmLog);
  Application.Run;
end.
