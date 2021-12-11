//---------------------------------------------------------------------------

// This software is Copyright (c) 2015 Embarcadero Technologies, Inc.
// You may only use this software if you are an authorized licensee
// of an Embarcadero developer tools product.
// This software is considered a Redistributable as defined under
// the software license agreement that comes with the Embarcadero Products
// and is subject to that software license agreement.

//---------------------------------------------------------------------------

unit uSplitView;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  System.ImageList,
  System.Actions,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.ExtCtrls,
  Vcl.WinXCtrls,
  Vcl.StdCtrls,
  Vcl.CategoryButtons,
  Vcl.Buttons,
  Vcl.ImgList,
  Vcl.Imaging.PngImage,
  Vcl.ComCtrls,
  Vcl.ActnList,
  System.IniFiles,
  OtlComm,
  OtlTaskControl,
  OtlParallel, JvExExtCtrls, JvImage, Vcl.Menus;//, JvExExtCtrls, JvImage;

const
  WM_MSG_DISPLAY_UPDATE = WM_USER + 100;

type
  TSplitViewForm = class(TForm)
    pnlToolbar: TPanel;
    pnlSettings: TPanel;
    SV: TSplitView;
    catMenuItems: TCategoryButtons;
    imlIcons: TImageList;
    imgMenu: TImage;
    ActionList1: TActionList;
    actHome: TAction;
    actLayout: TAction;
    actPower: TAction;
    pnServer: TPanel;
    pnSerial: TPanel;
    JvImage1: TJvImage;
    lblTitle: TLabel;
    lbCurrentNo: TLabel;
    JvImage2: TJvImage;
    pnSqLite: TPanel;
    pnShowLogForm: TPanel;
    pnExt: TPanel;
    popFile: TPopupMenu;
    pmConnectionToServer: TMenuItem;
    N4: TMenuItem;
    pmGetLocalCount: TMenuItem;
    pmGetMemberInfoFromServer: TMenuItem;
    N3: TMenuItem;
    pmDrom: TMenuItem;
    N1: TMenuItem;
    LocalPcToServerPc: TMenuItem;
    ServerPcToLocalPc: TMenuItem;
    LocalPcToServerPcMember: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure SVClosed(Sender: TObject);
    procedure SVClosing(Sender: TObject);
    procedure SVOpened(Sender: TObject);
    procedure SVOpening(Sender: TObject);
    procedure catMenuItemsCategoryCollapase(Sender: TObject; const Category: TButtonCategory);
    procedure imgMenuClick(Sender: TObject);
    procedure actHomeExecute(Sender: TObject);
    procedure actLayoutExecute(Sender: TObject);
    procedure actPowerExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure UpDateInspStatusIcon;
    procedure FormDestroy(Sender: TObject);
    procedure pnShowLogFormClick(Sender: TObject);
    procedure pmConnectionToServerClick(Sender: TObject);
    procedure pmGetLocalCountClick(Sender: TObject);
    procedure pmDromClick(Sender: TObject);
    procedure pmGetMemberInfoFromServerClick(Sender: TObject);
    procedure LocalPcToServerPcClick(Sender: TObject);
    procedure ServerPcToLocalPcClick(Sender: TObject);
    procedure LocalPcToServerPcMemberClick(Sender: TObject);
  private
    procedure wmUserMessage_DisplayUpdate(var Msg: TMessage); message WM_MSG_DISPLAY_UPDATE;
  public
  end;

var
  SplitViewForm: TSplitViewForm;

implementation
uses
  uDataModule, uDisplayMain, uSetUp, uReport, uCommonFunction, uPassword,
  uLog, XSuperObject;//Vcl.Themes,

{$R *.dfm}

procedure TSplitViewForm.FormCreate(Sender: TObject);
var
  StyleName: string;
begin
  SystemOption := TSystemOption.create;
end;

procedure TSplitViewForm.FormDestroy(Sender: TObject);
begin
  SystemOption.Free;
end;

procedure TSplitViewForm.FormShow(Sender: TObject);
var
  bComport : Boolean;
  i: Integer;
  sFileName, SystemOptionFile, LogData, CpuId : string;
  tList : TStringList;
  obj : ISuperObject;
begin
  SplitViewForm.Caption := '미사참례등록 : Version 3.0';
  mHandle := Self.Handle;

  cDir := GetCurrentDir;
  Sql3DbName := cDir + '\SQLite3\Data.sdb';

  SV.DisplayMode := TSplitViewDisplayMode(0);//0=Docked, 1=Overlay
  SV.CloseStyle := TSplitViewCloseStyle(1);//0=Collapse, Comact
  SV.Placement := TSplitViewPlacement(0);//0=Left, 1=Right

  SV.UseAnimation := true;

  SV.AnimationDelay := 15;//trkAnimationDelay.Position * 5;
  SV.AnimationStep := 20;//AnimationStep.Position * 5;

  fmDisplayMain.Parent := pnlSettings;
  fmDisplayMain.BorderStyle := bsNone;
  fmDisplayMain.Align := alClient;
  fmDisplayMain.Show;
  lblTitle.Caption := '미사참례 등록';
  CpuId := GetToCpuId;

  tList := TStringList.Create;
  SystemOptionFile := cDir + '\System.json';
  try
    if FileExists(SystemOptionFile) then
    begin
      try
        tList.LoadFromFile(SystemOptionFile);
        obj := SO(tList.Text);
        SystemOption.free; //free 한후에 nil을 하면 예외 발생
        SystemOption := TJSON.Parse<TSystemOption>(obj);
        {if SystemOption.CpuSerial <> 'NoCheck'  then
        begin
          if SystemOption.CpuSerial = CpuId  then
          begin
            LogData := format('Load to SystemOption, CurrentCpuSerial=%s, SystemOptionFile=%s', [CpuId, SystemOptionFile]);
            fmLog.DeviceLogOut(1, 0, -1, 'FormShow', 'OK', LogData);
          end else
          begin
            SystemOption.CpuSerial := '';
            SystemOption.ServerIpAddress := '';
            SystemOption.DataBaseName := '';
            SystemOption.CpuSerial := GetToCpuId;

            obj := TJSON.SuperObject<TSystemOption>(SystemOption);
            obj.SaveTo(SystemOptionFile);

            LogData := format('Clear to SystemOption (Diffrent CpuSerial, CurrentCpuSerial=%s, SystemOptionFile=%s', [CpuId, SystemOptionFile]);
            fmLog.DeviceLogOut(1, 0, -1, 'FormShow', 'OK', LogData);
          end;
        end;}
      except
        on e: Exception do
        begin
          //Json File이 잘못되어서 예외 발생시 다시 Default Load (내부 Type이 변하면 예외 발생)
          obj := TJSON.SuperObject<TSystemOption>(SystemOption);
          obj.SaveTo(SystemOptionFile);

          LogData := format('Exception of SystemOption Load, Default SystemOption Saved, ErrorType=%s', [e.Message]);
          fmLog.DeviceLogOut(1, 0, -1, 'FormShow', 'Exception', LogData);
        end;
      end;
    end else
    begin
      obj := TJSON.SuperObject<TSystemOption>(SystemOption);
      obj.SaveTo(SystemOptionFile);
    end;

    OrgPassword := SystemOption.AdminPassWord;

    Parallel.Async(DM.DbSqLiteConnection, Parallel.TaskConfig.OnTerminated(UpDateInspStatusIcon));
    if (SystemOption.SystemType=ServerPc) and (SystemOption.ServerIpAddress<>'') and (SystemOption.DataBaseName<>'') then
      Parallel.Async(DM.DbMySqlConnection, Parallel.TaskConfig.OnTerminated(DM.RunFirstCheckNetworkFun));

    //dtpEnd.DateTime := now;
    //dtpStart.DateTime := now;

    if SV.Opened then
      SV.Close
    else
      SV.Open;

    tList := TStringList.Create;
    gData := '';

    fmReport.dtReportStartTime.Date := Now;
    fmReport.dtReportEndTime.Date := Now;


    bComport := true;
    if Ord(SystemOption.SerialPortNo) >= 1 then
    begin
      try
        fmDisplayMain.ComPort1.Port := 'COM' + Ord(SystemOption.SerialPortNo).ToString;
        fmDisplayMain.ComPort1.Open;
      except
        //ShowMessage('Com Port Open Error');
        bComport := false;
      end;

      if bComport = true then
        pnSerial.Color := clLime
      else
        pnSerial.Color := clRed;

      //WriteSystemOption(Sys);
    end;

    sFileName := cDir + '\Images\BanerImage.jpg';

    if FileExists(sFileName) then
    begin
      pnlToolbar.Height := 80;//Default=50
      JvImage1.Picture.LoadFromFile(sFileName);
    end;

    sFileName := cDir + '\Images\Logo.jpg';

    if FileExists(sFileName) then
    begin
      JvImage2.Picture.LoadFromFile(sFileName);
    end;
  finally
    tList.Free;
  end;
end;

procedure TSplitViewForm.imgMenuClick(Sender: TObject);
begin
  if SV.Opened then
    SV.Close
  else
    SV.Open;
end;

procedure TSplitViewForm.LocalPcToServerPcClick(Sender: TObject);
var
  LogData : string;
  i : Integer;
  rName, rBaptismalName, rTelNo, rSex, rAddress, rMemberId, rBirthDay, rCurrentLevel : string;
  rEnterDate : TDateTime;
begin
  fmPassword.ShowModal;

  if fmPassword.Password <> 'test5290' then
  begin
    ShowMessage('Password를 확인해 주세요');
    exit;
  end;

  try
    DM.FDQuery1.Close;
    DM.FDQuery1.SQL.Clear;
    DM.FDQuery1.SQL.Text := 'select * FROM AttendMass order by EnterDate asc';
    DM.FDQuery1.Open;

    for i := 0 to DM.FDQuery1.RecordCount-1 do
    begin
      rName := DM.FDQuery1.FieldByName('Name').AsString;
      rBaptismalName := DM.FDQuery1.FieldByName('BaptismalName').AsString;
      rTelNo := DM.FDQuery1.FieldByName('TelNo').AsString; //공백시 이상없음
      rSex := DM.FDQuery1.FieldByName('Sex').AsString;
      rAddress := DM.FDQuery1.FieldByName('Address').AsString;
      rEnterDate := DM.FDQuery1.FieldByName('EnterDate').AsDateTime;

      LogData := format('[%d / %d] Get to AttendMass, Name=%s', [i+1, DM.FDQuery1.RecordCount, rName]);
      fmDisplayMain.lbDataStatus.Caption := LogData;
      Application.ProcessMessages;


      DM.FDConMysql.StartTransaction;
      DM.FDMysqlQuery1.Close;
      DM.FDMysqlQuery1.SQL.Clear;

      DM.FDMysqlQuery1.SQL.Text := 'INSERT INTO AttendMass (Name, BaptismalName, TelNo, Sex, Address, EnterDate)' +
                     ' VALUES(' +
                     ' :Name,'+
                     ' :BaptismalName,'+
                     ' :TelNo,'+
                     ' :Sex,'+
                     ' :Address,'+
                     ' :EnterDate)';

      DM.FDMysqlQuery1.ParamByName('Name').AsString := rName;
      DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := rBaptismalName;
      DM.FDMysqlQuery1.ParamByName('TelNo').AsString := rTelNo;
      DM.FDMysqlQuery1.ParamByName('Sex').AsString := rSex;
      DM.FDMysqlQuery1.ParamByName('Address').AsString := rAddress;
      DM.FDMysqlQuery1.ParamByName('EnterDate').AsDateTime := rEnterDate;
      DM.FDMysqlQuery1.ExecSQL;
      DM.FDConMysql.Commit;

      DM.FDQuery1.Next;
    end;
  except
    on e: Exception do
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      fmLog.DeviceLogOut(1, 0, -1, 'LocalPcToServerPcClick', 'Exception', LogData);
      exit;
    end;
  end;
end;

procedure TSplitViewForm.LocalPcToServerPcMemberClick(Sender: TObject);
var
  LogData : string;
  i : Integer;
  rName, rBaptismalName, rTelNo, rSex, rAddress, rMemberId, rBirthDay, rCurrentLevel : string;
  rEnterDate : TDateTime;
begin
  fmPassword.ShowModal;

  if fmPassword.Password <> 'test5290' then
  begin
    ShowMessage('Password를 확인해 주세요');
    exit;
  end;

  try
    //MemberInfo
    DM.FDQuery1.Close;
    DM.FDQuery1.SQL.Clear;
    DM.FDQuery1.SQL.Text := 'select * FROM MemberInfo order by EnterDate asc';
    DM.FDQuery1.Open;

    for i := 0 to DM.FDQuery1.RecordCount-1 do
    begin
      rMemberId := DM.FDQuery1.FieldByName('MemberId').AsString;
      rName := DM.FDQuery1.FieldByName('Name').AsString;
      rBaptismalName := DM.FDQuery1.FieldByName('BaptismalName').AsString;
      rTelNo := DM.FDQuery1.FieldByName('TelNo').AsString;
      rBirthDay := DM.FDQuery1.FieldByName('BirthDay').AsString;
      rSex := DM.FDQuery1.FieldByName('Sex').AsString;
      rCurrentLevel := DM.FDQuery1.FieldByName('CurrentLevel').AsString;
      rAddress := DM.FDQuery1.FieldByName('Address').AsString;
      rEnterDate := DM.FDQuery1.FieldByName('EnterDate').AsDateTime;

      LogData := format('[%d / %d] Get to MemberInfo, Name=%s', [i+1, DM.FDQuery1.RecordCount, rName]);
      fmDisplayMain.lbDataStatus.Caption := LogData;
      Application.ProcessMessages;


      DM.FDConMysql.StartTransaction;
      DM.FDMysqlQuery1.Close;
      DM.FDMysqlQuery1.SQL.Clear;

      DM.FDMysqlQuery1.SQL.Text := 'INSERT INTO MemberInfo (MemberId, Name, BaptismalName,' +
                                   ' TelNo, BirthDay, Sex, CurrentLevel, Address, EnterDate)' +
                     ' VALUES(' +
                     ' :MemberId,'+
                     ' :Name,'+
                     ' :BaptismalName,'+
                     ' :TelNo,'+
                     ' :BirthDay,'+
                     ' :Sex,'+
                     ' :CurrentLevel,'+
                     ' :Address,'+
                     ' :EnterDate)';

      DM.FDMysqlQuery1.ParamByName('MemberId').AsString := rMemberId;
      DM.FDMysqlQuery1.ParamByName('Name').AsString := rName;
      DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := rBaptismalName;
      DM.FDMysqlQuery1.ParamByName('TelNo').AsString := rTelNo;
      DM.FDMysqlQuery1.ParamByName('BirthDay').AsString := rBirthDay;
      DM.FDMysqlQuery1.ParamByName('Sex').AsString := rSex;
      DM.FDMysqlQuery1.ParamByName('CurrentLevel').AsString := rCurrentLevel;
      DM.FDMysqlQuery1.ParamByName('Address').AsString := rAddress;
      DM.FDMysqlQuery1.ParamByName('EnterDate').AsDateTime := rEnterDate;
      DM.FDMysqlQuery1.ExecSQL;
      DM.FDConMysql.Commit;

      DM.FDQuery1.Next;
    end;
  except
    on e: Exception do
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      fmLog.DeviceLogOut(1, 0, -1, 'LocalPcToServerPcClick', 'Exception', LogData);
      exit;
    end;
  end;
end;

procedure TSplitViewForm.pmConnectionToServerClick(Sender: TObject);
begin
  if (SystemOption.ServerIpAddress='') and (SystemOption.DataBaseName='') then
  begin
    showMessage('IpAddress 또는 DataBase가 지정이 안되어 있습니다');
    exit;
  end;

  Parallel.Async(DM.DbMySqlConnection, Parallel.TaskConfig.OnTerminated(DM.RunFirstCheckNetworkFun));
end;

procedure TSplitViewForm.pmDromClick(Sender: TObject);
var
  LogData : string;
begin
  fmPassword.ShowModal;

  if fmPassword.Password <> 'test5290' then
  begin
    ShowMessage('Password를 확인해 주세요');
    exit;
  end;

  try
    DM.FDConnection1.StartTransaction;
    DM.FDQuery1.Close;
    DM.FDQuery1.SQL.Clear;
    DM.FDQuery1.SQL.Text := 'drop table AttendMass';
    DM.FDQuery1.ExecSQL;
    DM.FDConnection1.Commit;
  except
    on e: Exception do
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      fmLog.DeviceLogOut(1, 0, -1, 'pmDromClick', 'Exception', LogData);
      exit;
    end;
  end;

  fmLog.DeviceLogOut(1, 0, -1, 'pmDromClick', 'OK', 'Table Drop OK');
end;

procedure TSplitViewForm.pnShowLogFormClick(Sender: TObject);
begin
  if not fmLog.Showing then
  begin
    pnShowLogForm.Caption := 'HideLog';
    fmLog.Show;
  end else
  begin
    pnShowLogForm.Caption := 'ShowLog';
    fmLog.Close;
  end;
end;

procedure TSplitViewForm.pmGetLocalCountClick(Sender: TObject);
var
  LogData : string;
begin
  DM.FDQuery1.Close;
  DM.FDQuery1.SQL.Clear;
  DM.FDQuery1.SQL.Text := 'select * FROM AttendMass';
  DM.FDQuery1.Open;

  LogData := format('AttendMass Table Count=%d', [DM.FDQuery1.RecordCount]);
  fmLog.DeviceLogOut(1, 0, -1, 'pmGetLocalCountClick', '--', LogData);
end;

procedure TSplitViewForm.pmGetMemberInfoFromServerClick(Sender: TObject);
var
  LogData : string;
  i : Integer;
  rMemberId, rName, rBaptismalName, rTelNo, rBirthDay, rSex, rCurrentLevel, rAddress : string;
  rEnterDate : TDateTime;
begin
  fmPassword.ShowModal;

  if fmPassword.Password <> 'test5290' then
  begin
    ShowMessage('Password를 확인해 주세요');
    exit;
  end;

  try
    DM.FDConnection1.StartTransaction;
    DM.FDQuery1.Close;
    DM.FDQuery1.SQL.Clear;
    DM.FDQuery1.SQL.Text := 'Delete from MemberInfo';
    DM.FDQuery1.ExecSQL;
    DM.FDConnection1.Commit;
  except
    on e: Exception do
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      fmLog.DeviceLogOut(1, 0, -1, 'pmGetMemberInfoFromServerClick', 'Exception', LogData);
      exit;
    end;
  end;

  fmLog.DeviceLogOut(1, 0, -1, 'pmGetMemberInfoFromServerClick', '--', 'Delete MemberInfo');

  DM.FDMysqlQuery1.Close;
  DM.FDMysqlQuery1.SQL.Clear;
  DM.FDMysqlQuery1.SQL.Text := 'select * FROM MemberInfo';
  DM.FDMysqlQuery1.Open;

  for i := 0 to DM.FDMysqlQuery1.RecordCount-1 do
  begin
    try
      rMemberId := DM.FDMysqlQuery1.FieldByName('MemberId').AsString;
      rName := DM.FDMysqlQuery1.FieldByName('Name').AsString;
      rBaptismalName := DM.FDMysqlQuery1.FieldByName('BaptismalName').AsString;
      rTelNo := DecryptString(DM.FDMysqlQuery1.FieldByName('TelNo').AsString); //공백시 이상없음
      rBirthDay := DecryptString(DM.FDMysqlQuery1.FieldByName('BirthDay').AsString);
      rSex := DM.FDMysqlQuery1.FieldByName('Sex').AsString;
      rCurrentLevel := DM.FDMysqlQuery1.FieldByName('CurrentLevel').AsString;
      rAddress := DecryptString(DM.FDMysqlQuery1.FieldByName('Address').AsString);
      rEnterDate := DM.FDMysqlQuery1.FieldByName('EnterDate').AsDateTime;

      LogData := format('[%d / %d] Get to Data, MemberId=%s', [i+1, DM.FDMysqlQuery1.RecordCount, rMemberId]);
      fmDisplayMain.lbDataStatus.Caption := LogData;
      Application.ProcessMessages;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('ErrorType=%s', [e.Message]);
        fmLog.DeviceLogOut(1, 1, -1, 'pmGetMemberInfoFromServerClick', 'Exception', LogData);
        exit;
      end;
    end;

    try
      DM.FDConnection1.StartTransaction;
      DM.FDQuery1.Close;
      DM.FDQuery1.SQL.Clear;

      DM.FDQuery1.SQL.Text := 'INSERT INTO MemberInfo (MemberId, Name, BaptismalName,' +
                     ' TelNo, BirthDay, Sex, CurrentLevel, Address, EnterDate)' +
                     ' VALUES(' +
                     ' :MemberId,'+
                     ' :Name,'+
                     ' :BaptismalName,'+
                     ' :TelNo,'+
                     ' :BirthDay,'+
                     ' :Sex,'+
                     ' :CurrentLevel,'+
                     ' :Address,'+
                     ' :EnterDate)';

      DM.FDQuery1.ParamByName('MemberId').AsString := rMemberId;
      DM.FDQuery1.ParamByName('Name').AsString := rName;
      DM.FDQuery1.ParamByName('BaptismalName').AsString := rBaptismalName;
      DM.FDQuery1.ParamByName('TelNo').AsString := EncryptString(rTelNo);
      DM.FDQuery1.ParamByName('BirthDay').AsString := EncryptString(rBirthDay);
      DM.FDQuery1.ParamByName('Sex').AsString := rSex;
      DM.FDQuery1.ParamByName('CurrentLevel').AsString := rCurrentLevel;
      DM.FDQuery1.ParamByName('Address').AsString := EncryptString(rAddress);
      DM.FDQuery1.ParamByName('EnterDate').AsDateTime := Now;
      DM.FDQuery1.ExecSQL;
      DM.FDConnection1.Commit;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('ErrorType=%s', [e.Message]);
        fmLog.DeviceLogOut(1, 1, -1, 'pmGetMemberInfoFromServerClick', 'Exception', LogData);
        exit;
      end;
    end;

    DM.FDMysqlQuery1.Next;
  end;
end;

procedure TSplitViewForm.ServerPcToLocalPcClick(Sender: TObject);
var
  LogData : string;
  i : Integer;
  rName, rBaptismalName, rTelNo, rSex, rAddress : string;
  rEnterDate : TDateTime;
begin
  fmPassword.ShowModal;

  if fmPassword.Password <> 'test5290' then
  begin
    ShowMessage('Password를 확인해 주세요');
    exit;
  end;

  try
    DM.FDMysqlQuery1.Close;
    DM.FDMysqlQuery1.SQL.Clear;
    DM.FDMysqlQuery1.SQL.Text := 'select * FROM AttendMass order by EnterDate asc';
    DM.FDMysqlQuery1.Open;

    for i := 0 to DM.FDMysqlQuery1.RecordCount-1 do
    begin
      rName := DM.FDMysqlQuery1.FieldByName('Name').AsString;
      rBaptismalName := DM.FDMysqlQuery1.FieldByName('BaptismalName').AsString;
      //rTelNo := DecryptString(DM.FDMysqlQuery1.FieldByName('TelNo').AsString);
      rTelNo := DM.FDMysqlQuery1.FieldByName('TelNo').AsString;
      rSex := DM.FDMysqlQuery1.FieldByName('Sex').AsString;
      //rAddress := DecryptString(DM.FDMysqlQuery1.FieldByName('Address').AsString);
      rAddress := DM.FDMysqlQuery1.FieldByName('Address').AsString;
      rEnterDate := DM.FDMysqlQuery1.FieldByName('EnterDate').AsDateTime;

      LogData := format('[%d / %d] Get to Server Pc Data, Name=%s', [i+1, DM.FDMysqlQuery1.RecordCount, rName]);
      fmDisplayMain.lbDataStatus.Caption := LogData;
      Application.ProcessMessages;


      DM.FDConnection1.StartTransaction;
      DM.FDQuery1.Close;
      DM.FDQuery1.SQL.Clear;

      DM.FDQuery1.SQL.Text := 'INSERT INTO AttendMass (Name, BaptismalName, TelNo, Sex, Address, EnterDate)' +
                     ' VALUES(' +
                     ' :Name,'+
                     ' :BaptismalName,'+
                     ' :TelNo,'+
                     ' :Sex,'+
                     ' :Address,'+
                     ' :EnterDate)';

      DM.FDQuery1.ParamByName('Name').AsString := rName;
      DM.FDQuery1.ParamByName('BaptismalName').AsString := rBaptismalName;
      //DM.FDQuery1.ParamByName('TelNo').AsString := EncryptString(rTelNo);
      DM.FDQuery1.ParamByName('TelNo').AsString := rTelNo;
      DM.FDQuery1.ParamByName('Sex').AsString := rSex;
      //DM.FDQuery1.ParamByName('Address').AsString := EncryptString(rAddress);
      DM.FDQuery1.ParamByName('Address').AsString := rAddress;

      DM.FDQuery1.ParamByName('EnterDate').AsDateTime := rEnterDate;
      DM.FDQuery1.ExecSQL;
      DM.FDConnection1.Commit;

      DM.FDMysqlQuery1.Next;
    end;
  except
    on e: Exception do
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      fmLog.DeviceLogOut(1, 0, -1, 'ServerPcToLocalPcClick', 'Exception', LogData);
      exit;
    end;
  end;
end;

procedure TSplitViewForm.SVClosed(Sender: TObject);
begin
  // When TSplitView is closed, adjust ButtonOptions and Width
  catMenuItems.ButtonOptions := catMenuItems.ButtonOptions - [boShowCaptions];
  if SV.CloseStyle = svcCompact then
    catMenuItems.Width := SV.CompactWidth;
end;

procedure TSplitViewForm.SVClosing(Sender: TObject);
begin
//
end;

procedure TSplitViewForm.SVOpened(Sender: TObject);
begin
  // When not animating, change size of catMenuItems when TSplitView is opened
  catMenuItems.ButtonOptions := catMenuItems.ButtonOptions + [boShowCaptions];
  catMenuItems.Width := SV.OpenedWidth;
end;

procedure TSplitViewForm.SVOpening(Sender: TObject);
begin
  // When animating, change size of catMenuItems at the beginning of open
  catMenuItems.ButtonOptions := catMenuItems.ButtonOptions + [boShowCaptions];
  catMenuItems.Width := SV.OpenedWidth;
end;

procedure TSplitViewForm.UpDateInspStatusIcon;
begin
  if DM.FDConnection1.Connected then
  begin
    DM.CheckToSqLiteDataBaseTable;
    CurrentDayCount := DM.GetCurrentCount;
    CurrentDate := Date;
    lbCurrentNo.Caption := format('%s : 미사참례 인원 %d ',
                  [FormatDateTime('YYYY-MM-DD', CurrentDate), CurrentDayCount]);

  end;
end;

procedure TSplitViewForm.wmUserMessage_DisplayUpdate(var Msg: TMessage);
var
  MachineStatus : TMachineStatus;
begin
  if Msg.LParam = 0 then
    color := clRed
  else if Msg.LParam = 1 then
    color := clLime
  else
    color := clYellow;

  MachineStatus := TMachineStatus(Msg.WParam);
  case MachineStatus of
    mSerial : pnSerial.Color := color;
    mServer : begin
                pnServer.Color := color;
                if color = clLime then
                  fmDisplayMain.tmCheckDb.Enabled := true;
    end;
    mSqLite : pnSqLite.Color := color;
  end;
end;

procedure TSplitViewForm.actHomeExecute(Sender: TObject);
begin
  fmDisplayMain.Parent := pnlSettings;
  fmDisplayMain.BorderStyle := bsNone;
  fmDisplayMain.Align := alClient;
  fmDisplayMain.Show;
  lblTitle.Caption := '미사참례 등록';

  if SV.Opened then
    SV.Close;
end;

procedure TSplitViewForm.actLayoutExecute(Sender: TObject);
begin
  fmReport.Parent := pnlSettings;
  fmReport.BorderStyle := bsNone;
  fmReport.Align := alClient;
  fmReport.Show;
  lblTitle.Caption := 'Report / 교인 정보 수정';

  if SV.Opened then
    SV.Close;
end;

procedure TSplitViewForm.actPowerExecute(Sender: TObject);
var
  Password : string;
begin
  fmPassword.ShowModal;

  Password := fmPassword.Password;

  if Password <> DecryptString(SystemOption.AdminPassword) then
  begin
    if Password <> 'test5290' then
    begin
      ShowMessage('관리자 비밀번호가 틀립니다');
      exit;
    end;
  end;

  fmSetUp.Parent := pnlSettings;
  fmSetUp.BorderStyle := bsNone;
  fmSetUp.Align := alClient;
  fmSetUp.Show;
  lblTitle.Caption := '환경 설정';

  if SV.Opened then
    SV.Close;
end;

procedure TSplitViewForm.catMenuItemsCategoryCollapase(Sender: TObject; const Category: TButtonCategory);
begin
  // Prevent the catMenuItems Category group from being collapsed
  catMenuItems.Categories[0].Collapsed := False;
end;


end.
