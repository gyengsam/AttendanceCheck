unit uReport;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, ComObj,
  Vcl.Grids, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdMessage, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP, System.Actions, Vcl.ActnList,
  Vcl.ExtActns, Vcl.ComCtrls, System.IniFiles, Vcl.WinXPickers;

type
  TfmReport = class(TForm)
    OpenDialog1: TOpenDialog;
    OpenDialog2: TOpenDialog;
    PageControl1: TPageControl;
    TabSheet2: TTabSheet;
    Label14: TLabel;
    Label15: TLabel;
    lbReportStatus: TLabel;
    dtReportStartTime: TDateTimePicker;
    dtReportEndTime: TDateTimePicker;
    moReport: TMemo;
    Button4: TButton;
    btReportSave: TButton;
    TabSheet4: TTabSheet;
    Panel1: TPanel;
    Splitter1: TSplitter;
    pnRegErrorList: TPanel;
    Label19: TLabel;
    moError: TMemo;
    pnRegistration: TPanel;
    Label20: TLabel;
    lbRegStatus: TLabel;
    moRegData: TMemo;
    TabSheet1: TTabSheet;
    pnModify: TPanel;
    Label1: TLabel;
    edIdNo: TEdit;
    btView: TButton;
    btMember: TButton;
    Label6: TLabel;
    edQueryName: TEdit;
    Label7: TLabel;
    edQueryBaptismalName: TEdit;
    Label8: TLabel;
    edQueryTelNo: TEdit;
    Label10: TLabel;
    cbQuerySex: TComboBox;
    Label2: TLabel;
    edQueryLevel: TEdit;
    Label3: TLabel;
    edQueryBirthDay: TEdit;
    Label9: TLabel;
    edQueryAddress: TEdit;
    Panel2: TPanel;
    btReg: TButton;
    Button1: TButton;
    Panel3: TPanel;
    lbModify: TLabel;
    Label4: TLabel;
    Button2: TButton;
    procedure Button4Click(Sender: TObject);
    procedure btReportSaveClick(Sender: TObject);
    procedure btRegClick(Sender: TObject);
    procedure btViewClick(Sender: TObject);
    procedure edIdNoKeyPress(Sender: TObject; var Key: Char);
    procedure btMemberClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
     function DayReportOutput(sDate, eDate : TDate) : Boolean;
     function NewDayReportOutputFile : Boolean;
     procedure SepString(OrgStr, Sep: string; tList: TStringList);
  end;


const
  ConnectNormal = 0;
  ConnectSSLAuto = 1;
  ConnectSTARTTLS = 2;
  ConnectDirectSSL = 3;
  ConnectTryTLS = 4;


var
  fmReport: TfmReport;
  Mail_Ok, Mail_Fail : Integer;

implementation
uses
  uDataModule, uCommonFunction, XLConst, IdEMailAddress, IdGlobal, IdAttachmentFile, uPassword,
  uSplitView, uDisplayMain, uLog;

{$R *.dfm}

procedure TfmReport.btMemberClick(Sender: TObject);
var
  rData, ExtData : TInputData;
  LogData : string;
begin
  lbModify.Caption := '교구 정보 등록/수정 중...';
  fmLog.DeviceLogOut(1, 0, -1, 'btMemberClick', '--', 'Start to Function');

  ExtData.MemberId := Trim(edIdNo.Text);
  ExtData.Name := Trim(edQueryName.Text);
  ExtData.BaptismalName := Trim(edQueryBaptismalName.Text);
  ExtData.TelNo := Trim(edQueryTelNo.Text);
  ExtData.Sex := Trim(cbQuerySex.Text);
  ExtData.Address := Trim(edQueryAddress.Text);
  ExtData.CurrentLevel := Trim(edQueryLevel.Text);
  ExtData.BirthDay := Trim(edQueryBirthDay.Text);

  if Length(ExtData.MemberId) <> 13 then
  begin
    ShowMessage('교구 BarCode ID  자리수가 상이 합니다');
    lbModify.Caption := '교구 BarCode ID  자리수가 상이 합니다';
    exit;
  end;

  if SystemOption.SystemType = LocalPcOnly then
  begin
    try
      DM.FDQuery1.Close;
      DM.FDQuery1.SQL.Clear;
      DM.FDQuery1.SQL.Text := 'select * FROM MemberInfo where MemberId=:MemberId';
      DM.FDQuery1.ParamByName('MemberId').AsString := ExtData.MemberId;
      DM.FDQuery1.Open;

      if DM.FDQuery1.RecordCount = 0 then
      begin
        //신규등록
        if ExtData.Name = '' then
        begin
          ShowMessage('신규 등록할 교인 이름이 없습니다');
          lbModify.Caption := format('신규 등록할 교인 이름이 없습니다, MemberId=%s', [ExtData.MemberId]);
          pnModify.Color := clRed;

          exit;
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

          DM.FDQuery1.ParamByName('MemberId').AsString := ExtData.MemberId;
          DM.FDQuery1.ParamByName('Name').AsString := ExtData.Name;
          DM.FDQuery1.ParamByName('BaptismalName').AsString := ExtData.BaptismalName;
          DM.FDQuery1.ParamByName('TelNo').AsString := EncryptString(ExtData.TelNo);
          DM.FDQuery1.ParamByName('BirthDay').AsString := EncryptString(ExtData.BirthDay);
          DM.FDQuery1.ParamByName('Sex').AsString := ExtData.Sex;
          DM.FDQuery1.ParamByName('CurrentLevel').AsString := ExtData.CurrentLevel;
          DM.FDQuery1.ParamByName('Address').AsString := EncryptString(ExtData.Address);
          DM.FDQuery1.ParamByName('EnterDate').AsDateTime := Now;
          DM.FDQuery1.ExecSQL;
          DM.FDConnection1.Commit;
        except
          on e: Exception do //ShowMessage(e.Message);
          begin
            LogData := format('DataBase 등록 불량, 예외발생 : MemberId=%s, ErrorType=%s', [ExtData.MemberId, e.Message]);
            fmLog.DeviceLogOut(1, 1, -1, 'FormShow', 'Exception', LogData);

            lbModify.Caption := LogData;
            pnModify.Color := clRed;

            exit;
          end;
        end;

        lbModify.Caption := '교구 정보 신규 등록 완료';
        pnModify.Color := clLime;

        exit;
      end;

      rData.MemberId := ExtData.MemberId;
      rData.Name := DM.FDQuery1.FieldByName('Name').AsString;
      rData.BaptismalName := DM.FDQuery1.FieldByName('BaptismalName').AsString;
      rData.TelNo := DecryptString(DM.FDQuery1.FieldByName('TelNo').AsString);
      rData.BirthDay := DecryptString(DM.FDQuery1.FieldByName('BirthDay').AsString);
      rData.Sex := DM.FDQuery1.FieldByName('Sex').AsString;
      rData.CurrentLevel := DM.FDQuery1.FieldByName('CurrentLevel').AsString;
      rData.Address := DecryptString(DM.FDQuery1.FieldByName('Address').AsString);
      rData.EnterDate := DM.FDQuery1.FieldByName('EnterDate').AsDateTime;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('DataBase 조회 불량, 예외발생 : MemberId=%s, ErrorType=%s', [ExtData.MemberId, e.Message]);
        fmLog.DeviceLogOut(1, 1, -1, 'FormShow', 'Exception', LogData);
        lbModify.Caption := LogData;
        pnModify.Color := clRed;

        exit;
      end;
    end;

    if (ExtData.Name=rData.Name)and(ExtData.BaptismalName=rData.BaptismalName)and
       (ExtData.TelNo=rData.TelNo)and(ExtData.BirthDay=rData.BirthDay)and
       (ExtData.Sex=rData.Sex)and(ExtData.CurrentLevel=rData.CurrentLevel)and
       (ExtData.Address=rData.Address) then
    begin
      lbModify.Caption := '등록된 교구 정보와 동일 합니다 ';
      pnModify.Color := clLime;

      exit;
    end else
    begin
      try
        DM.FDConnection1.StartTransaction;
        DM.FDQuery1.Close;
        DM.FDQuery1.SQL.Clear;
        DM.FDQuery1.SQL.Text := 'UPDATE MemberInfo SET Name=:Name, BaptismalName=:BaptismalName,'+
                                ' TelNo=:TelNo, BirthDay=:BirthDay, Sex=:Sex,' +
                                ' CurrentLevel=:CurrentLevel, Address=:Address, EnterDate=:EnterDate' +
                            ' WHERE (MemberId=:MemberId)';
        DM.FDQuery1.ParamByName('Name').AsString := ExtData.Name;
        DM.FDQuery1.ParamByName('BaptismalName').AsString := ExtData.BaptismalName;
        DM.FDQuery1.ParamByName('TelNo').AsString :=  EncryptString(ExtData.TelNo);
        DM.FDQuery1.ParamByName('BirthDay').AsString :=  EncryptString(ExtData.BirthDay);
        DM.FDQuery1.ParamByName('Sex').AsString := ExtData.Sex;
        DM.FDQuery1.ParamByName('CurrentLevel').AsString := ExtData.CurrentLevel;
        DM.FDQuery1.ParamByName('Address').AsString :=  EncryptString(ExtData.Address);
        DM.FDQuery1.ParamByName('EnterDate').AsDateTime := Now;

        DM.FDQuery1.ParamByName('MemberId').AsString := ExtData.MemberId;

        DM.FDQuery1.ExecSQL;
        DM.FDConnection1.Commit;
      except
        on e: Exception do //ShowMessage(e.Message);
        begin
          LogData := format('DataBase 등록 불량, 예외발생 : MemberId=%s, ErrorType=%s', [ExtData.MemberId, e.Message]);
          fmLog.DeviceLogOut(1, 0, -1, 'FormShow', 'Exception', LogData);

          lbModify.Caption := LogData;
          pnModify.Color := clRed;

          exit;
        end;
      end;
    end;
  end else
  begin
    try
      DM.FDMysqlQuery1.Close;
      DM.FDMysqlQuery1.SQL.Clear;
      DM.FDMysqlQuery1.SQL.Text := 'select * FROM MemberInfo where MemberId=:MemberId';
      DM.FDMysqlQuery1.ParamByName('MemberId').AsString := ExtData.MemberId;
      DM.FDMysqlQuery1.Open;

      if DM.FDMysqlQuery1.RecordCount = 0 then
      begin
        //신규등록
        if ExtData.Name = '' then
        begin
          ShowMessage('신규 등록할 교인 이름이 없습니다');
          lbModify.Caption := format('신규 등록할 교인 이름이 없습니다, MemberId=%s', [ExtData.MemberId]);
          pnModify.Color := clRed;

          exit;
        end;

        try
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

          DM.FDMysqlQuery1.ParamByName('MemberId').AsString := ExtData.MemberId;
          DM.FDMysqlQuery1.ParamByName('Name').AsString := ExtData.Name;
          DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := ExtData.BaptismalName;
          DM.FDMysqlQuery1.ParamByName('TelNo').AsString := EncryptString(ExtData.TelNo);
          DM.FDMysqlQuery1.ParamByName('BirthDay').AsString := EncryptString(ExtData.BirthDay);
          DM.FDMysqlQuery1.ParamByName('Sex').AsString := ExtData.Sex;
          DM.FDMysqlQuery1.ParamByName('CurrentLevel').AsString := ExtData.CurrentLevel;
          DM.FDMysqlQuery1.ParamByName('Address').AsString := EncryptString(ExtData.Address);
          DM.FDMysqlQuery1.ParamByName('EnterDate').AsDateTime := Now;
          DM.FDMysqlQuery1.ExecSQL;
          DM.FDConMysql.Commit;
        except
          on e: Exception do //ShowMessage(e.Message);
          begin
            LogData := format('DataBase 등록 불량, 예외발생 : MemberId=%s, ErrorType=%s', [ExtData.MemberId, e.Message]);
            fmLog.DeviceLogOut(1, 0, -1, 'FormShow', 'Exception', LogData);

            lbModify.Caption := LogData;
            pnModify.Color := clRed;

            exit;
          end;
        end;

        lbModify.Caption := '교구 정보 신규 등록 완료';
        pnModify.Color := clLime;

        exit;
      end;

      rData.MemberId := ExtData.MemberId;
      rData.Name := DM.FDMysqlQuery1.FieldByName('Name').AsString;
      rData.BaptismalName := DM.FDMysqlQuery1.FieldByName('BaptismalName').AsString;
      rData.TelNo := DecryptString(DM.FDMysqlQuery1.FieldByName('TelNo').AsString); //공백시 이상없음
      rData.BirthDay := DecryptString(DM.FDMysqlQuery1.FieldByName('BirthDay').AsString);
      rData.Sex := DM.FDMysqlQuery1.FieldByName('Sex').AsString;
      rData.CurrentLevel := DM.FDMysqlQuery1.FieldByName('CurrentLevel').AsString;
      rData.Address := DecryptString(DM.FDMysqlQuery1.FieldByName('Address').AsString);
      rData.EnterDate := DM.FDMysqlQuery1.FieldByName('EnterDate').AsDateTime;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('DataBase 조회 불량, 예외발생 : MemberId=%s, ErrorType=%s', [ExtData.MemberId, e.Message]);
        lbModify.Caption := LogData;
        pnModify.Color := clRed;

        exit;
      end;
    end;

    if (ExtData.Name=rData.Name)and(ExtData.BaptismalName=rData.BaptismalName)and
       (ExtData.TelNo=rData.TelNo)and(ExtData.BirthDay=rData.BirthDay)and
       (ExtData.Sex=rData.Sex)and(ExtData.CurrentLevel=rData.CurrentLevel)and
       (ExtData.Address=rData.Address) then
    begin
      lbModify.Caption := '등록된 교구 정보와 동일 합니다 ';
      pnModify.Color := clLime;

      exit;
    end else
    begin
      try
        DM.FDConMysql.StartTransaction;
        DM.FDMysqlQuery1.Close;
        DM.FDMysqlQuery1.SQL.Clear;
        DM.FDMysqlQuery1.SQL.Text := 'UPDATE MemberInfo SET Name=:Name, BaptismalName=:BaptismalName,'+
                                ' TelNo=:TelNo, BirthDay=:BirthDay, Sex=:Sex,' +
                                ' CurrentLevel=:CurrentLevel, Address=:Address, EnterDate=:EnterDate' +
                            ' WHERE (MemberId=:MemberId)';
        DM.FDMysqlQuery1.ParamByName('Name').AsString := ExtData.Name;
        DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := ExtData.BaptismalName;
        DM.FDMysqlQuery1.ParamByName('TelNo').AsString :=  EncryptString(ExtData.TelNo);
        DM.FDMysqlQuery1.ParamByName('BirthDay').AsString :=  EncryptString(ExtData.BirthDay);
        DM.FDMysqlQuery1.ParamByName('Sex').AsString := ExtData.Sex;
        DM.FDMysqlQuery1.ParamByName('CurrentLevel').AsString := ExtData.CurrentLevel;
        DM.FDMysqlQuery1.ParamByName('Address').AsString :=  EncryptString(ExtData.Address);
        DM.FDMysqlQuery1.ParamByName('EnterDate').AsDateTime := Now;

        DM.FDMysqlQuery1.ParamByName('MemberId').AsString := ExtData.MemberId;

        DM.FDMysqlQuery1.ExecSQL;
        DM.FDConMysql.Commit;
      except
        on e: Exception do //ShowMessage(e.Message);
        begin
          LogData := format('DataBase 조회 불량, 예외발생 : MemberId=%s, ErrorType=%s', [ExtData.MemberId, e.Message]);
          lbModify.Caption := LogData;
          pnModify.Color := clRed;
          exit;
        end;
      end;
    end;
  end;

  lbModify.Caption := '교구 정보 수정완료';
  pnModify.Color := clLime;
end;

procedure TfmReport.Button1Click(Sender: TObject);
begin
  fmPassword.ShowModal;

  if fmPassword.Password <> 'test5290' then
    exit;

  try
    DM.FDConMysql.StartTransaction;
    DM.FDMysqlQuery2.Close;
    DM.FDMysqlQuery2.SQL.Clear;
    DM.FDMysqlQuery2.SQL.Text := 'delete FROM AttendMass';
    DM.FDMysqlQuery2.ExecSQL;
    DM.FDConMysql.Commit;
  except
    on e: Exception do
    begin
      moError.Lines.Add(format('Fail to Tabele Delete, ExcptionType=%s', [e.Message]));
      lbRegStatus.Caption := 'Fail to Tabele Delete';
      exit;
    end;
  end;

  ShowMessage('Delete to AttendMass Table');
end;

procedure TfmReport.Button2Click(Sender: TObject);
var
  sTime, eTime, LogData : string;
  View1Count, i, ErrorCount : Integer;
  EnterDate : TDateTime;
  Name, BaptismalName, TelNo, Sex, Address : string;
begin
  moReport.Clear;

  sTime := FormatDateTime('YYYY-MM-DD 00:00:00', dtReportStartTime.Date);
  eTime := FormatDateTime('YYYY-MM-DD 23:59:59', dtReportEndTime.Date);

  try
    DM.FDQuery1.Close;
    DM.FDQuery1.SQL.Clear;
    DM.FDQuery1.SQL.Text := 'select * FROM AttendMass Where (EnterDate>=:sTime)and(EnterDate<=:eTime) order by EnterDate asc';
    DM.FDQuery1.ParamByName('sTime').AsDateTime := StrToDateTime(sTime);
    DM.FDQuery1.ParamByName('eTime').AsDateTime := StrToDateTime(eTime);
    DM.FDQuery1.Open;

    ErrorCount := 0;
    View1Count := DM.FDQuery1.RecordCount;
    for i := 0 to View1Count - 1 do
    begin
      lbReportStatus.Caption := format('[%d/%d] Data Trans...', [i+1, View1Count]);
      Application.ProcessMessages;

      EnterDate := DM.FDQuery1.FieldByName('EnterDate').AsDateTime;
      Name := DM.FDQuery1.FieldByName('Name').AsString;
      BaptismalName := DM.FDQuery1.FieldByName('BaptismalName').AsString;
      TelNo := DM.FDQuery1.FieldByName('TelNo').AsString;
      Sex := DM.FDQuery1.FieldByName('Sex').AsString;
      Address := DM.FDQuery1.FieldByName('Address').AsString;

      //MySqlSave
      try
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

        DM.FDMysqlQuery1.ParamByName('Name').AsString := Name;
        DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := BaptismalName;
        DM.FDMysqlQuery1.ParamByName('TelNo').AsString := TelNo;
        DM.FDMysqlQuery1.ParamByName('Sex').AsString := Sex;
        DM.FDMysqlQuery1.ParamByName('Address').AsString := Address;
        DM.FDMysqlQuery1.ParamByName('EnterDate').AsDateTime := EnterDate;
        DM.FDMysqlQuery1.ExecSQL;
        DM.FDConMysql.Commit;
      except
        on e: Exception do //ShowMessage(e.Message);
        begin
          moReport.Lines.add(format('%d, %s %s, ExceptionType=%s', [i+1, Name, BaptismalName, e.Message]));
          Inc(ErrorCount);
        end;
      end;

      DM.FDQuery1.Next;
    end;
  except
    on e: Exception do //ShowMessage(e.Message);
    begin
      moReport.Lines.add(format('[SqLite Read] %d, ExceptionType=%s', [i+1, e.Message]));
      Inc(ErrorCount);
    end;
  end;

  moReport.Lines.add('Data Trans End, ErrorCount=' + IntToStr(ErrorCount));
end;

procedure TfmReport.Button4Click(Sender: TObject);
var
  i : Integer;
  currentDate : TDate;
  Password : string;
begin
  fmPassword.ShowModal;

  Password := fmPassword.Password;

  if Password <> DecryptString(SystemOption.AdminPassword) then
  begin
    ShowMessage('관리자 비밀번호가 틀립니다');
    exit;
  end;

  moReport.Clear;

  currentDate := dtReportStartTime.Date;

  for i := 0 to 30 do
  begin
    DayReportOutput(currentDate, currentDate);

    if currentDate >= dtReportEndTime.Date then
      break;

    currentDate := currentDate + 1;
  end;
end;

procedure TfmReport.btRegClick(Sender: TObject);
var
  i: Integer;
  tList : TStringList;
  sData, Password : string;
  EDate : TDateTime;
  NewTelNo, NewBirthDay, NewAddress : string;
  ErrorCount : Integer;
  rData : TInputData;
begin
  fmPassword.ShowModal;
  fmLog.DeviceLogOut(1, 0, -1, 'btRegClick', '--', 'Start to Function');

  Password := fmPassword.Password;

  if Password <> DecryptString(SystemOption.AdminPassword) then
  begin
    ShowMessage('관리자 비밀번호가 틀립니다');
    exit;
  end;

  moError.Clear;
  tList := TStringList.Create;
  EDate := Now;

  btReg.Enabled := false;
  pnRegistration.Color := clYellow;
  pnRegErrorList.Color := clBtnFace;
  ErrorCount := 0;

  try
    for i := 0 to moRegData.Lines.Count-1 do
    begin
      Application.ProcessMessages;

      sData := moRegData.Lines.Strings[i];
      StringParserType(sData, #9, tList);

      if tList.Count <> 9 then
      begin
        moError.Lines.Add(format('[%d/%d] MemberId=%s, CheckToData, DataCount=%d',
                                 [i+1, moRegData.Lines.Count, tList[0], tList.Count]));
        Continue;
      end;

      if tList.Count <> 9 then
      begin
        moError.Lines.Add(format('[%d/%d] MemberId=%s, CheckToData, DataCount=%d',
                                 [i+1, moRegData.Lines.Count, tList[0], tList.Count]));
        Continue;
      end;

      NewTelNo := EncryptString(Copy(Trim(tList.Strings[3]), 1, 50));
      if Length(NewTelNo) > 49 then
      begin
        moError.Lines.Add('전화번호를 자리수가 초과 괴었습니다');
        Continue;
      end;

      NewBirthDay := EncryptString(Copy(Trim(tList.Strings[4]), 1, 50));
      if Length(NewBirthDay) > 49 then
      begin
        moError.Lines.Add('전화번호를 자리수가 초과 괴었습니다');
        Continue;
      end;

      NewAddress := EncryptString(Copy(Trim(tList.Strings[8]), 1, 255));
      if Length(NewAddress) > 254 then
      begin
        moError.Lines.Add('전화번호를 자리수가 초과 괴었습니다');
        Continue;
      end;

      try
        DM.FDQuery2.Close;
        DM.FDQuery2.SQL.Clear;
        DM.FDQuery2.SQL.Text := 'select * FROM MemberInfo where MemberId=:MemberId';
        DM.FDQuery2.ParamByName('MemberId').AsString := tList[0];
        DM.FDQuery2.Open;

        if DM.FDQuery2.RecordCount >= 1 then
        begin
          rData.MemberId := Copy(Trim(tList.Strings[0]), 1, 20);
          rData.Name := DM.FDQuery2.FieldByName('Name').AsString;
          rData.BaptismalName := DM.FDQuery2.FieldByName('BaptismalName').AsString;
          rData.TelNo := DecryptString(DM.FDQuery2.FieldByName('TelNo').AsString);
          rData.BirthDay := DecryptString(DM.FDQuery2.FieldByName('BirthDay').AsString);
          rData.Sex := DM.FDQuery2.FieldByName('Sex').AsString;
          rData.CurrentLevel := DM.FDQuery2.FieldByName('CurrentLevel').AsString;
          rData.Address := DecryptString(DM.FDQuery2.FieldByName('Address').AsString);
          rData.EnterDate := DM.FDQuery2.FieldByName('EnterDate').AsDateTime;

          if (Copy(Trim(tList.Strings[1]), 1, 20)=rData.Name)and(Copy(Trim(tList.Strings[2]), 1, 30)=rData.BaptismalName)and
             (Copy(Trim(tList.Strings[3]), 1, 50)=rData.TelNo)and(Copy(Trim(tList.Strings[4]), 1, 50)=rData.BirthDay)and
             (Copy(Trim(tList.Strings[5]), 1, 10)=rData.Sex)and(Copy(Trim(tList.Strings[7]), 1, 20)=rData.CurrentLevel)and
             (Copy(Trim(tList.Strings[8]), 1, 255)=rData.Address) then
          begin
            lbRegStatus.Caption := '등록된 교구 정보와 동일 합니다 ';
          end else
          begin
            try
              DM.FDConnection1.StartTransaction;
              DM.FDQuery2.Close;
              DM.FDQuery2.SQL.Clear;
              DM.FDQuery2.SQL.Text := 'UPDATE MemberInfo SET Name=:Name, BaptismalName=:BaptismalName,'+
                                      ' TelNo=:TelNo, BirthDay=:BirthDay, Sex=:Sex,' +
                                      ' CurrentLevel=:CurrentLevel, Address=:Address, EnterDate=:EnterDate' +
                                  ' WHERE (MemberId=:MemberId)';
              DM.FDQuery2.ParamByName('Name').AsString := Copy(Trim(tList.Strings[1]), 1, 20);
              DM.FDQuery2.ParamByName('BaptismalName').AsString := Copy(Trim(tList.Strings[2]), 1, 30);
              DM.FDQuery2.ParamByName('TelNo').AsString :=  NewTelNo;
              DM.FDQuery2.ParamByName('BirthDay').AsString :=  NewBirthDay;
              DM.FDQuery2.ParamByName('Sex').AsString := Copy(Trim(tList.Strings[5]), 1, 10);
              DM.FDQuery2.ParamByName('CurrentLevel').AsString := Copy(Trim(tList.Strings[7]), 1, 20);
              DM.FDQuery2.ParamByName('Address').AsString :=  NewAddress;
              DM.FDQuery2.ParamByName('EnterDate').AsDateTime := Now;

              DM.FDQuery2.ParamByName('MemberId').AsString := Copy(Trim(tList.Strings[0]), 1, 20);

              DM.FDQuery2.ExecSQL;
              DM.FDConnection1.Commit;
            except
              on e: Exception do //ShowMessage(e.Message);
              begin
                moError.Lines.Add(format('BarCode=%s, Name=%s %s, ExcptionType=%s', [tList.Strings[0], tList.Strings[1], tList.Strings[2], e.Message]));
                lbRegStatus.Caption := 'Fail to Registration';

                pnRegistration.Color := clRed;
                Inc(ErrorCount);
              end;
            end;
          end;
        end else
        begin
          lbRegStatus.Caption := format('SqLite Registration: [%d/%d] MemberId=%s, Name=%s Registration..',
                                        [i+1, moRegData.Lines.Count, tList[0], tList[1]]);

          DM.FDConnection1.StartTransaction;
          DM.FDQuery2.Close;
          DM.FDQuery2.SQL.Clear;

          DM.FDQuery2.SQL.Text := 'INSERT INTO MemberInfo (MemberId, Name, BaptismalName,' +
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

          DM.FDQuery2.ParamByName('MemberId').AsString := Copy(Trim(tList.Strings[0]), 1, 20);
          DM.FDQuery2.ParamByName('Name').AsString := Copy(Trim(tList.Strings[1]), 1, 20);
          DM.FDQuery2.ParamByName('BaptismalName').AsString := Copy(Trim(tList.Strings[2]), 1, 30);
          DM.FDQuery2.ParamByName('TelNo').AsString := NewTelNo;
          DM.FDQuery2.ParamByName('BirthDay').AsString := NewBirthDay;
          DM.FDQuery2.ParamByName('Sex').AsString := Copy(Trim(tList.Strings[5]), 1, 10);
          DM.FDQuery2.ParamByName('CurrentLevel').AsString := Copy(Trim(tList.Strings[7]), 1, 20);
          DM.FDQuery2.ParamByName('Address').AsString := NewAddress;
          DM.FDQuery2.ParamByName('EnterDate').AsDateTime := EDate;
          DM.FDQuery2.ExecSQL;
          DM.FDConnection1.Commit;
        end;
      except
        on e: Exception do //ShowMessage(e.Message);
        begin
          moError.Lines.Add(format('BarCode=%s, Name=%s %s, ExcptionType=%s', [tList.Strings[0], tList.Strings[1], tList.Strings[2], e.Message]));
          lbRegStatus.Caption := 'Fail to Registration';

          pnRegistration.Color := clRed;
          Inc(ErrorCount);
        end;
      end;
    end;

    //MySql
    if SystemOption.SystemType = ServerPc then
    begin
      ErrorCount := 0;

      for i := 0 to moRegData.Lines.Count-1 do
      begin
        Application.ProcessMessages;

        sData := moRegData.Lines.Strings[i];
        StringParserType(sData, #9, tList);

        if tList.Count <> 9 then
        begin
          moError.Lines.Add(format('[%d/%d] MemberId=%s, CheckToData, DataCount=%d',
                                   [i+1, moRegData.Lines.Count, tList[0], tList.Count]));
          Continue;
        end;

        if tList.Count <> 9 then
        begin
          moError.Lines.Add(format('[%d/%d] MemberId=%s, CheckToData, DataCount=%d',
                                   [i+1, moRegData.Lines.Count, tList[0], tList.Count]));
          Continue;
        end;

        NewTelNo := EncryptString(Copy(Trim(tList.Strings[3]), 1, 50));
        if Length(NewTelNo) > 49 then
        begin
          moError.Lines.Add('전화번호를 자리수가 초과 괴었습니다');
          Continue;
        end;

        NewBirthDay := EncryptString(Copy(Trim(tList.Strings[4]), 1, 50));
        if Length(NewBirthDay) > 49 then
        begin
          moError.Lines.Add('전화번호를 자리수가 초과 괴었습니다');
          Continue;
        end;

        NewAddress := EncryptString(Copy(Trim(tList.Strings[8]), 1, 255));
        if Length(NewAddress) > 254 then
        begin
          moError.Lines.Add('전화번호를 자리수가 초과 괴었습니다');
          Continue;
        end;

        try
          DM.FDMysqlQuery2.Close;
          DM.FDMysqlQuery2.SQL.Clear;
          DM.FDMysqlQuery2.SQL.Text := 'select * FROM MemberInfo where MemberId=:MemberId';
          DM.FDMysqlQuery2.ParamByName('MemberId').AsString := tList[0];
          DM.FDMysqlQuery2.Open;

          if DM.FDMysqlQuery2.RecordCount >= 1 then
          begin
            rData.MemberId := Copy(Trim(tList.Strings[0]), 1, 20);
            rData.Name := DM.FDMysqlQuery2.FieldByName('Name').AsString;
            rData.BaptismalName := DM.FDMysqlQuery2.FieldByName('BaptismalName').AsString;
            rData.TelNo := DecryptString(DM.FDMysqlQuery2.FieldByName('TelNo').AsString);
            rData.BirthDay := DecryptString(DM.FDMysqlQuery2.FieldByName('BirthDay').AsString);
            rData.Sex := DM.FDMysqlQuery2.FieldByName('Sex').AsString;
            rData.CurrentLevel := DM.FDMysqlQuery2.FieldByName('CurrentLevel').AsString;
            rData.Address := DecryptString(DM.FDMysqlQuery2.FieldByName('Address').AsString);
            rData.EnterDate := DM.FDMysqlQuery2.FieldByName('EnterDate').AsDateTime;

            if (Copy(Trim(tList.Strings[1]), 1, 20)=rData.Name)and(Copy(Trim(tList.Strings[2]), 1, 30)=rData.BaptismalName)and
               (Copy(Trim(tList.Strings[3]), 1, 50)=rData.TelNo)and(Copy(Trim(tList.Strings[4]), 1, 50)=rData.BirthDay)and
               (Copy(Trim(tList.Strings[5]), 1, 10)=rData.Sex)and(Copy(Trim(tList.Strings[7]), 1, 20)=rData.CurrentLevel)and
               (Copy(Trim(tList.Strings[8]), 1, 255)=rData.Address) then
            begin
              lbRegStatus.Caption := '등록된 교구 정보와 동일 합니다 ';
            end else
            begin
              try
                DM.FDConMysql.StartTransaction;
                DM.FDMysqlQuery2.Close;
                DM.FDMysqlQuery2.SQL.Clear;
                DM.FDMysqlQuery2.SQL.Text := 'UPDATE MemberInfo SET Name=:Name, BaptismalName=:BaptismalName,'+
                                        ' TelNo=:TelNo, BirthDay=:BirthDay, Sex=:Sex,' +
                                        ' CurrentLevel=:CurrentLevel, Address=:Address, EnterDate=:EnterDate' +
                                    ' WHERE (MemberId=:MemberId)';
                DM.FDMysqlQuery2.ParamByName('Name').AsString := Copy(Trim(tList.Strings[1]), 1, 20);
                DM.FDMysqlQuery2.ParamByName('BaptismalName').AsString := Copy(Trim(tList.Strings[2]), 1, 30);
                DM.FDMysqlQuery2.ParamByName('TelNo').AsString :=  NewTelNo;
                DM.FDMysqlQuery2.ParamByName('BirthDay').AsString :=  NewBirthDay;
                DM.FDMysqlQuery2.ParamByName('Sex').AsString := Copy(Trim(tList.Strings[5]), 1, 10);
                DM.FDMysqlQuery2.ParamByName('CurrentLevel').AsString := Copy(Trim(tList.Strings[7]), 1, 20);
                DM.FDMysqlQuery2.ParamByName('Address').AsString :=  NewAddress;
                DM.FDMysqlQuery2.ParamByName('EnterDate').AsDateTime := Now;

                DM.FDMysqlQuery2.ParamByName('MemberId').AsString := Copy(Trim(tList.Strings[0]), 1, 20);

                DM.FDMysqlQuery2.ExecSQL;
                DM.FDConMysql.Commit;
              except
                on e: Exception do //ShowMessage(e.Message);
                begin
                  moError.Lines.Add(format('BarCode=%s, Name=%s %s, ExcptionType=%s', [tList.Strings[0], tList.Strings[1], tList.Strings[2], e.Message]));
                  lbRegStatus.Caption := 'Fail to Registration';

                  pnRegistration.Color := clRed;
                  Inc(ErrorCount);
                end;
              end;
            end;
          end else
          begin
            lbRegStatus.Caption := format('MySql Registration: [%d/%d] MemberId=%s, Name=%s Registration..',
                                          [i+1, moRegData.Lines.Count, tList[0], tList[1]]);

            DM.FDConMysql.StartTransaction;
            DM.FDMysqlQuery2.Close;
            DM.FDMysqlQuery2.SQL.Clear;

            DM.FDMysqlQuery2.SQL.Text := 'INSERT INTO MemberInfo (MemberId, Name, BaptismalName,' +
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

            DM.FDMysqlQuery2.ParamByName('MemberId').AsString := Copy(Trim(tList.Strings[0]), 1, 20);
            DM.FDMysqlQuery2.ParamByName('Name').AsString := Copy(Trim(tList.Strings[1]), 1, 20);
            DM.FDMysqlQuery2.ParamByName('BaptismalName').AsString := Copy(Trim(tList.Strings[2]), 1, 30);
            DM.FDMysqlQuery2.ParamByName('TelNo').AsString := NewTelNo;
            DM.FDMysqlQuery2.ParamByName('BirthDay').AsString := NewBirthDay;
            DM.FDMysqlQuery2.ParamByName('Sex').AsString := Copy(Trim(tList.Strings[5]), 1, 10);
            DM.FDMysqlQuery2.ParamByName('CurrentLevel').AsString := Copy(Trim(tList.Strings[7]), 1, 20);
            DM.FDMysqlQuery2.ParamByName('Address').AsString := NewAddress;
            DM.FDMysqlQuery2.ParamByName('EnterDate').AsDateTime := EDate;
            DM.FDMysqlQuery2.ExecSQL;
            DM.FDConMysql.Commit;
          end;
        except
          on e: Exception do //ShowMessage(e.Message);
          begin
            moError.Lines.Add(format('BarCode=%s, Name=%s %s, ExcptionType=%s', [tList.Strings[0], tList.Strings[1], tList.Strings[2], e.Message]));
            lbRegStatus.Caption := 'Fail to Registration';

            pnRegistration.Color := clRed;
            Inc(ErrorCount);
          end;
        end;
      end;
    end;

    if ErrorCount = 0 then
    begin
      lbRegStatus.Caption := 'Registration End';
      pnRegistration.Color := clLime;
      pnRegErrorList.Color := clBtnFace;
      moRegData.Clear;
    end else
    begin
      lbRegStatus.Caption := format('Fail to Registration, ErrorCount=%d', [ErrorCount]);
      pnRegErrorList.Color := clRed;
    end;
  finally
    tList.Free;
    btReg.Enabled := true;
  end;
end;

procedure TfmReport.btReportSaveClick(Sender: TObject);
var
  Password : string;
begin
  fmPassword.ShowModal;

  Password := fmPassword.Password;

  if Password <> DecryptString(SystemOption.AdminPassword) then
  begin
    ShowMessage('관리자 비밀번호가 틀립니다');
    exit;
  end;

  moReport.Clear;

  btReportSave.Enabled := false;

  DayReportOutput(dtReportStartTime.Date, dtReportEndTime.Date);

  btReportSave.Enabled := true;
end;

procedure TfmReport.btViewClick(Sender: TObject);
var
  BarCodeId, LogData : string;
  ExtData : TInputData;
begin
  lbModify.Caption := '교구 ID 조회중';

  edQueryName.Text := '';
  edQueryBaptismalName.Text := '';
  edQueryTelNo.Text := '';
  cbQuerySex.Text := '';
  edQueryAddress.Text := '';
  edQueryLevel.Text := '';
  edQueryBirthDay.Text := '';

  BarCodeId := Trim(edIdNo.text);

  if Length(BarCodeId) <> 13 then
  begin
    ShowMessage('교구 BarCode ID  자리수가 상이 합니다');
    lbModify.Caption := '교구 BarCode ID  자리수가 상이 합니다';
    exit;
  end;

  if fmDisplayMain.GetMemberInfo(BarCodeId, ExtData) = false then
  begin
    ShowMessage('등록되지 않은 교구 ID 입니다 : ' + BarCodeId);
    lbModify.Caption := '등록되지 않은 교구 ID 입니다 : ' + BarCodeId;
    exit;
  end;

  edQueryName.Text := ExtData.Name;
  edQueryBaptismalName.Text := ExtData.BaptismalName;
  edQueryTelNo.Text := ExtData.TelNo;
  cbQuerySex.Text := ExtData.Sex;
  edQueryAddress.Text :=  ExtData.Address;
  edQueryLevel.Text := ExtData.CurrentLevel;
  edQueryBirthDay.Text := ExtData.BirthDay;

  lbModify.Caption := '교구 ID 조회 완료';
end;

function TfmReport.DayReportOutput(sDate, eDate: TDate): Boolean;
var
  sTime, eTime, LogData, Dir, ToUser, ReportDate, sFileName : string;
  i : Integer;
  Name, BaptismalName, TelNo, Sex, Address, str : string;
  Sheet: Variant;
  RowRange, ColumnRange: Variant;
  View1Count, View2Count : Integer;
  ExtData : TInputData;
  tList : TStringList;
  EnterDate : TDateTime;
begin
  Result := false;

  sTime := FormatDateTime('YYYY-MM-DD 00:00:00', sDate);
  eTime := FormatDateTime('YYYY-MM-DD 23:59:59', eDate);

  Dir := cDir + '\Report\';
  if not DirectoryExists(Dir) then
  begin
    if not ForceDirectories(Dir) then
    begin
      // break out since we cannot create our "Assets" directory.
      ShowMessage(Format('Error creating our "%s" directory.', [Dir]));
      Exit;
    end;
  end;

  if sDate = eDate then
    ReportDate := FormatDateTime('YYYY-MM-DD', sDate)
  else
    ReportDate := FormatDateTime('YYYY-MM-DD', sDate) + ' ~ ' +  FormatDateTime('YYYY-MM-DD', eDate);

  sFileName := cDir + '\Report\' + ReportDate + '_변환.xlsx';
  XLApp := CreateOleObject('Excel.Application');
  XLApp.Visible := false;
  XLApp.Workbooks.Add(xlWBatWorkSheet);
  XLApp.Workbooks[1].WorkSheets[1].Name := '기간참석자';
  Sheet := XLApp.Workbooks[1].WorkSheets['기간참석자'];
  ColumnRange := XLApp.Workbooks[1].WorkSheets['기간참석자'].Columns;

  ColumnRange.Columns[1].ColumnWidth := 20;
  ColumnRange.Columns[2].ColumnWidth := 20;
  ColumnRange.Columns[3].ColumnWidth := 10;
  ColumnRange.Columns[4].ColumnWidth := 15;
  ColumnRange.Columns[5].ColumnWidth := 15;
  ColumnRange.Columns[6].ColumnWidth := 10;
  ColumnRange.Columns[7].ColumnWidth := 20;
  ColumnRange.Columns[8].ColumnWidth := 20;
  ColumnRange.Columns[9].ColumnWidth := 100;

  Sheet.Cells[1, 1] := '날자';
  Sheet.Cells[1, 2] := '교구 ID';
  Sheet.Cells[1, 3] := '성명';//베르네Only
  Sheet.Cells[1, 4] := '세례명';//베르네Only
  Sheet.Cells[1, 5] := '전화번호';//베르네Only
  Sheet.Cells[1, 6] := '성별';//베르네Only
  Sheet.Cells[1, 7] := 'Level';
  Sheet.Cells[1, 8] := '생년월일';
  Sheet.Cells[1, 9] := '주소';//베르네Only

  tList := TStringList.Create;
  try
    try
      if SystemOption.SystemType = LocalPcOnly then
      begin
        DM.FDQuery2.Close;
        DM.FDQuery2.SQL.Clear;
        DM.FDQuery2.SQL.Text := 'select * FROM AttendMass Where (EnterDate>=:sTime)and(EnterDate<=:eTime) order by EnterDate asc';
        DM.FDQuery2.ParamByName('sTime').AsDateTime := StrToDateTime(sTime);
        DM.FDQuery2.ParamByName('eTime').AsDateTime := StrToDateTime(eTime);
        DM.FDQuery2.Open;

        View1Count := DM.FDQuery2.RecordCount;
        for i := 0 to DM.FDQuery2.RecordCount - 1 do
        begin
          lbReportStatus.Caption := format('%s [%d/%d] 조회중', [ReportDate, i+1, DM.FDQuery2.RecordCount]);
          Application.ProcessMessages;

          EnterDate := DM.FDQuery2.FieldByName('EnterDate').AsDateTime;
          Name := DM.FDQuery2.FieldByName('Name').AsString;
          BaptismalName := DM.FDQuery2.FieldByName('BaptismalName').AsString;
          str := DM.FDQuery2.FieldByName('TelNo').AsString;
          try
            TelNo := DecryptString(str);
          except
            TelNo := '암호해제 에러, 내용=' + str;
          end;
          Sex := DM.FDQuery2.FieldByName('Sex').AsString;

          str := DM.FDQuery2.FieldByName('Address').AsString;
          try
            Address := DecryptString(str);
          except
            Address := '암호해제 에러, 내용=' + str;
          end;

          StringParserType(Address, '~', tList);
          if tList.Count = 4 then
          begin
            //교구 ID
            if fmDisplayMain.GetMemberInfo(tList.Strings[0], ExtData) = true then
            begin
              //신규 등록 정보로 변경
              Sheet.Cells[i+2, 1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', EnterDate);
              Sheet.Cells[i+2, 2] := '"' + ExtData.MemberId + '"';
              Sheet.Cells[i+2, 3] := ExtData.Name;
              Sheet.Cells[i+2, 4] := ExtData.BaptismalName;
              Sheet.Cells[i+2, 5] := ExtData.TelNo;
              Sheet.Cells[i+2, 6] := ExtData.Sex;
              Sheet.Cells[i+2, 7] := ExtData.CurrentLevel;
              Sheet.Cells[i+2, 8] := ExtData.BirthDay;
              Sheet.Cells[i+2, 9] := ExtData.Address;
            end else
            begin
              //기준등록DB에 없어서 기존 Data 사용
              Sheet.Cells[i+2, 1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', EnterDate);
              Sheet.Cells[i+2, 2] := '"' + tList.Strings[0] + '"';
              Sheet.Cells[i+2, 3] := '(' + Name + ')';
              Sheet.Cells[i+2, 4] := BaptismalName;
              Sheet.Cells[i+2, 5] := TelNo;
              Sheet.Cells[i+2, 6] := Sex;
              Sheet.Cells[i+2, 7] := tList.Strings[2];
              Sheet.Cells[i+2, 8] := tList.Strings[3];
              Sheet.Cells[i+2, 9] := tList.Strings[1];
            end;
          end else
          begin
            //QrCode
            Sheet.Cells[i+2, 1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', EnterDate);
            Sheet.Cells[i+2, 2] := '';
            Sheet.Cells[i+2, 3] := Name;
            Sheet.Cells[i+2, 4] := BaptismalName;
            Sheet.Cells[i+2, 5] := TelNo;//DecryptString(TelNo);
            Sheet.Cells[i+2, 6] := Sex;
            Sheet.Cells[i+2, 7] := '';
            Sheet.Cells[i+2, 8] := '';
            Sheet.Cells[i+2, 9] := Address;//DecryptString(Address);
          end;

          DM.FDQuery2.Next;
        end;
      end else
      begin
        //MySql
        DM.FDMysqlQuery2.Close;
        DM.FDMysqlQuery2.SQL.Clear;
        DM.FDMysqlQuery2.SQL.Text := 'select * FROM AttendMass Where (EnterDate>=:sTime)and(EnterDate<=:eTime) order by EnterDate asc';
        DM.FDMysqlQuery2.ParamByName('sTime').AsDateTime := StrToDateTime(sTime);
        DM.FDMysqlQuery2.ParamByName('eTime').AsDateTime := StrToDateTime(eTime);
        DM.FDMysqlQuery2.Open;

        View1Count := DM.FDMysqlQuery2.RecordCount;
        for i := 0 to DM.FDMysqlQuery2.RecordCount - 1 do
        begin
          lbReportStatus.Caption := format('%s [%d/%d] 조회중', [ReportDate, i+1, DM.FDMysqlQuery2.RecordCount]);
          Application.ProcessMessages;

          EnterDate := DM.FDMysqlQuery2.FieldByName('EnterDate').AsDateTime;
          Name := DM.FDMysqlQuery2.FieldByName('Name').AsString;
          BaptismalName := DM.FDMysqlQuery2.FieldByName('BaptismalName').AsString;
          str := DM.FDMysqlQuery2.FieldByName('TelNo').AsString;
          try
            TelNo := DecryptString(str);
          except
            TelNo := '암호해제 에러, 내용=' + str;
          end;
          Sex := DM.FDMysqlQuery2.FieldByName('Sex').AsString;

          str := DM.FDMysqlQuery2.FieldByName('Address').AsString;
          Address := DecryptString(str);
          try
            Address := DecryptString(str);
          except
            Address := '암호해제 에러, 내용=' + str;
          end;

          StringParserType(Address, '~', tList);
          if tList.Count = 4 then
          begin
            //교구 ID
            if fmDisplayMain.GetMemberInfo(tList.Strings[0], ExtData) = true then
            begin
              //신규 등록 정보로 변경
              Sheet.Cells[i+2, 1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', EnterDate);
              Sheet.Cells[i+2, 2] := '"' + ExtData.MemberId + '"';
              Sheet.Cells[i+2, 3] := ExtData.Name;
              Sheet.Cells[i+2, 4] := ExtData.BaptismalName;
              Sheet.Cells[i+2, 5] := ExtData.TelNo;
              Sheet.Cells[i+2, 6] := ExtData.Sex;
              Sheet.Cells[i+2, 7] := ExtData.CurrentLevel;
              Sheet.Cells[i+2, 8] := ExtData.BirthDay;
              Sheet.Cells[i+2, 9] := ExtData.Address;
            end else
            begin
              //기준등록DB에 없어서 기존 Data 사용
              Sheet.Cells[i+2, 1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', EnterDate);
              Sheet.Cells[i+2, 2] := '"' + tList.Strings[0] + '"';
              Sheet.Cells[i+2, 3] := '(' + Name + ')';
              Sheet.Cells[i+2, 4] := BaptismalName;
              Sheet.Cells[i+2, 5] := TelNo;
              Sheet.Cells[i+2, 6] := Sex;
              Sheet.Cells[i+2, 7] := tList.Strings[2];
              Sheet.Cells[i+2, 8] := tList.Strings[3];
              Sheet.Cells[i+2, 9] := tList.Strings[1];
            end;
          end else
          begin
            //QrCode
            Sheet.Cells[i+2, 1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', EnterDate);
            Sheet.Cells[i+2, 2] := '';
            Sheet.Cells[i+2, 3] := Name;
            Sheet.Cells[i+2, 4] := BaptismalName;
            Sheet.Cells[i+2, 5] := TelNo;
            Sheet.Cells[i+2, 6] := Sex;
            Sheet.Cells[i+2, 7] := '';
            Sheet.Cells[i+2, 8] := '';
            Sheet.Cells[i+2, 9] := Address;
          end;

          DM.FDMysqlQuery2.Next;
        end;
      end;

      XLApp.WorkBooks[1].SaveAs(sFileName);   //Save excel file
      //moReport.Lines.add('엑셀 파일 저장 완료');

      if not VarIsEmpty(XLApp) then
      begin
        XLApp.DisplayAlerts := False;  // Discard unsaved files....
        XLApp.Quit;
      end;

      moReport.Lines.add(format('[%s] Report 출력 완료 : %d 건', [ReportDate, View1Count + View2Count]))
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('ExcptionType=%s', [e.Message]);
        moReport.Lines.add(LogData);

        exit;
      end;
    end;
  finally
    tList.Free;
  end;
end;




procedure TfmReport.edIdNoKeyPress(Sender: TObject; var Key: Char);
var
  BarCodeId, LogData : string;
  ExtData : TInputData;
begin
  if Key = #13 then
  begin
    edQueryName.Text := '';
    edQueryBaptismalName.Text := '';
    edQueryTelNo.Text := '';
    cbQuerySex.Text := '';
    edQueryAddress.Text := '';
    edQueryLevel.Text := '';
    edQueryBirthDay.Text := '';

    BarCodeId := Trim(edIdNo.text);

    if Length(BarCodeId) <> 13 then
    begin
      ShowMessage('교구 BarCode ID  자리수가 상이 합니다');
      exit;
    end;

    if fmDisplayMain.GetMemberInfo(BarCodeId, ExtData) = false then
    begin
      ShowMessage('등록되지 않은 교구 ID 입니다 : ' + BarCodeId);
      exit;
    end;

    edQueryName.Text := ExtData.Name;
    edQueryBaptismalName.Text := ExtData.BaptismalName;
    edQueryTelNo.Text := ExtData.TelNo;
    cbQuerySex.Text := ExtData.Sex;
    edQueryAddress.Text :=  ExtData.Address;
    edQueryLevel.Text := ExtData.CurrentLevel;
    edQueryBirthDay.Text := ExtData.BirthDay;
  end;
end;



function TfmReport.NewDayReportOutputFile: Boolean;
var
  List1, tList : TStringList;
  i : Integer;
  str, sFileName : string;
  Sheet: Variant;
  RowRange, ColumnRange: Variant;
begin
  tList := TStringList.Create;
  List1 := TStringList.Create;

  OpenDialog2.InitialDir := cDir + '\CheckHistory\';

  try
    if OpenDialog2.Execute then
    begin
      str := ExtractFileName(OpenDialog2.FileName);
      str := Copy(str, 1, Length(str)-4);

      lbReportStatus.Caption := format('%s FileLoading...', [str]);
      tList.LoadFromFile(OpenDialog2.FileName);
      sFileName := cDir + '\Report\' + str + '_변환.xlsx';
      XLApp := CreateOleObject('Excel.Application');
      XLApp.Visible := false;
      XLApp.Workbooks.Add(xlWBatWorkSheet);
      XLApp.Workbooks[1].WorkSheets[1].Name := '기간참석자';
      Sheet := XLApp.Workbooks[1].WorkSheets['기간참석자'];
      ColumnRange := XLApp.Workbooks[1].WorkSheets['기간참석자'].Columns;

      ColumnRange.Columns[1].ColumnWidth := 20;
      ColumnRange.Columns[2].ColumnWidth := 10;
      ColumnRange.Columns[3].ColumnWidth := 10;
      ColumnRange.Columns[4].ColumnWidth := 15;
      ColumnRange.Columns[5].ColumnWidth := 10;
      ColumnRange.Columns[6].ColumnWidth := 100;

      Sheet.Cells[1, 1] := '날자';
      Sheet.Cells[1, 2] := '성명';
      Sheet.Cells[1, 3] := '세례명';
      Sheet.Cells[1, 4] := '전화번호';
      Sheet.Cells[1, 5] := '성별';
      Sheet.Cells[1, 6] := '주소/교구 추가 정보';

      for i := 0 to tList.Count-1 do
      begin
        lbReportStatus.Caption := format('[%d / %d] Data Converting...', [i+1, tList.Count]);

        str := tList.Strings[i];
        SepString(str, ',', List1);

        if List1.Count = 3 then
        begin
          Sheet.Cells[i+2, 1] := List1[0];
          Sheet.Cells[i+2, 2] := List1[1];
          Sheet.Cells[i+2, 6] := List1[2];
        end else
        begin
          Sheet.Cells[i+2, 1] := List1[0];
          Sheet.Cells[i+2, 2] := List1[1];
          Sheet.Cells[i+2, 3] := List1[2];
          Sheet.Cells[i+2, 4] := DecryptString(List1[3]);
          Sheet.Cells[i+2, 5] := List1[4];
          Sheet.Cells[i+2, 6] := DecryptString(List1[5]);
        end;
      end;

      XLApp.WorkBooks[1].SaveAs(sFileName);   //Save excel file
      if not VarIsEmpty(XLApp) then
      begin
        XLApp.DisplayAlerts := False;  // Discard unsaved files....
        XLApp.Quit;
      end;

      moReport.Lines.add(format('[%s] Report 출력 완료', [ExtractFileName(sFileName)]))
    end;
  finally
    tList.free;
    List1.free;
  end;
end;





procedure TfmReport.SepString(OrgStr, Sep: string; tList: TStringList);
var
  i, Count : Integer;
  oStr, str : string;
begin
  tList.clear;

  oStr := OrgStr;

  repeat
    i := Pos(Sep, oStr);

    if i >= 1 then
    begin
      str := Copy(oStr, 1, i-1);
      Delete(oStr, 1, i);
      tList.Add(str);
    end else
    begin
      tList.Add(oStr);
      break;
    end;
  until str = '';
end;



end.
