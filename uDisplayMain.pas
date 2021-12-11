unit uDisplayMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  System.DateUtils, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  CPort, Vcl.ComCtrls, Vcl.Grids, Vcl.ExtCtrls, Vcl.CheckLst, Vcl.WinXPickers,
  Grijjy.TextToSpeech, ComObj, JvExExtCtrls, uCommonFunction, System.Diagnostics;

type

  TfmDisplayMain = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    sgData: TStringGrid;
    ComPort1: TComPort;
    pnDataInsert: TPanel;
    Label1: TLabel;
    edName: TEdit;
    Label2: TLabel;
    edbaptismalName: TEdit;
    Label3: TLabel;
    edTelNo: TEdit;
    Label4: TLabel;
    edAddress: TEdit;
    Label5: TLabel;
    edSex: TEdit;
    pnQueryInsert: TPanel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    edQueryName: TEdit;
    edQueryBaptismalName: TEdit;
    edQueryTelNo: TEdit;
    edQueryAddress: TEdit;
    btQuery: TButton;
    cbQuerySex: TComboBox;
    ckQueryList: TCheckListBox;
    moQueryStatus: TMemo;
    Label12: TLabel;
    Timer1: TTimer;
    btManualReg: TButton;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Label18: TLabel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    Panel11: TPanel;
    lbTtsStatus: TLabel;
    lbDataStatus: TLabel;
    Label11: TLabel;
    edExtData: TEdit;
    tmCheckDb: TTimer;
    procedure FormShow(Sender: TObject);
    procedure btQueryClick(Sender: TObject);
    procedure ComPort1RxChar(Sender: TObject; Count: Integer);
    procedure Timer1Timer(Sender: TObject);
    procedure btManualRegClick(Sender: TObject);
    procedure ckQueryListClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure tmCheckDbTimer(Sender: TObject);
  private
    { Private declarations }
    FTextToSpeech: IgoTextToSpeech;
    procedure Log(const AMsg: String);
    procedure TextToSpeechAvailable(Sender: TObject);
    procedure TextToSpeechStarted(Sender: TObject);
    procedure TextToSpeechFinished(Sender: TObject);
    procedure UpdateControls;
  public
    { Public declarations }
    function SaveToServer_AttendMass(Data : TInputData; Memo : TLabel) : Integer;
    procedure AddCheckData(Data : TInputData);
    function SaveToFileDataHistory(Data: TInputData): Boolean;
    function SaveToFileDataHistory_DontoSaveDatabase(Data: TInputData): Boolean;
    function SaveToFileDataHistory_DontoSaveMySqlDatabase(Data: TInputData): Boolean;
    function SaveToTextFileDataHistory(FileName : string; Data: TInputData): Boolean;
    function UpdateMassCount : Boolean;
    function GetMemberInfo(MemberId : string; var Data : TInputData) : boolean;
  end;

var
  fmDisplayMain: TfmDisplayMain;

//ckPcbCodeList.Checked[No] = false
//Copy(ckPcbCodeList.Items.Strings[No], 1, 11);

implementation
uses
  uDataModule, XLConst, uSplitView, uReport, uSetUp, uLog;

{$R *.dfm}

procedure TfmDisplayMain.AddCheckData(Data: TInputData);
var
  NewTelNo : String;
  i : Integer;
begin
  NewTelNo := Data.TelNo;
  for i := 1 to Length(NewTelNo)-4 do
    NewTelNo[i] := '*';

  sgData.RowCount := sgData.RowCount+1;
  sgData.Cells[0, sgData.RowCount-1] := sgData.RowCount.ToString;
  sgData.Cells[1, sgData.RowCount-1] := FormatDateTime('YYYY-MM-DD hh:nn:ss', Data.EnterDate);
  sgData.Cells[2, sgData.RowCount-1] := Data.Name;
  sgData.Cells[3, sgData.RowCount-1] := Data.BaptismalName;
  sgData.Cells[4, sgData.RowCount-1] := NewTelNo;
  sgData.Cells[5, sgData.RowCount-1] := Data.Sex;
  sgData.Cells[6, sgData.RowCount-1] := '*******';

  sgData.Row := sgData.RowCount-1;

  //sgData.AddRow;
  //sgData.Row := SgIndexNo;
end;

procedure TfmDisplayMain.btManualRegClick(Sender: TObject);
var
  Data, ExtData : TInputData;
  ret, i, DataLength : Integer;
  str, TtsText, ExtId : string;
begin
  btManualReg.Enabled := false;
  fmLog.DeviceLogOut(1, 0, -1, 'btManualRegClick', '--', 'Start to Function');

  try
    ExtId := Trim(edExtData.Text);
    if Length(ExtId) = 13 then
    begin
      if GetMemberInfo(ExtId, ExtData) = false then
      begin
        //if GetMemberInfo(ExtId, ExtData) = false then
        //begin
          ShowMessage('등록되지 않은 교구 ID 입니다, 사무실로 방문 바랍니다 : ' + ExtId);

          pnDataInsert.Color := clRed;
          exit;
        //end;
      end;

      Data.MemberId := ExtData.MemberId;
      Data.Name := ExtData.Name;
      Data.BaptismalName := ExtData.BaptismalName;
      Data.TelNo := ExtData.TelNo;
      Data.Sex := ExtData.Sex;
      Data.Address :=  ExtData.MemberId + '~' + ExtData.Address + '~' + ExtData.CurrentLevel + '~' + ExtData.BirthDay;

      if (Data.Name='')or(Data.BaptismalName='') then
      begin
        ShowMessage('입력 안한 정보가 있습니다,  입력을 완료후 다시 등록 바랍니다');
        exit;
      end;
    end else
    begin
      Data.MemberId := '';
      Data.Name := Trim(edQueryName.Text);
      Data.BaptismalName := Trim(edQueryBaptismalName.Text);
      Data.TelNo := Trim(edQueryTelNo.Text);
      Data.Sex := Trim(cbQuerySex.Text);
      Data.Address := Trim(edQueryAddress.Text);
      Data.CurrentLevel := '';
      Data.BirthDay := '';

      Data.TelNo.Replace('-', '');
      Data.Address.Replace(',', ' ');

      if Pos('-', Data.TelNo) >= 1 then
      begin
        ShowMessage(format('에러 : %s님의 전화번호에 "-"이 포함 되어 있습니다', [Name]));
        exit;
      end;

      if (Data.Name='')or(Data.BaptismalName='')or(Data.TelNo='')or(Data.Sex='')or(Data.Address='') then
      begin
        ShowMessage('입력 안한 정보가 있습니다,  입력을 완료후 다시 등록 바랍니다');
        exit;
      end;
    end;

    Data.EnterDate := Now;

    ret := SaveToServer_AttendMass(Data, lbDataStatus);
    //if ret = 0 then
    //  ret := SaveToServer_AttendMass(Data, lbDataStatus);

    case ret of
      0 : pnQueryInsert.Color := clRed;
      1 : pnQueryInsert.Color := clLime;
      2 : pnQueryInsert.Color := clTeal;
    end;

    UpdateMassCount;

  finally
    edName.Text := '';
    edbaptismalName.Text := '';
    edTelNo.Text := '';
    edSex.Text := '';
    edAddress.Text := '';
    edExtData.Text := '';
    pnDataInsert.Color := clBtnFace;
    btManualReg.Enabled := true;
  end;
end;

procedure TfmDisplayMain.btQueryClick(Sender: TObject);
var
  Name, BaptismalName, TelNo, Sex, Address, LogData : string;
  i : Integer;
begin
  pnQueryInsert.Color := clBtnFace;
  moQueryStatus.Clear;
  ckQueryList.Clear;

  Name := Trim(edQueryName.text);
  edQueryBaptismalName.Text := '';
  edQueryTelNo.text := '';
  edQueryAddress.text := '';
  cbQuerySex.ItemIndex := -1;

  moQueryStatus.Lines.Add(format('%s님 성당 참석 현황 조회...', [Name]));

  DM.FDQuery1.Close;
  DM.FDQuery1.SQL.Clear;
  DM.FDQuery1.SQL.Text := 'select * FROM AttendMass Where (Name=:Name) Group by Name, BaptismalName, TelNo, Sex, Address';
  DM.FDQuery1.ParamByName('Name').AsString := Name;
  DM.FDQuery1.Open;

  if DM.FDQuery1.RecordCount = 0 then
  begin
    moQueryStatus.Lines.Add('미사 참례 이력이 없습니다,  수작업으로 입력 바랍니다');
  end else
  begin
    for i := 0 to DM.FDQuery1.RecordCount-1 do
    begin
      BaptismalName := DM.FDQuery1.FieldByName('BaptismalName').AsString;
      TelNo := DM.FDQuery1.FieldByName('TelNo').AsString;
      Sex := DM.FDQuery1.FieldByName('Sex').AsString;
      Address := DM.FDQuery1.FieldByName('Address').AsString;

      ckQueryList.Items.Add(format('%s, %s, %s, %s, %s', [Name, BaptismalName, TelNo, Sex, Address]));
      DM.FDQuery1.Next;
    end;

    if DM.FDQuery1.RecordCount = 1 then
    begin
      LogData := format('%s님 이름으로 미사 참례 이력이 확인 되었습니다', [Name]);

      edQueryBaptismalName.Text := BaptismalName;
      edQueryTelNo.text := telNo;
      edQueryAddress.text := Address;
      cbQuerySex.Text := Sex;

      ckQueryList.Checked[0] := true;
    end else
    begin
      LogData := format('%s님 이름으로 %d 분이 확인 되었습니다,  본인을 선택하여 주십시오', [Name, DM.FDQuery1.RecordCount]);
      moQueryStatus.Lines.Add(LogData);
    end;
  end;
end;

procedure TfmDisplayMain.ckQueryListClick(Sender: TObject);
var
  str : string;
begin
  str := ckQueryList.Items.Strings[ckQueryList.ItemIndex];
  StringParserType(str, ',', tList);

  if tList.Count < 5 then
  begin
    moQueryStatus.Lines.Add('Data가 잘못 되었습니다, 다시선택 또는 수작업으로 입력 바랍니다');
    exit;
  end;

  edQueryName.Text := Trim(tList.Strings[0]);
  edQueryBaptismalName.Text := Trim(tList.Strings[1]);
  edQueryTelNo.Text := Trim(tList.Strings[2]);
  cbQuerySex.Text := Trim(tList.Strings[3]);
  edQueryAddress.Text := Trim(tList.Strings[4]);
end;



procedure TfmDisplayMain.ComPort1RxChar(Sender: TObject; Count: Integer);
var
  rData : AnsiString;//UnicodeString;
  //rData : string;//UnicodeString;
  cData : string;
  i, No : Integer;
begin
  //ComPort1.WriteStr(Data); //WriteStr를 string -> ansiString 으로 젼경해야 제대로 출력이 됨
  //ComPort1.ReadUnicodeString(rData, i);
  ComPort1.ReadStr(rData, i);//이걸로 읽으면 한글이 깨짐
  cData := UTF8ToString(rData);//moQueryStatus.Lines.Add(rData);로 하면 글자 깨짐
  //moQueryStatus.Lines.Add(cData);

  gData := gData + cData;
  No := Length(gData);
  if gData[No] = #$0a then
  begin
    gData := Copy(gData, 1, No-2);
    Timer1.Enabled := true;
  end;
  //아래 Code가 있으면 gData가 없어짐??
  {if gData[No] = #$0a then
  begin
    StringParserType(gData, ',', tList);
    if tList.Count <> 5 then
    begin
      moStatus.Lines.Add('잘못된 BarCode 입니다,  다시한번 읽어 주세요 : ' + cData);
      gData := '';
      exit;
    end;

    edName.Text := Trim(tList.Strings[0]);
    edbaptismalName.Text := Trim(tList.Strings[1]);
    edTelNo.Text := Trim(tList.Strings[2]);
    edSex.Text := Trim(tList.Strings[3]);
    edAddress.Text := Trim(tList.Strings[4]);

    gData := '';
  end;}
end;



procedure TfmDisplayMain.FormCreate(Sender: TObject);
begin
  FTextToSpeech := TgoTextToSpeech.Create;
  FTextToSpeech.OnAvailable := TextToSpeechAvailable;
  FTextToSpeech.OnSpeechStarted := TextToSpeechStarted;
  FTextToSpeech.OnSpeechFinished := TextToSpeechFinished;
end;

procedure TfmDisplayMain.FormShow(Sender: TObject);
begin
  sgData.ColWidths[0] := 40;//No
  sgData.ColWidths[1] := 150;//Date
  sgData.ColWidths[2] := 100;//이름
  sgData.ColWidths[3] := 120;//세레명
  sgData.ColWidths[4] := 120;//전화번호
  sgData.ColWidths[5] := 70;//성별
  sgData.ColWidths[6] := 450;//주소

  sgData.Cells[0, 0] := 'No';
  sgData.Cells[1, 0] := '날자';
  sgData.Cells[2, 0] := '성명';
  sgData.Cells[3, 0] := '세례명';
  sgData.Cells[4, 0] := '전화번호';
  sgData.Cells[5, 0] := '성별';
  sgData.Cells[6, 0] := '주소';

  PageControl1.ActivePage := TabSheet1;
end;


function TfmDisplayMain.GetMemberInfo(MemberId: string;
  var Data: TInputData): boolean;
var
  LogData : string;
begin
  Result := false;
  if SystemOption.SystemType = LocalPcOnly then
  begin
    try
      DM.FDQuery3.Close;
      DM.FDQuery3.SQL.Clear;
      DM.FDQuery3.SQL.Text := 'select * FROM MemberInfo where MemberId=:MemberId';
      DM.FDQuery3.ParamByName('MemberId').AsString := MemberId;
      DM.FDQuery3.Open;

      if DM.FDQuery3.RecordCount = 0 then
        exit;

      Data.MemberId := MemberId;
      Data.Name := DM.FDQuery3.FieldByName('Name').AsString;
      Data.BaptismalName := DM.FDQuery3.FieldByName('BaptismalName').AsString;
      Data.TelNo := DecryptString(DM.FDQuery3.FieldByName('TelNo').AsString);
      Data.BirthDay := DecryptString(DM.FDQuery3.FieldByName('BirthDay').AsString);
      Data.Sex := DM.FDQuery3.FieldByName('Sex').AsString;
      Data.CurrentLevel := DM.FDQuery3.FieldByName('CurrentLevel').AsString;
      Data.Address := DecryptString(DM.FDQuery3.FieldByName('Address').AsString);
      Data.EnterDate := DM.FDQuery3.FieldByName('EnterDate').AsDateTime;

      Result := true;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('ExcptionType=%s', [e.Message]);
        fmLog.DeviceLogOut(1, 0, -1, 'GetMemberInfo', 'Exception', LogData);
        exit;
      end;
    end;
  end else
  begin
    try
      DM.FDMysqlQuery3.Close;
      DM.FDMysqlQuery3.SQL.Clear;
      DM.FDMysqlQuery3.SQL.Text := 'select * FROM MemberInfo where MemberId=:MemberId';
      DM.FDMysqlQuery3.ParamByName('MemberId').AsString := MemberId;
      DM.FDMysqlQuery3.Open;

      if DM.FDMysqlQuery3.RecordCount = 0 then
        exit;

      Data.MemberId := MemberId;
      Data.Name := DM.FDMysqlQuery3.FieldByName('Name').AsString;
      Data.BaptismalName := DM.FDMysqlQuery3.FieldByName('BaptismalName').AsString;
      Data.TelNo := DecryptString(DM.FDMysqlQuery3.FieldByName('TelNo').AsString);
      Data.BirthDay := DecryptString(DM.FDMysqlQuery3.FieldByName('BirthDay').AsString);
      Data.Sex := DM.FDMysqlQuery3.FieldByName('Sex').AsString;
      Data.CurrentLevel := DM.FDMysqlQuery3.FieldByName('CurrentLevel').AsString;
      Data.Address := DecryptString(DM.FDMysqlQuery3.FieldByName('Address').AsString);
      Data.EnterDate := DM.FDMysqlQuery3.FieldByName('EnterDate').AsDateTime;

      Result := true;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        LogData := format('ExcptionType=%s', [e.Message]);
        fmLog.DeviceLogOut(1, 0, -1, 'GetMemberInfo', 'Exception', LogData);
        exit;
      end;
    end;
  end;
end;

procedure TfmDisplayMain.Log(const AMsg: String);
begin
  //MemoLog.Lines.Add(AMsg);
  lbTtsStatus.Caption := AMsg;
end;

function TfmDisplayMain.SaveToFileDataHistory(Data: TInputData): Boolean;
var
  Dir, sFileName, OnDate, StrData : String;
  F : TextFile;
begin
  Dir := cDir + '\CheckHistory\';
  if not DirectoryExists(Dir) then
  begin
    if not ForceDirectories(Dir) then
    begin
      ShowMessage(Format('Error creating our "%s" directory.', [Dir]));
      Exit;
    end;
  end;

  OnDate := FormatDateTime('YYYY-MM-DD', Data.EnterDate);
  sFileName := Dir + OnDate + '.Log';

  AssignFile(F, sFileName);
  try
    if FileExists(sFileName) then
      Append(F)
    else
      Rewrite(F);

      StrData := format('%s,%s,%s,%s,%s,%s', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Data.EnterDate), Data.Name, Data.BaptismalName, EncryptString(Data.TelNo),
                 Data.Sex, EncryptString(Data.Address)]);

      Writeln(F, StrData);
  except
    on e: Exception do
    begin
      CloseFile(F);
      DeviceLogOutNew('SaveToFileDataHistory', 'Exception=' + e.Message, StrData);

      exit;
    end;
  end;

  CloseFile(F);
end;

function TfmDisplayMain.SaveToFileDataHistory_DontoSaveDatabase(
  Data: TInputData): Boolean;
var
  Dir, sFileName, StrData : String;
  F : TextFile;
begin
  Dir := cDir + '\Error\';
  if not DirectoryExists(Dir) then
  begin
    if not ForceDirectories(Dir) then
    begin
      ShowMessage(Format('Error creating our "%s" directory.', [Dir]));
      Exit;
    end;
  end;

  sFileName := Dir + 'DonotSavedServerList.Log';

  AssignFile(F, sFileName);
  try
    if FileExists(sFileName) then
      Append(F)
    else
      Rewrite(F);

      StrData := format('%s,%s,%s,%s,%s,%s', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Data.EnterDate), Data.Name, Data.BaptismalName, EncryptString(Data.TelNo),
                 Data.Sex, EncryptString(Data.Address)]);

      Writeln(F, StrData);
  except
    on e: Exception do
    begin
      CloseFile(F);
      DeviceLogOutNew('SaveToFileDataHistory_DontoSaveDatabase', 'Exception=' + e.Message, StrData);

      exit;
    end;
  end;

  CloseFile(F);
end;

function TfmDisplayMain.SaveToFileDataHistory_DontoSaveMySqlDatabase(
  Data: TInputData): Boolean;
var
  Dir, sFileName, StrData, LogData : String;
  F : TextFile;
begin
  Dir := cDir + '\Error\';
  if not DirectoryExists(Dir) then
  begin
    if not ForceDirectories(Dir) then
    begin
      ShowMessage(Format('Error creating our "%s" directory.', [Dir]));
      Exit;
    end;
  end;

  sFileName := Dir + 'DonotSavedSqlServerList.Log';

  AssignFile(F, sFileName);
  try
    if FileExists(sFileName) then
      Append(F)
    else
      Rewrite(F);

      StrData := format('%s,%s,%s,%s,%s,%s', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Data.EnterDate), Data.Name, Data.BaptismalName, EncryptString(Data.TelNo),
                 Data.Sex, EncryptString(Data.Address)]);

      Writeln(F, StrData);
  except
    on e: Exception do
    begin
      CloseFile(F);
      DeviceLogOutNew('SaveToFileDataHistory_DontoSaveMySqlDatabase', 'Exception=' + e.Message, StrData);

      exit;
    end;
  end;

  CloseFile(F);

  //Table에 저장
  try
    DM.FDConnection1.StartTransaction;
    DM.FDQuery1.Close;
    DM.FDQuery1.SQL.Clear;

    DM.FDQuery1.SQL.Text := 'INSERT INTO AttendMass_Ext (Name, BaptismalName, TelNo, Sex, Address, EnterDate)' +
                   ' VALUES(' +
                   ' :Name,'+
                   ' :BaptismalName,'+
                   ' :TelNo,'+
                   ' :Sex,'+
                   ' :Address,'+
                   ' :EnterDate)';

    DM.FDQuery1.ParamByName('Name').AsString := Data.Name;
    DM.FDQuery1.ParamByName('BaptismalName').AsString := Data.BaptismalName;
    DM.FDQuery1.ParamByName('TelNo').AsString := EncryptString(Data.TelNo);
    DM.FDQuery1.ParamByName('Sex').AsString := Data.Sex;
    DM.FDQuery1.ParamByName('Address').AsString := EncryptString(Data.Address);
    DM.FDQuery1.ParamByName('EnterDate').AsDateTime := Data.EnterDate;
    DM.FDQuery1.ExecSQL;
    DM.FDConnection1.Commit;
  except
    on e: Exception do //ShowMessage(e.Message);
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      fmLog.DeviceLogOut(1, 0, -1, 'SaveToFileDataHistory_DontoSaveMySqlDatabase', 'Exception', LogData);
    end;
  end;
end;

function TfmDisplayMain.SaveToServer_AttendMass(Data: TInputData; Memo : TLabel): Integer;
var
  LogData, TtsText, SexStr : String;
  cTime, EnterTime : TDateTime;
  CheckTime : Int64;
  i : Integer;
  sw: TStopWatch;
begin
  Result := 0;//0=입력불량, 1=입력 완료, 2=기등록
  sw := TStopwatch.StartNew;
  cTime := Now;
  //OpTime := ((SecondsBetween(MaxDate, MinDate)/60)/60);
  //StrToDateTIme(FormatDateTime('YYYY-MM-DD 00:00:00', sTime));
  //SubStartDay := FormatDateTime(format('YYYY-MM-DD %2d:00:00', [SubNo]), sTime);
  //SubEndDay := FormatDateTime(format('YYYY-MM-DD %2d:59:59', [SubNo]), sTime);

  fmLog.DeviceLogOut(1, 0, sw.ElapsedMilliseconds, 'SaveToServer_AttendMass', '--', 'Start to Function');

  if Length(EncryptString(Data.TelNo)) > 49 then
  begin
    TtsText := '전화번호를 확인해 주십시오(Length 초과)';

    if (not FTextToSpeech.Speak(TtsText)) then
      Log('Unable to speak text');
    exit;
  end;

  if Length(EncryptString(Data.Address)) > 254 then
  begin
    TtsText := '주소를 확인해 주십시오(Length 초과)';

    if (not FTextToSpeech.Speak(TtsText)) then
      Log('Unable to speak text');
    exit;
  end;

  if Data.MemberId = '' then
  begin
    SexStr := Data.Sex;
  end else
  begin
    if Data.Sex = '남' then
      SexStr := '형제'
    else
      SexStr := '자매';
  end;

  if SystemOption.SystemType = LocalPcOnly then
  begin
    try
      DM.FDQuery1.Close;
      DM.FDQuery1.SQL.Clear;
      DM.FDQuery1.SQL.Text := 'select * FROM AttendMass Where (Name=:Name)and(BaptismalName=:BaptismalName)and(TelNo=:TelNo) order by EnterDate Desc';
      DM.FDQuery1.ParamByName('Name').AsString := Data.Name;
      DM.FDQuery1.ParamByName('BaptismalName').AsString := Data.BaptismalName;
      DM.FDQuery1.ParamByName('TelNo').AsString := EncryptString(Data.TelNo);
      DM.FDQuery1.Open;

      if DM.FDQuery1.RecordCount >= 1 then
      begin
        EnterTime := DM.FDQuery1.FieldByName('EnterDate').AsDateTime;

        //if cTime < (EnterTime + (1/24)) then
        if cTime < (EnterTime + ((SystemOption.ReCheckLimitTime/60)/24)) then
        begin
          Result := 2;
          CheckTime := MinutesBetween(EnterTime, Now);

          Memo.Caption := format('[%s] %s %s님은 %d 분전에 출석확인 했습니다(%d 분후 재등록 가능), 등록시간=%s',
                                [FormatDateTime('YYYY-MM-DD hh:nn:ss', Now), Data.Name, Data.BaptismalName, CheckTime, SystemOption.ReCheckLimitTime, DateTimeToStr(EnterTime)]);

          TtsText := format('%s %s %s님은 %d 분전에 출석확인 했습니다', [Data.Name, Data.BaptismalName, SexStr, CheckTime]);
          fmLog.DeviceLogOut(1, 0, -1, 'SaveToServer_AttendMass', 'OK', TtsText);

          if (not FTextToSpeech.Speak(TtsText)) then
            Log('Unable to speak text');

          exit;
        end;
      end;

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

      DM.FDQuery1.ParamByName('Name').AsString := Data.Name;
      DM.FDQuery1.ParamByName('BaptismalName').AsString := Data.BaptismalName;
      DM.FDQuery1.ParamByName('TelNo').AsString := EncryptString(Data.TelNo);
      DM.FDQuery1.ParamByName('Sex').AsString := Data.Sex;
      DM.FDQuery1.ParamByName('Address').AsString := EncryptString(Data.Address);
      DM.FDQuery1.ParamByName('EnterDate').AsDateTime := Data.EnterDate;
      DM.FDQuery1.ExecSQL;
      DM.FDConnection1.Commit;

      SaveToFileDataHistory(Data);
      AddCheckData(Data);

      LogData := format('[SqLite] %s %s 미사 참례 입력 완료', [Data.Name, Data.BaptismalName]);
      fmLog.DeviceLogOut(1, 0, sw.ElapsedMilliseconds, 'SaveToServer_AttendMass', 'OK', LogData);

      Memo.Caption := (format('[%s] %s %s 미사 참례 입력 완료', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Now), Data.Name, Data.BaptismalName]));

      TtsText := format('%s %s %s님 반갑습니다', [Data.Name, Data.BaptismalName, SexStr]);

      if (not FTextToSpeech.Speak(TtsText)) then
        Log('Unable to speak text');

      Inc(CurrentDayCount);

      Result := 1;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        SaveToFileDataHistory_DontoSaveDatabase(Data);

        LogData := format('Data Base 입력 Error, ExcptionType=%s', [e.Message]);
        fmLog.DeviceLogOut(1, 0, -1, 'SaveToServer_AttendMass', 'Exception', LogData);

        Memo.Caption := format('[%s] Data Base 입력 Error : %s', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Now), e.Message]);

        TtsText := '시스템 문제입니다, 관리자에게 문의 바랍니다';

        if (not FTextToSpeech.Speak(TtsText)) then
          Log('Unable to speak text');

        exit;
      end;
    end;
  end else
  begin
    try
      DM.FDMysqlQuery1.Close;
      DM.FDMysqlQuery1.SQL.Clear;
      DM.FDMysqlQuery1.SQL.Text := 'select * FROM AttendMass Where (Name=:Name)and(BaptismalName=:BaptismalName)and(TelNo=:TelNo) order by EnterDate Desc';
      DM.FDMysqlQuery1.ParamByName('Name').AsString := Data.Name;
      DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := Data.BaptismalName;
      DM.FDMysqlQuery1.ParamByName('TelNo').AsString := EncryptString(Data.TelNo);
      DM.FDMysqlQuery1.Open;

      if DM.FDMysqlQuery1.RecordCount >= 1 then
      begin
        EnterTime := DM.FDMysqlQuery1.FieldByName('EnterDate').AsDateTime;

        //if cTime < (EnterTime + (1/24)) then
        if cTime < (EnterTime + ((SystemOption.ReCheckLimitTime/60)/24)) then
        begin
          Result := 2;
          CheckTime := MinutesBetween(EnterTime, Now);

          Memo.Caption := format('[%s] %s %s님은 %d 분전에 출석확인 했습니다(%d 분후 재등록 가능), 등록시간=%s',
                                [FormatDateTime('YYYY-MM-DD hh:nn:ss', Now), Data.Name, Data.BaptismalName, CheckTime, SystemOption.ReCheckLimitTime, DateTimeToStr(EnterTime)]);

          TtsText := format('%s %s %s님은 %d 분전에 출석확인 했습니다', [Data.Name, Data.BaptismalName, SexStr, CheckTime]);
          fmLog.DeviceLogOut(1, 0, -1, 'SaveToServer_AttendMass', 'OK', TtsText);

          if (not FTextToSpeech.Speak(TtsText)) then
            Log('Unable to speak text');

          exit;
        end;
      end;

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

      DM.FDMysqlQuery1.ParamByName('Name').AsString := Data.Name;
      DM.FDMysqlQuery1.ParamByName('BaptismalName').AsString := Data.BaptismalName;
      DM.FDMysqlQuery1.ParamByName('TelNo').AsString := EncryptString(Data.TelNo);
      DM.FDMysqlQuery1.ParamByName('Sex').AsString := Data.Sex;
      DM.FDMysqlQuery1.ParamByName('Address').AsString := EncryptString(Data.Address);
      DM.FDMysqlQuery1.ParamByName('EnterDate').AsDateTime := Data.EnterDate;
      DM.FDMysqlQuery1.ExecSQL;
      DM.FDConMysql.Commit;

      SaveToFileDataHistory(Data);
      AddCheckData(Data);

      LastDbCheckTime := Now;

      LogData := format('[MySql] %s %s 미사 참례 입력 완료', [Data.Name, Data.BaptismalName]);
      fmLog.DeviceLogOut(1, 0, sw.ElapsedMilliseconds, 'SaveToServer_AttendMass', 'OK', LogData);

      Memo.Caption := (format('[%s] %s %s 미사 참례 입력 완료', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Now), Data.Name, Data.BaptismalName]));

      TtsText := format('%s %s %s님 반갑습니다', [Data.Name, Data.BaptismalName, SexStr]);

      if (not FTextToSpeech.Speak(TtsText)) then
        Log('Unable to speak text');

      Inc(CurrentDayCount);

      Result := 1;
    except
      on e: Exception do //ShowMessage(e.Message);
      begin
        SaveToFileDataHistory_DontoSaveMySqlDatabase(Data);

        LogData := format('Data Base 입력 Error, ExcptionType=%s', [e.Message]);
        fmLog.DeviceLogOut(1, 0, -1, 'SaveToServer_AttendMass', 'Exception', LogData);

        Memo.Caption := format('[%s] Data Base 입력 Error : %s', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Now), e.Message]);

        TtsText := '시스템 문제입니다, 관리자에게 문의 바랍니다';

        if (not FTextToSpeech.Speak(TtsText)) then
          Log('Unable to speak text');

        exit;
      end;
    end;
  end;
end;


function TfmDisplayMain.SaveToTextFileDataHistory(FileName : string; Data: TInputData): Boolean;
var
  Dir, OnDate, StrData : String;
  F : TextFile;
begin
  Dir := cDir + '\Report\';
  if not DirectoryExists(Dir) then
  begin
    if not ForceDirectories(Dir) then
    begin
      ShowMessage(Format('Error creating our "%s" directory.', [Dir]));
      Exit;
    end;
  end;

  AssignFile(F, FileName);
  try
    if FileExists(FileName) then
      Append(F)
    else
      Rewrite(F);

      StrData := format('%s,%s,%s,%s,%s,%s', [FormatDateTime('YYYY-MM-DD hh:nn:ss', Data.EnterDate), Data.Name, Data.BaptismalName, Data.TelNo,
                 Data.Sex, Data.Address]);
      Writeln(F, StrData);
  except
    on e: Exception do
    begin
      CloseFile(F);
      DeviceLogOutNew('SaveToTextFileDataHistory', 'Exception=' + e.Message, StrData);

      exit;
    end;
  end;

  CloseFile(F);
end;

procedure TfmDisplayMain.TextToSpeechAvailable(Sender: TObject);
begin
  Log('Text-to-Speech engine is available');
  UpdateControls;
end;

procedure TfmDisplayMain.TextToSpeechFinished(Sender: TObject);
begin
  Log('Speech finished');
  UpdateControls;
end;

procedure TfmDisplayMain.TextToSpeechStarted(Sender: TObject);
begin
  Log('Speech started');
  UpdateControls;
end;

procedure TfmDisplayMain.Timer1Timer(Sender: TObject);
var
  Data : TInputData;
  ret, i, DataLen : Integer;
  TtsText, Addr, str, FileName : string;
  tList : TStringList;
begin
  Timer1.Enabled := false;

  PageControl1.ActivePage := TabSheet1;
  if SplitViewForm.lblTitle.Caption <> '미사참례 등록' then
    SplitViewForm.actHomeExecute(Sender);

  tList := TStringList.Create;
  try
    DataLen := Length(gData);
    if DataLen = 13 then
    begin
      if GetMemberInfo(gData, Data) = false then
      begin
        lbDataStatus.Caption := '등록되지 않은 교구 ID 입니다, 사무실로 방문 바랍니다 : ' + gData;

        TtsText := '등록되지 않은 교구 ID 입니다, 사무실로 방문 바랍니다';
        if (not FTextToSpeech.Speak(TtsText)) then
          Log('Unable to speak text');

        pnDataInsert.Color := clRed;
        gData := '';

        exit;
      end;

      edName.Text := Data.Name;
      edbaptismalName.Text := Data.BaptismalName;
      edTelNo.Text := Data.TelNo;
      edSex.Text := Data.Sex;
      edAddress.Text := Data.MemberId + '~' + Data.Address + '~' + Data.CurrentLevel + '~' + Data.BirthDay;

      Data.Address := Data.MemberId + '~' + Data.Address + '~' + Data.CurrentLevel + '~' + Data.BirthDay;
      Data.EnterDate := Now;

      gData := '';

      pnDataInsert.Color := clBtnFace;
      Application.ProcessMessages;

      ret := SaveToServer_AttendMass(Data, lbDataStatus);

      case ret of
        0 : pnDataInsert.Color := clRed;
        1 : pnDataInsert.Color := clLime;
        2 : pnDataInsert.Color := clTeal;
      end;

      UpdateMassCount;
    end else
    begin
      StringParserType(gData, ',', tList);
      if tList.Count < 5 then
      begin
        lbDataStatus.Caption := '잘못된 큐알코드 입니다,  다시한번 읽어 주세요 : ' + gData;

        TtsText := '잘못된 큐알코드 입니다,  다시한번 읽어 주세요';
        if (not FTextToSpeech.Speak(TtsText)) then
          Log('Unable to speak text');

        pnDataInsert.Color := clRed;
        gData := '';
        exit;
      end;

      edName.Text := Trim(tList.Strings[0]);
      edbaptismalName.Text := Trim(tList.Strings[1]);
      edTelNo.Text := Trim(tList.Strings[2]);
      edSex.Text := Trim(tList.Strings[3]);
      str := Trim(tList.Strings[4]);
      str.Replace(',', ' ');
      edAddress.Text := str;

      gData := '';

      pnDataInsert.Color := clBtnFace;
      Application.ProcessMessages;

      Data.MemberId := '';
      Data.Name := Trim(edName.Text);
      Data.BaptismalName := Trim(edBaptismalName.Text);
      Data.TelNo := Trim(edTelNo.Text);
      Data.Sex := Trim(edSex.Text);
      Data.Address := Trim(edAddress.Text);
      Data.CurrentLevel := '';
      Data.BirthDay := '';

      Data.EnterDate := Now;

      ret := SaveToServer_AttendMass(Data, lbDataStatus);

      case ret of
        0 : pnDataInsert.Color := clRed;
        1 : pnDataInsert.Color := clLime;
        2 : pnDataInsert.Color := clTeal;
      end;

      UpdateMassCount;
    end;

    Application.ProcessMessages;

    edName.Text := '';
    edbaptismalName.Text := '';
    edTelNo.Text := '';
    edSex.Text := '';
    edAddress.Text := '';
    pnDataInsert.Color := clBtnFace;
  finally
    tList.Free;
  end;
end;

procedure TfmDisplayMain.tmCheckDbTimer(Sender: TObject);
var
  LogData : string;
  ret : Boolean;
begin
  if now >= (LastDbCheckTime + (2/24)) then //2시간이 넘으면 update
  begin
    ret := DM.FDConMysql.Ping;
    LastDbCheckTime := Now;

    if ret then
      LogData := 'Ping Test OK'
    else
      LogData := 'Ping Test NG';

    fmLog.DeviceLogOut(1, 0, -1, 'tmCheckDbTimer', 'Check', LogData);
  end;
end;

procedure TfmDisplayMain.UpdateControls;
begin
  //ButtonSpeak.Enabled := (not FTextToSpeech.IsSpeaking);
  //ButtonStop.Enabled := (not ButtonSpeak.Enabled);
end;

function TfmDisplayMain.UpdateMassCount: Boolean;
begin
  if CurrentDate <> Date then
  begin
    CurrentDate := Date;
    CurrentDayCount := 1;
  end;

  SplitViewForm.lbCurrentNo.Caption := format('%s : 미사참례 인원 %d ',
                        [FormatDateTime('YYYY-MM-DD', CurrentDate), CurrentDayCount]);
end;


end.

