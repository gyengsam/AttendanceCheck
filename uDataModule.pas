unit uDataModule;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MySQL,
  FireDAC.Phys.MySQLDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, Data.DB, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, uCommonFunction, System.IniFiles, FireDAC.Stan.ExprFuncs,
  FireDAC.Phys.SQLiteDef, FireDAC.Phys.SQLite, Winapi.Windows;

type
  TDM = class(TDataModule)
    FDQuery1: TFDQuery;
    FDQuery2: TFDQuery;
    FDPhysSQLiteDriverLink1: TFDPhysSQLiteDriverLink;
    FDConnection1: TFDConnection;
    FDConMysql: TFDConnection;
    FDMysqlQuery1: TFDQuery;
    FDMysqlQuery2: TFDQuery;
    FDQuery3: TFDQuery;
    FDMysqlQuery3: TFDQuery;
    procedure FDConMysqlAfterConnect(Sender: TObject);
    procedure FDConMysqlAfterDisconnect(Sender: TObject);
    procedure FDConnection1AfterConnect(Sender: TObject);
    procedure FDConnection1AfterDisconnect(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure DbSqLiteConnection;
    procedure DbMySqlConnection;
    function CheckTableName(Data: string; List: TStringList): Boolean;
    function CheckToSqLiteDataBaseTable: Boolean;
    function CheckToMySqlDataBaseTable: Boolean;
    function GetCurrentCount: Integer;
    procedure RunFirstCheckNetworkFun;
  end;

var
  DM: TDM;

implementation
uses
  uLog;

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TDM }

function TDM.CheckTableName(Data: string; List: TStringList): Boolean;
var
  i, No: Integer;
  TableName : string;
begin
  Result := false;

  for i := 0 to List.Count-1 do
  begin
    TableName := Uppercase(List.Strings[i]);
    No := Pos('.', TableName);
    if No >= 1 then
      TableName := Copy(TableName, No+1, No + Length(TableName));

    if Uppercase(Data) = TableName then
    begin
      Result:= true;
      exit;
    end;
  end;
end;

function TDM.CheckToMySqlDataBaseTable: Boolean;
var
  tList : TStringList;
  tName, tColumeName : string;
  Index : Integer;
  ret : Boolean;
begin
  Result := false;

  tList := TStringList.Create;
  try
    //DM.FDConMysql.GetTableNames('','church','', tList);
    DM.FDConMysql.GetTableNames('','dongtan','', tList);

    tName := 'AttendMass';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConMysql.StartTransaction;
      FDMysqlQuery2.Close;
      FDMysqlQuery2.SQL.Clear;
      FDMysqlQuery2.SQL.Text := 'Create Table AttendMass (' +
                      //' Id INTERGER NOT NULL PRIMARY KEY,' +
                      ' Name Varchar(20),' +
                      ' BaptismalName Varchar(30),' +
                      ' TelNo Varchar(50),' +
                      ' Sex Varchar(10),' +
                      ' Address Varchar(255),' +
                      ' EnterDate DateTime)';
      try
        FDMysqlQuery2.ExecSQL;
        FDConMysql.Commit;
      except
        FDConMysql.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToMySqlDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToMySqlDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;

    tName := 'MemberInfo';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConMysql.StartTransaction;
      FDMysqlQuery2.Close;
      FDMysqlQuery2.SQL.Clear;
      FDMysqlQuery2.SQL.Text := 'Create Table MemberInfo (' +
                      ' MemberId Varchar(20) PRIMARY KEY,' +
                      ' Name Varchar(20),' +
                      ' BaptismalName Varchar(30),' +
                      ' TelNo Varchar(50),' +
                      ' BirthDay Varchar(50),' +
                      ' Sex Varchar(10),' +
                      ' CurrentLevel Varchar(20),' +
                      ' Address Varchar(255),' +
                      ' EnterDate DateTime)';
      try
        FDMysqlQuery2.ExecSQL;
        FDConMysql.Commit;
      except
        FDConMysql.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToMySqlDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToMySqlDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;

    {tName := 'MemberInfo_UpdateList';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConMysql.StartTransaction;
      FDMysqlQuery2.Close;
      FDMysqlQuery2.SQL.Clear;
      FDMysqlQuery2.SQL.Text := 'Create Table MemberInfo_UpdateList (' +
                      ' MemberId Varchar(20) PRIMARY KEY,' +
                      ' Name Varchar(20),' +
                      ' BaptismalName Varchar(30),' +
                      ' TelNo Varchar(50),' +
                      ' BirthDay Varchar(50),' +
                      ' Sex Varchar(10),' +
                      ' CurrentLevel Varchar(20),' +
                      ' Address Varchar(255),' +
                      ' EnterDate DateTime)';
      try
        FDMysqlQuery2.ExecSQL;
        FDConMysql.Commit;
      except
        FDConMysql.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToMySqlDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToMySqlDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;}
  finally
    tList.Free;
  end;

  Result := true;
end;

function TDM.CheckToSqLiteDataBaseTable: Boolean;
var
  tList : TStringList;
  tName, tColumeName : string;
  Index : Integer;
  ret : Boolean;
begin
  Result := false;

  tList := TStringList.Create;
  try
    DM.FDConnection1.GetTableNames('','','', tList);

    tName := 'AttendMass';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConnection1.StartTransaction;
      FDQuery1.Close;
      FDQuery1.SQL.Clear;
      FDQuery1.SQL.Text := 'Create Table AttendMass (' +
                      //' Id INTERGER NOT NULL PRIMARY KEY,' +
                      //' Id INTERGER PRIMARY KEY,' +//   기존 table 삭제후 사용
                      ' Name Varchar(20),' +
                      ' BaptismalName Varchar(30),' +
                      ' TelNo Varchar(50),' +
                      ' Sex Varchar(10),' +
                      ' Address Varchar(255),' +
                      ' EnterDate DateTime)';
      try
        FDQuery1.ExecSQL;
        FDConnection1.Commit;
      except
        FDConnection1.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;

    tName := 'AttendMass_Ext';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConnection1.StartTransaction;
      FDQuery1.Close;
      FDQuery1.SQL.Clear;
      FDQuery1.SQL.Text := 'Create Table AttendMass_Ext (' +
                      ' Name Varchar(20),' +
                      ' BaptismalName Varchar(30),' +
                      ' TelNo Varchar(50),' +
                      ' Sex Varchar(10),' +
                      ' Address Varchar(255),' +
                      ' EnterDate DateTime)';
      try
        FDQuery1.ExecSQL;
        FDConnection1.Commit;
      except
        FDConnection1.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;

    tName := 'SystemInfo';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConnection1.StartTransaction;
      FDQuery1.Close;
      FDQuery1.SQL.Clear;
      FDQuery1.SQL.Text := 'Create Table SystemInfo (' +
                      ' Item Varchar(20),' +
                      ' Data Varchar(255))';
      try
        FDQuery1.ExecSQL;
        FDConnection1.Commit;
      except
        FDConnection1.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;

    tName := 'MemberInfo';
    ret := CheckTableName(tName, tList);
    if ret = false then
    begin
      FDConnection1.StartTransaction;
      FDQuery1.Close;
      FDQuery1.SQL.Clear;
      FDQuery1.SQL.Text := 'Create Table MemberInfo (' +
                      ' MemberId Varchar(20) PRIMARY KEY,' +
                      ' Name Varchar(20),' +
                      ' BaptismalName Varchar(30),' +
                      ' TelNo Varchar(50),' +
                      ' BirthDay Varchar(50),' +
                      ' Sex Varchar(10),' +
                      ' CurrentLevel Varchar(20),' +
                      ' Address Varchar(255),' +
                      ' EnterDate DateTime)';
      try
        FDQuery1.ExecSQL;
        FDConnection1.Commit;
      except
        FDConnection1.Rollback;
        //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'Exception', 'RegInspectionSpec, TableCreate');
        exit;
      end;
      //DeviceLogOutNew(-1, 0, 'CheckToSqLiteDataBaseTable', 'OK', 'RegInspectionSpec, TableCreate');
    end;
  finally
    tList.Free;
  end;

  Result := true;
end;

procedure TDM.DbMySqlConnection;
begin
  //MySql3
  if SystemOption.ServerIpAddress <> '' then
  begin
    PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mServer), 2);//w, l

    if FDConMysql.Connected  then
      FDConMysql.Close;

    FDConMysql.Params.Values['Server'] := SystemOption.ServerIpAddress;
    FDConMysql.Params.Values['DataBase'] := SystemOption.DataBaseName;//'church';

    try
      FDConMysql.Open;
    except
      PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mServer), 0);//w, l
      exit;
    end;
  end;
end;

procedure TDM.DbSqLiteConnection;
begin
  //SqLite3
  if FDConnection1.Connected  then
    FDConnection1.Close;

  FDConnection1.Params.Values['DriverID'] := 'SQLite';
  FDConnection1.Params.Values['Database'] := Sql3DbName;
  FDConnection1.Params.Values['UserName'] := 'key_server';
  FDConnection1.Params.Values['Password'] := 'test5290';

  try
    FDConnection1.Open;
  except
    PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mSqLite), 0);//w, l
    exit;
  end;
end;


procedure TDM.FDConMysqlAfterConnect(Sender: TObject);
var
  LogData : string;
begin
  PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mServer), 1);//w, l
  LogData := 'MySql Connected';
  fmLog.DeviceLogOut(1, 0, -1, 'FDConMysqlAfterConnect', 'OK', LogData);

  LastDbCheckTime := Now;
end;

procedure TDM.FDConMysqlAfterDisconnect(Sender: TObject);
var
  LogData : string;
begin
  PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mServer), 0);//w, l
  LogData := 'MySql DisConnected';
  fmLog.DeviceLogOut(1, 0, -1, 'FDConMysqlAfterDisconnect', 'OK', LogData);
end;

procedure TDM.FDConnection1AfterConnect(Sender: TObject);
var
  LogData : string;
begin
  PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mSqLite), 1);//w, l
  LogData := 'SqLite Connected';
  fmLog.DeviceLogOut(1, 0, -1, 'FDConnection1AfterConnect', 'OK', LogData);
end;

procedure TDM.FDConnection1AfterDisconnect(Sender: TObject);
var
  LogData : string;
begin
  PostMessage(mHandle, WM_MSG_DISPLAY_UPDATE, Ord(mSqLite), 0);//w, l
  LogData := 'SqLite DisConnected';
  fmLog.DeviceLogOut(1, 0, -1, 'FDConnection1AfterDisconnect', 'OK', LogData);
end;

function TDM.GetCurrentCount: Integer;
var
  LogData : string;
  sTime, eTime : string;
begin
  Result := 0;

  sTime := FormatDateTime('YYYY-MM-DD 00:00:00', Now);
  eTime := FormatDateTime('YYYY-MM-DD 23:59:59', Now);

  try
    if SystemOption.SystemType = LocalPcOnly then
    begin
      FDQuery1.Close;
      FDQuery1.SQL.Clear;
      FDQuery1.SQL.Text := 'select * FROM AttendMass Where (EnterDate>=:sTime)and(EnterDate<=:eTime) order by EnterDate asc';
      FDQuery1.ParamByName('sTime').AsDateTime := StrToDateTime(sTime);
      FDQuery1.ParamByName('eTime').AsDateTime := StrToDateTime(eTime);
      FDQuery1.Open;

      Result := FDQuery1.RecordCount;
    end else
    begin
      //MySql
      FDMysqlQuery1.Close;
      FDMysqlQuery1.SQL.Clear;
      FDMysqlQuery1.SQL.Text := 'select * FROM AttendMass Where (EnterDate>=:sTime)and(EnterDate<=:eTime) order by EnterDate asc';
      FDMysqlQuery1.ParamByName('sTime').AsDateTime := StrToDateTime(sTime);
      FDMysqlQuery1.ParamByName('eTime').AsDateTime := StrToDateTime(eTime);
      FDMysqlQuery1.Open;

      Result := FDMysqlQuery1.RecordCount;
    end;
  except
    on e: Exception do
    begin
      LogData := format('ExcptionType=%s', [e.Message]);
      DeviceLogOutNew('GetCurrentCount', 'Excption', LogData);
    end;
  end;
end;


procedure TDM.RunFirstCheckNetworkFun;
begin
  CheckToMySqlDataBaseTable;
end;

end.
