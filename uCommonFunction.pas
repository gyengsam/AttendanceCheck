unit uCommonFunction;

interface
uses
  System.IniFiles, System.Classes, System.SysUtils, FileCtrl, Winapi.Windows, Winapi.Messages;

const
  WM_MSG_DISPLAY_UPDATE = WM_USER + 100;
  MaxPostMsgLength = 1024;
  ID_BIT   =  $200000;       // EFLAGS ID bit

type
  TCPUID   = array[1..4] of Longint;
  TVendor  = array [0..11] of char;


  TInputData = record
    MemberId : string;//교구앱
    Name : string;
    BaptismalName : string;
    TelNo : string;
    BirthDay : string;//교구앱, 암호화
    Sex : string;//교구=남/여,  베르네=형제/자매
    CurrentLevel : string;//교구앱, 암호화
    Address : string;
    EnterDate : TDateTime;
    //NewTelNo : string;//암호화
    //NewAddress : string;//암호화
  end;


  TComportNo = (doNotUse, COM1, COM2, COM3, COM4, COM5, COM6, COM7, COM8, COM9, COM10,
                COM11, COM12, COM13, COM14, COM15, COM16, COM17, COM18, COM19, COM20,
                COM21, COM22, COM23, COM24, COM25, COM26, COM27, COM28, COM29, COM30,
                COM31, COM32, COM33, COM34, COM35, COM36, COM37, COM38, COM39, COM40,
                COM41, COM42, COM43, COM44, COM45, COM46, COM47, COM48, COM49, COM50,
                COM51, COM52, COM53, COM54, COM55, COM56, COM57, COM58, COM59, COM60,
                COM61, COM62, COM63, COM64, COM65, COM66, COM67, COM68, COM69, COM70);

  TMachineStatus = (mSerial, mServer, mSqLite);
  TSystemType = (LocalPcOnly, ServerPc);

  TSystemOption = class(TObject)
    private
      FServerIpAddress : string;
      FDataBaseName : string;
      FSerialPortNo : TComportNo;
      FSystemType : TSystemType;
      FAdminPassWord : string;
      FReCheckLimitTime : Integer;
      FCpuSerial : string;
    public
      constructor Create;
      property ServerIpAddress: string read FServerIpAddress write FServerIpAddress;
      property DataBaseName: string read FDataBaseName write FDataBaseName;
      property SerialPortNo: TComportNo read FSerialPortNo write FSerialPortNo;
      property SystemType: TSystemType read FSystemType write FSystemType;
      property AdminPassWord: string read FAdminPassWord write FAdminPassWord;
      property ReCheckLimitTime: Integer read FReCheckLimitTime write FReCheckLimitTime;
      property CpuSerial: string read FCpuSerial write FCpuSerial;
  end;



function StringParserType(orgStr, SepStr: string; Strings: TStringList): Integer;
function DeviceLogOutNew(FuncType, OkNg, Msg : string): Boolean;
function EncryptString(Data : string) : string;
function DecryptString(Data : string) : string;
function GetToCpuId : string;


var
  cDir : string;
  //pckmodule : HMODULE;
  tList : TStringList;
  gData : string;
  XlRead, XLApp, xlSheet, xlRange : Variant;
  CurrentDayCount : Integer;
  CurrentDate : TDate;
  Sql3DbName : string;
  SystemOption : TSystemOption;
  mHandle : HWND;
  OrgPassword : string;
  LastDbCheckTime : TDateTime;

implementation
uses
  uTPLb_BaseNonVisualComponent,
  uTPLb_Codec,
  uTPLb_CryptographicLibrary,
  uTPLb_Constants;

function EncryptString(Data : string) : string;
var
  EncStr: string;
  CryptographicLibrary: TCryptographicLibrary;
  Codec: TCodec;
begin
  CryptographicLibrary := TCryptographicLibrary.Create(nil);

  Codec := TCodec.Create(nil);

  Codec.CryptoLibrary := CryptographicLibrary;
  Codec.StreamCipherId := BlockCipher_ProgID;
  //Codec.BlockCipherId := Blowfish_ProgId;
  Codec.BlockCipherId := Format(AES_ProgId, [256]);
  Codec.ChainModeId := ECB_ProgId;//CBC_ProgId=일경우 암호화시마다 값이 바뀜;

  Codec.Password := 'test5290~';

  Codec.EncryptString(Data, EncStr, TEncoding.UTF8);
  Result := EncStr;
end;


function DecryptString(Data : string) : string;
var
  DecStr: string;
  CryptographicLibrary: TCryptographicLibrary;
  Codec: TCodec;
begin
  CryptographicLibrary := TCryptographicLibrary.Create(nil);

  Codec := TCodec.Create(nil);

  Codec.CryptoLibrary := CryptographicLibrary;
  Codec.StreamCipherId := BlockCipher_ProgID;
  //Codec.BlockCipherId := Blowfish_ProgId;
  Codec.BlockCipherId := Format(AES_ProgId, [256]);
  Codec.ChainModeId := ECB_ProgId;//CBC_ProgId=일경우 암호화시마다 값이 바뀜;

  Codec.Password := 'test5290~';

  Codec.DecryptString(DecStr, Data, TEncoding.UTF8);
  Result := DecStr;
end;


function StringParserType(orgStr, SepStr: string;
  Strings: TStringList): Integer;
var
  i, No : Integer;
  str, FieldSep, strData : string;
  Det : Boolean;
begin
  i := 0;
  str := orgStr;
  Det := false;

  Strings.Clear;

  repeat
    FieldSep := SepStr;
    No := Pos(FieldSep, str);

    if No >= 1 then
    begin
      Det := true;
      strData := copy(str,1, No - 1);

      Strings.Add(strData);
    end else
    begin
      if Length(str) = 0 then
      begin
        if Det then
          Strings.Add('');

        break;
      end else
      begin
        Strings.Add(str);
        break;
      end;
    end;

    delete(str,1,No);
    Inc(i);
  until false;

  Result := i;
end;

function DeviceLogOutNew(FuncType, OkNg, Msg : string): Boolean;
var
  Dir, sFileName, OnDate, StrData: String;
  F : TextFile;
begin
  Dir := cDir + '\InspLog\';
  if not DirectoryExists(Dir) then
  begin
    if not ForceDirectories(Dir) then
    begin
      //ShowMessage(Format('Error creating our "%s" directory.', [Dir]));
      Exit;
    end;
  end;

  OnDate := FormatDateTime('YYYY-MM-DD', Now);
  sFileName := Dir + OnDate + '.Log';

  AssignFile(F, sFileName);
  try
    if FileExists(sFileName) then
      Append(F)
    else
      Rewrite(F);

      StrData := FormatDateTime('YYYY-MM-DD hh:nn:ss', Now) + #$9 + FuncType + #$9 + OkNg + #$9 + Msg;
      Writeln(F, StrData);
  except
    CloseFile(F);
    exit;
  end;

  CloseFile(F);
end;

{ TSystemOption }

constructor TSystemOption.Create;
begin
  FServerIpAddress := '';
  FDataBaseName := '';
  FSerialPortNo := doNotUse;
  FSystemType := LocalPcOnly;
  FAdminPassWord := EncryptString('test5290');
  FReCheckLimitTime := 10;
  FCpuSerial := '';
end;


function IsCPUID_Available: Boolean; register;
asm
  PUSHFD             {direct access to flags no possible, only via stack}
  POP     EAX        {flags to EAX}
  MOV     EDX,EAX    {save current flags}
  XOR     EAX,ID_BIT {not ID bit}
  PUSH    EAX        {onto stack}
  POPFD              {from stack to flags, with not ID bit}
  PUSHFD             {back to stack}
  POP     EAX        {get back to EAX}
  XOR     EAX,EDX    {check if ID bit affected}
  JZ      @exit      {no, CPUID not availavle}
  MOV     AL,True    {Result=True}
@exit:
end;

function GetCPUID: TCPUID; assembler; register;
asm
  PUSH    EBX        {Save affected register}
  PUSH    EDI
  MOV     EDI,EAX    {@Resukt}
  MOV     EAX,1
  DW      $A20F      {CPUID Command}
  STOSD              {CPUID[1]}
  MOV     EAX,EBX
  STOSD              {CPUID[2]}
  MOV     EAX,ECX
  STOSD              {CPUID[3]}
  MOV     EAX,EDX
  STOSD              {CPUID[4]}
  POP     EDI        {Restore registers}
  POP     EBX
end;

function GetCPUVendor: TVendor; assembler; register;
asm
  PUSH    EBX        {Save affected register}
  PUSH    EDI
  MOV     EDI,EAX    {@Result (TVendor)}
  MOV     EAX,0
  DW      $A20F      {CPUID Command}
  MOV     EAX,EBX
  XCHG    EBX,ECX    {save ECX result}
  MOV     ECX,4
@1:
  STOSB
  SHR     EAX,8
  LOOP    @1
  MOV     EAX,EDX
  MOV     ECX,4
@2:
  STOSB
  SHR     EAX,8
  LOOP    @2
  MOV     EAX,EBX
  MOV     ECX,4
@3:
  STOSB
  SHR     EAX,8
  LOOP    @3
  POP     EDI        {Restore registers}
  POP     EBX
end;

function GetToCpuId : string;
var
  CPUID: TCPUID;
  I: Integer;
  S: TVendor;
begin
  Result := '';

  for I := Low(CPUID) to High(CPUID) do
    CPUID[I] := -1;
  if IsCPUID_Available then
  begin
    CPUID  := GetCPUID;
    //Result := IntToHex(CPUID[1],8) + IntToHex(CPUID[4],8) + IntToHex(CPUID[3],8);
    Result := IntToHex(CPUID[4],8) + IntToHex(CPUID[1],8)// + IntToHex(CPUID[3],8);
    //IntToHex(CPUID[1],8) + IntToHex(CPUID[2],8) +  IntToHex(CPUID[3],8) + IntToHex(CPUID[4],8);
    //SerialNo : IntToHex(CPUID[1],8) + IntToHex(CPUID[4],8) + IntToHex(CPUID[3],8);
    //Result := IntToHex(CPUID[3],8) + IntToHex(CPUID[4],8);
    {Label1.Caption := 'CPUID[1] = ' + IntToHex(CPUID[1],8);
    Label2.Caption := 'CPUID[2] = ' + IntToHex(CPUID[2],8); //실행시마다 자주 변함
    Label3.Caption := 'CPUID[3] = ' + IntToHex(CPUID[3],8);
    Label4.Caption := 'CPUID[4] = ' + IntToHex(CPUID[4],8);
    PValue.Caption := IntToStr(CPUID[1] shr 12 and 3);
    FValue.Caption := IntToStr(CPUID[1] shr 8 and $f);
    MValue.Caption := IntToStr(CPUID[1] shr 4 and $f);
    SValue.Caption := IntToStr(CPUID[1] and $f);
    S := GetCPUVendor;
    Label5.Caption := 'Vendor: ' + S;}
  end;
  //else
  //begin
  //  Result := 'CPUID not available';
  //end;
end;

end.
