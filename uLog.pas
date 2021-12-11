unit uLog;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, System.Generics.Collections,
  JvComponentBase, JvThread,Vcl.Grids;

type
  TfmLog = class(TForm)
    Panel1: TPanel;
    JvThLog: TJvThread;
    Button2: TButton;
    Memo1: TMemo;
    sgLogInfo: TStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure JvThLogExecute(Sender: TObject; Params: Pointer);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure sgLogInfoDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure GridTitleUpDate;
    function DeviceLogOut(Level1, Level2, InspTime: Integer; FuncType, OkNg, Data: string): Boolean;
  end;

var
  fmLog: TfmLog;
  LogQ : TQueue<String>;


implementation
uses
  uCommonFunction;


{$R *.dfm}

procedure TfmLog.Button2Click(Sender: TObject);
var
  i  :Integer;
begin
  GridTitleUpDate;

  sgLogInfo.RowCount := 1;
  sgLogInfo.Row := sgLogInfo.RowCount-1;
  //for i := 0 to 6 do
  //  sgLogInfo.Cells[i, 1] := '';


end;

function TfmLog.DeviceLogOut(Level1, Level2, InspTime: Integer; FuncType, OkNg,
  Data: string): Boolean;
var
  sTime, wData : string;
begin
  //Level1=0, Level2=0 : Main 검사 시작 / 진행
  //Level1=0, Level2=1 : Main 검사 종료
  //Level1=1, Level2=0 : Sub 검사 시작 / 진행
  //Level1=1, Level2=1 : Sub 검사 종료
  //Level1=2, Level2=0 : FaNet

  if InspTime = -1 then
    sTime := '--'
  else
    sTime := IntToStr(InspTime);

  wData := format('%d%s%d%s%s%s%s%s%s%s%s%s%s',
                  [Level1, #9, Level2, #9, FormatDateTime('hh:nn:ss.zzz', Now), #9, sTime, #9,
                  FuncType, #9, OkNg, #9, Data]);

  LogQ.Enqueue(wData);
end;

procedure TfmLog.FormCreate(Sender: TObject);
begin
  LogQ := TQueue<String>.Create;
  JvThLog.Execute(Self);
end;

procedure TfmLog.FormDestroy(Sender: TObject);
var
  i : Integer;
begin
  if not JvThLog.Terminated then
    JvThLog.Terminate;

  LogQ.Free;
  //JvThLog.Free;
end;

procedure TfmLog.FormResize(Sender: TObject);
begin
  sgLogInfo.ColWidths[6] := fmLog.Width - 30 - 30 - 80 - 50 - 150 - 50 - 50;
end;

procedure TfmLog.FormShow(Sender: TObject);
begin
  if fmLog.Showing then
  begin
    fmLog.Caption := 'Set %d Log History';
    GridTitleUpDate;
  end;
end;

procedure TfmLog.GridTitleUpDate;
var
  i : Integer;
begin
  sgLogInfo.Cells[0, 0] := 'Lv1';
  sgLogInfo.Cells[1, 0] := 'Lv2';
  sgLogInfo.Cells[2, 0] := 'Time';
  sgLogInfo.Cells[3, 0] := 'T/T';
  sgLogInfo.Cells[4, 0] := 'FuncName';
  sgLogInfo.Cells[5, 0] := 'Result';
  sgLogInfo.Cells[6, 0] := 'Data';
  sgLogInfo.ColWidths[0] := 30;
  sgLogInfo.ColWidths[1] := 30;
  sgLogInfo.ColWidths[2] := 80;
  sgLogInfo.ColWidths[3] := 50;
  sgLogInfo.ColWidths[4] := 150;
  sgLogInfo.ColWidths[5] := 50;
  sgLogInfo.ColWidths[6] := fmLog.Width - 30 - 30 - 80 - 50 - 150 - 50 - 50;
end;

procedure TfmLog.JvThLogExecute(Sender: TObject; Params: Pointer);
var
  i, x : Integer;
  F : TextFile;
  tExceptionDir, LogData, sFileName, tFileName, tDir, OnDate, strInspTime : string;
  rData : string;
  tList : TStringList;
  sStr : PChar;
  vLParam : DWORD;
begin
  tList := TStringList.Create;

  try
    while JvThLog.Terminated = false do
    begin
      try
        if JvThLog.Terminated then
          break;

        if LogQ.Count = 0 then
        begin
          Sleep(10);
          Application.ProcessMessages;
          Continue;
        end;

        for i := 0 to LogQ.Count - 1 do
        begin
          rData := LogQ.Dequeue;

          StringParserType(rData, #9, tList);
          if tList.Count = 1 then
          begin
            if rData = 'CLEAR' then
            begin
              GridTitleUpDate;

              sgLogInfo.RowCount := 2;
              for x := 0 to 6 do
                sgLogInfo.Cells[x, 1] := '';

              sgLogInfo.Row := sgLogInfo.RowCount-1;

              continue;
            end;
          end
          else if tList.Count < 7 then
          begin
            continue;
          end;

          //화면 UpDate
          sgLogInfo.RowCount := sgLogInfo.RowCount + 1;
          sgLogInfo.Row := sgLogInfo.RowCount-1;//Set1Thread1Pos;

          for x := 0 to 6 do
            sgLogInfo.Cells[x, sgLogInfo.RowCount-1] := tList[x];

          LogData := format('[ %s ] [%s, %s] %s', [tList[5], tList[0], tList[1], tList[6]]);
          if LogData.Length >= MaxPostMsgLength then
            LogData := Copy(LogData, 1, MaxPostMsgLength-1);

          //sStr := Pchar(LogData);
          //vLParam := GlobalAddAtom(sStr);
          //PostMessage(mInspHandle, WM_MSG_INSP_STATUS, 0, vLParam);
          ////////////////////////////////

          tDir := cDir + '\NewInspLog\TotalLogInfo\';
          if not DirectoryExists(tDir) then
            ForceDirectories(tDir);

          //OnDate := Copy(tList.Strings[0], 1, 8);
          OnDate := FormatDateTime('yyyymmdd', Now);
          sFileName := tDir + OnDate + '.LOG';

          AssignFile(F, sFileName);
          try
            if FileExists(sFileName) then
              Append(F)
            else
              Rewrite(F);

              Writeln(F, rData);
          except
            LogData := FormatDateTime('hh:nn:ss.zzz', Now) + ', File Write to \NewInspLog\TotalLogInfo';
            Memo1.Lines.Add(LogData);

            CloseFile(F);
            //exit;
          end;

          CloseFile(F);

          Application.ProcessMessages;
        end;
      except
        on e: Exception do
        begin
          //if _ProgramTerminate_ = false then
          //begin
            LogData := format('%s, ExceptionType=%s', [FormatDateTime('hh:nn:ss.zzz', Now), e.Message]);
            Memo1.Lines.Add(LogData);
            //exit;
          //end;
        end;
      end;
    end;
  finally
    tList.free;
  end;
end;

procedure TfmLog.sgLogInfoDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  with (Sender as TStringGrid) do
  begin
    if (ARow = 0) then
      Canvas.Brush.Color := clBlue//clBtnFace
    else
    begin
      if ARow mod 2 = 0 then
        Canvas.Brush.Color := clInfoBk;//$00E1FFF9
      //else
      //  Canvas.Brush.Color := $00FFEBDF;

      if (Cells[0, aRow] = '0') and (Cells[1, aRow] = '0') then
      begin
        //InspStart
        Canvas.Font.Color := clGray;
      end
      else if (Cells[0, aRow] = '0') and (Cells[1, aRow] = '1') then
      begin
        //InspEnd
        if Cells[5, aRow] = 'OK' then
          Canvas.Font.Color := clBlue
        else
          Canvas.Font.Color := clRed;
      end else
      begin
        //그외
        if Cells[5, aRow] = 'OK' then
        begin
          case ACol of
            5: Canvas.Font.Color := clBlue;
          end;
        end
        else if Cells[5, aRow] = 'NG' then
        begin
          case ACol of
            5: Canvas.Font.Color := clRed;
            6: Canvas.Font.Color := clRed;
          end;
        end
        else if Cells[5, aRow] = 'Exception' then
        begin
          case ACol of
            5: begin
                  Canvas.Brush.Color := clRed;
                  Canvas.Font.Color := clWhite;
            end;
            6: begin
                  Canvas.Brush.Color := clRed;
                  Canvas.Font.Color := clWhite;
            end;
          end;
          //Canvas.Brush.Color := clRed;
        end;
      end;

      {case ACol of
        1: Canvas.Font.Color := clBlack;
        2: Canvas.Font.Color := clBlue;
        3: Canvas.Font.Color := clred;
      end;}
      Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, cells[acol, arow]);
      Canvas.FrameRect(Rect);
    end;
  end;
end;

end.
