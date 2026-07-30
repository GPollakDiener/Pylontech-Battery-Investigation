unit Batteriestatus_To_Excel_Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, ComCtrls, JclLogic, RS232_DLL, JclStrings,
  SyncObjs, DateUtils, ComObj;

const
  AnzahlSensoren=6;
  Samplingintervall=30;
  Trennzeichen=';';

  Baudrate=CBR_115200;
  // SheetType
  xlChart = -4109;
  xlWorksheet = -4167;
  // WBATemplate
  xlWBATWorksheet = -4167;
  xlWBATChart = -4109;
  // Page Setup
  xlPortrait = 1;
  xlLandscape = 2;
  xlPaperA4 = 9;
  // Format Cells
  xlBottom = -4107;
  xlLeft = -4131;
  xlRight = -4152;
  xlTop = -4160;
  // Text Alignment
  xlHAlignCenter = -4108;
  xlVAlignCenter = -4108;
  // Cell Borders
  xlThick = 4;
  xlThin = 2;
  // Berechnung
  xlManual = -4135;
  xlAutomatic = -4105;
  // Diagramm
  xlSurface = 83;
  xlColumns = 2;
  xlValue = 2;
  xlCenter = -4108;
  xlUpward = -4171;
  //Dateien
  xlNormal = -4143;

type
  TForm1 = class(TForm)
    Start_Button: TButton;
    Ende: TButton;
    Stop_Button1: TButton;
    Timer1: TTimer;
    Datei_Waehlen_Button: TButton;
    Label1: TLabel;
    Speichern_Button: TButton;
    SaveDialog1: TSaveDialog;
    Label2: TLabel;
    IntervallEdit: TEdit;
    Anzahl_Bat_Edit: TEdit;
    Label3: TLabel;
    COM_Port_Edit: TEdit;
    Label4: TLabel;
    function OpenCOM (COM:byte):integer;
    procedure FormShow(Sender: TObject);
    procedure EndeClick(Sender: TObject);
    procedure Start_ButtonClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Stop_Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Datei_Waehlen_ButtonClick(Sender: TObject);
    procedure Speichern_ButtonClick(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  Form1: TForm1;
  Excel:Variant;
  f:TextFile;
  ready,gestartet,abort,Excel_was_running:boolean;
  COM_ID,Scan,Excel_Zeile,Anzahl_Bat:integer;
  WriteBuffer,ReadBuffer:Array [0..255] of Char;
  Read_Buffer:String;
  Filename,Answer:string;
  TimeoutTime,txt,COM_Port:integer;
  Timeout:boolean;

implementation

{$R *.dfm}

function TForm1.OpenCOM (COM:byte):integer;
var
  Id:integer;
  Name:String;
  DCB:TDCB;
  CommTimeouts:TCommTimeouts;
begin
  Name:=Format ('\\.\COM%d'#0,[COM]);
  Id:=CreateFile(@Name [1],GENERIC_READ or GENERIC_WRITE,
                 0,nil,OPEN_EXISTING,FILE_ATTRIBUTE_NORMAL,0);
  if Id >= 0 then
  begin
    GetCommState(Id,DCB);
    DCB.BaudRate:=Baudrate;
    DCB.Parity:=NOPARITY;       {Parity nicht überwachen}
    DCB.ByteSize:=8;            {8 Bit übertragen}
    DCB.StopBits:=ONESTOPBIT;   {1 Stop-Bit}
    DCB.Flags:=$00001011;       {RTS on, DTR on, binary mode}
    DCB.EofChar:=#0;
    DCB.EvtChar:=#0;
    SetCommState(Id,DCB);
    GetCommTimeouts (Id,CommTimeouts);
    CommTimeouts.ReadIntervalTimeout:=MAXDWORD;  {kein Timeout}
    CommTimeouts.ReadTotalTimeoutMultiplier:=0;
    CommTimeouts.ReadTotalTimeoutConstant:=0;
    CommTimeouts.WriteTotalTimeoutMultiplier:=0; {kein Timeout}
    CommTimeouts.WriteTotalTimeoutConstant:=0;
    SetCommTimeouts (Id,CommTimeouts);
    SetupComm (Id,1000,1000);
    PurgeComm (Id,PURGE_TXABORT or PURGE_RXABORT or
               PURGE_TXCLEAR or PURGE_RXCLEAR);
  end;
  OpenCOM:=Id;
end;

{***************************************************************}

procedure Read_double (var Str1:String; var Wert:double);
var
  code,Count:integer;
  Ch:char;
  WertStr:String;
begin
  WertStr:='';
  if Length (Str1) > 0 then
  begin
    if Str1 [1] = '+' then delete (Str1,1,1);
    repeat
      ch:=Str1[1];
      Str1:=copy (Str1,2,Length (Str1)-1);
      if (ch <> ' ') and (ord(ch) > 32) then
        WertStr:=WertStr + ch;
    until (STR1 = '') or (ch = ' ');
    if Length (WertStr) > 0 then
    begin
      val (WertStr,Wert,code);
      Count:=0;
      while (Count+1 <= Length(Str1)) and (Str1[Count+1] = ' ') do inc(Count);
      Delete(Str1,1,Count);
    end
    else
      Wert:=0.0;
  end
  else
    Wert:=0.0;
end;

{***************************************************************}

procedure Read_String (var Str1:String; var WertStr:String);
var
  Ch:char;
  Count:integer;
begin
  WertStr:='';
  if Length (Str1) > 0 then
  begin
    repeat
      ch:=Str1[1];
      Str1:=copy (Str1,2,Length (Str1)-1);
      if (ch <> ' ') and (ord(ch) > 32) then
        WertStr:=WertStr + ch;
    until (STR1 = '') or (ch = ' ');
    if Length(Str1) > 0 then
    begin
      Count:=0;
      while (Count+1 <= Length(Str1)) and (Str1[Count+1] = ' ') do inc(Count);
      Delete(Str1,1,Count);
    end;
  end;
end;

{***************************************************************}

procedure SendCommand (Command:String; TTimeout:integer);
var
  Len,ok,i:integer;
  CR_LF_erkannt:boolean;
  Buffer:Array [0..200] of char;
begin
  CR_LF_erkannt:=false;
  Len:=Length (Command);
  StrPCopy (Buffer,Command);
  FileWrite (COM_ID,Buffer,Len);
  Sleep(350);
  ok:=FileRead (COM_ID, ReadBuffer, 2000);
  if ok > 0 then
  begin
    for i:=0 to ok-1 do
    begin
      if (ReadBuffer [i] = #13) or (ReadBuffer [i] = #10) then
      begin
        if not CR_LF_erkannt then
        begin
          writeln(f);
          CR_LF_erkannt:=true;
        end;
      end
      else
      begin
        write (f, ReadBuffer [i]);
        CR_LF_erkannt:=false;
      end;
    end;
  end;
end;

{***************************************************************}

procedure TForm1.FormShow(Sender: TObject);
var
  Cell,Bat:integer;
begin
  Timer1.enabled:=false;
  Excel_was_running:=false;
end;

procedure TForm1.EndeClick(Sender: TObject);
begin
  Close;
end;

procedure TForm1.Start_ButtonClick(Sender: TObject);
var
Intervall:real;
Bat,Cell,Code:integer;
begin
  Val (COM_Port_Edit.Text,COM_Port,Code);
  if Code > 0 then
  begin
    ShowMessage('COM_Port nicht Integer');
    Exit;
  end;
  COM_ID:=OpenCom (COM_Port);
  if COM_ID < 0 then
  begin
    MessageDLG ('COM-Port nicht verfügbar!',mtError,[mbOK],0);
    Exit;
  end;
  Speichern_Button.Enabled:=false;
  Val (Anzahl_Bat_Edit.Text,Anzahl_Bat,Code);
  if Code > 0 then Anzahl_Bat:=1;
  if Anzahl_Bat > 16 then
  begin
    Anzahl_Bat:=16;
    Anzahl_Bat_Edit.Text:=Format ('%d',[Anzahl_Bat]);
  end;
  Val (IntervallEdit.Text,Intervall,Code);
  if Code > 0 then Intervall:=Anzahl_Bat;
  if Intervall < Anzahl_Bat then Intervall:=Anzahl_Bat;
  IntervallEdit.Text:=Format ('%3.1f',[Intervall]);
  Application.ProcessMessages;
  if not Excel_was_running then
  begin
    {Excel öffnen}
    try
      Excel:=CreateOleObject ('Excel.Application');
    except
      ShowMessage('Kann Excel nicht öffnen!');
      exit;
    end;
    Excel_was_running:=true;
    Excel.Visible:=true;
    Excel.Workbooks.Add;
    Excel.Application.displayalerts:=false;
    Excel.ActiveSheet.delete;
    Excel.ActiveSheet.delete;
    Excel.Application.displayalerts:=true;
    Excel.ActiveSheet.Name:='Balance';
    Excel.Cells[2,1].Value:='Zeit';
    Excel.ActiveSheet.Range['A:A'].NumberFormat:='TT.MM.JJ hh:mm:ss';
    Excel.ActiveSheet.Range['A:A'].ColumnWidth:=16;
    Excel.Cells[3,2].Select;
    Excel.ActiveWindow.FreezePanes:=true;
    for Bat:=0 to Anzahl_Bat-1 do
    begin
      Excel.Cells[1,Bat*15+2].Value:=Format ('Batterie %d',[Bat+1]);
      for Cell:=0 to 14 do
        Excel.Cells[2,Bat*15+2 + Cell].Value:=Format ('Zelle %d',[Cell]);
    end;

    Excel.Worksheets.Add;
    Excel.ActiveSheet.Name:='Ladung';
    Excel.Cells[2,1].Value:='Zeit';
    Excel.ActiveSheet.Range['A:A'].NumberFormat:='TT.MM.JJ hh:mm:ss';
    Excel.ActiveSheet.Range['A:A'].ColumnWidth:=16;
    Excel.ActiveSheet.Range['B:IV'].NumberFormat:='0.000';
    Excel.Cells[3,2].Select;
    Excel.ActiveWindow.FreezePanes:=true;
    for Bat:=0 to Anzahl_Bat-1 do
    begin
      Excel.Cells[1,Bat*15+2].Value:=Format ('Batterie %d',[Bat+1]);
      for Cell:=0 to 14 do
        Excel.Cells[2,Bat*15+2 + Cell].Value:=Format ('Zelle %d',[Cell]);
    end;

    Excel.Worksheets.Add;
    Excel.ActiveSheet.Name:='SOC';
    Excel.Cells[2,1].Value:='Zeit';
    Excel.ActiveSheet.Range['A:A'].NumberFormat:='TT.MM.JJ hh:mm:ss';
    Excel.ActiveSheet.Range['A:A'].ColumnWidth:=16;
    Excel.ActiveSheet.Range['B:IV'].NumberFormat:='0%';
    Excel.Cells[3,2].Select;
    Excel.ActiveWindow.FreezePanes:=true;
    for Bat:=0 to Anzahl_Bat-1 do
    begin
      Excel.Cells[1,Bat*15+2].Value:=Format ('Batterie %d',[Bat+1]);
      for Cell:=0 to 14 do
        Excel.Cells[2,Bat*15+2 + Cell].Value:=Format ('Zelle %d',[Cell]);
    end;

    Excel.Worksheets.Add;
    Excel.ActiveSheet.Name:='Temperatur';
    Excel.Cells[2,1].Value:='Zeit';
    Excel.ActiveSheet.Range['A:A'].NumberFormat:='TT.MM.JJ hh:mm:ss';
    Excel.ActiveSheet.Range['A:A'].ColumnWidth:=16;
    Excel.ActiveSheet.Range['B:IV'].NumberFormat:='0.0';
    Excel.Cells[3,2].Select;
    Excel.ActiveWindow.FreezePanes:=true;
    for Bat:=0 to Anzahl_Bat-1 do
    begin
      Excel.Cells[1,Bat*15+2].Value:=Format ('Batterie %d',[Bat+1]);
      for Cell:=0 to 14 do
        Excel.Cells[2,Bat*15+2 + Cell].Value:=Format ('Zelle %d',[Cell]);
    end;

    Excel.Worksheets.Add;
    Excel.ActiveSheet.Name:='Strom';
    Excel.Cells[2,1].Value:='Zeit';
    Excel.ActiveSheet.Range['A:A'].NumberFormat:='TT.MM.JJ hh:mm:ss';
    Excel.ActiveSheet.Range['A:A'].ColumnWidth:=16;
    Excel.ActiveSheet.Range['B:IV'].NumberFormat:='0.000';
    Excel.Cells[3,2].Select;
    Excel.ActiveWindow.FreezePanes:=true;
    for Bat:=0 to Anzahl_Bat-1 do
    begin
      Excel.Cells[2,2 + Bat].Value:=Format ('Batterie %d',[Bat+1]);
    end;

    Excel.Worksheets.Add;
    Excel.ActiveSheet.Name:='Spannung';
    Excel.Cells[2,1].Value:='Zeit';
    Excel.ActiveSheet.Range['A:A'].NumberFormat:='TT.MM.JJ hh:mm:ss';
    Excel.ActiveSheet.Range['A:A'].ColumnWidth:=16;
    Excel.ActiveSheet.Range['B:IV'].NumberFormat:='0.000';
    Excel.Cells[3,2].Select;
    Excel.ActiveWindow.FreezePanes:=true;
    for Bat:=0 to Anzahl_Bat-1 do
    begin
      Excel.Cells[1,Bat*15+2].Value:=Format ('Batterie %d',[Bat+1]);
      for Cell:=0 to 14 do
        Excel.Cells[2,Bat*15+2 + Cell].Value:=Format ('Zelle %d',[Cell]);
    end;
    Excel_Zeile:=3;
  end;

  AssignFile(f,'C:\temp.txt');
  Speichern_Button.Enabled:=false;
  Timer1.Interval:=round(Intervall*1000);
  Timer1.Enabled:=true;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Timer1.Enabled:=false;
  if COM_ID >= 0 then FileClose (COM_ID);
end;

procedure TForm1.Stop_Button1Click(Sender: TObject);
begin
  Timer1.Enabled:=false;
  Speichern_Button.Enabled:=true;
  if COM_ID >= 0 then FileClose (COM_ID);
  Speichern_Button.Enabled:=true;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
  jetzt:TDateTime;
  Bat,Cell,n,code:integer;
  Line,search_Str,CellStr,VoltStr,CurrStr:String;
  TempStr,DummyStr,socStr,QStr,BalStr:String;
  Voltage,Current,Temperature,soc,Ladung:double;
begin
  Cursor:=crHourglass;
  Application.ProcessMessages;
  Rewrite(f);
  if Anzahl_Bat = 1 then
    SendCommand ('bat' + chr(13),1)
  else
    for Bat:=1 to Anzahl_Bat do
    begin
      SendCommand (Format('bat %d',[Bat]) + chr(13),1);
    end;
  CloseFile(f);

  jetzt:=now;
  Reset(f);
  Excel.Workbooks[1].Sheets['Spannung'].Cells[Excel_Zeile,1].Value:=jetzt;
  Excel.Workbooks[1].Sheets['Strom'].Cells[Excel_Zeile,1].Value:=jetzt;
  Excel.Workbooks[1].Sheets['Temperatur'].Cells[Excel_Zeile,1].Value:=jetzt;
  Excel.Workbooks[1].Sheets['SOC'].Cells[Excel_Zeile,1].Value:=jetzt;
  Excel.Workbooks[1].Sheets['Ladung'].Cells[Excel_Zeile,1].Value:=jetzt;
  Excel.Workbooks[1].Sheets['Balance'].Cells[Excel_Zeile,1].Value:=jetzt;
  for Bat:=0 to Anzahl_Bat-1 do
  begin
    if Anzahl_Bat = 1 then
      search_Str:='Battery'
    else
      search_Str:=format('bat %d',[Bat+1]);
    repeat
      Readln(f,Line);
    until (Pos(search_Str,Line) <> 0) or eof(f);
    if eof(f) then
    begin
      CloseFile (f);
      Cursor:=crDefault;
      Application.ProcessMessages;
      exit;
    end
    else
    begin
      if Anzahl_Bat > 1 then
      begin
        Readln(f);
        Readln(f);
      end;
      for Cell:=0 to 14 do
      begin
        repeat Readln(f,Line) until Line <> '';
        Read_String(Line,CellStr);
        Val(CellStr,n,code);
        if (code = 0) and (n = Cell) then
        begin
          Read_double(Line,Voltage);
          Excel.Workbooks[1].Sheets['Spannung'].Cells[Excel_Zeile,Bat*15+2+Cell].Value:=
            Voltage/1000.0;

          Read_String(Line,CurrStr);
          if Cell = 0 then
          begin
            Val(CurrStr,Current,code);
            Excel.Workbooks[1].Sheets['Strom'].Cells[Excel_Zeile,2+Bat].Value:=
              Current/1000.0;
          end;

          Read_double(Line,Temperature);
          Excel.Workbooks[1].Sheets['Temperatur'].Cells[Excel_Zeile,Bat*15+2+Cell].Value:=
            Temperature/1000.0;

          Read_String(Line,DummyStr);
          Read_String(Line,DummyStr);
          Read_String(Line,DummyStr);
          Read_String(Line,DummyStr);

          Read_double(Line,soc);
          Excel.Workbooks[1].Sheets['SOC'].Cells[Excel_Zeile,Bat*15+2+Cell].Value:=
            soc/100.0;

          Read_double(Line,Ladung);
          Excel.Workbooks[1].Sheets['Ladung'].Cells[Excel_Zeile,Bat*15+2+Cell].Value:=
            Ladung/1000.0;

          Read_String(Line,DummyStr);

          Read_String(Line,BalStr);
          if BalStr = 'Y' then
            Excel.Workbooks[1].Sheets['Balance'].Cells[Excel_Zeile,Bat*15+2+Cell].Value:=1
          else
            Excel.Workbooks[1].Sheets['Balance'].Cells[Excel_Zeile,Bat*15+2+Cell].Value:=0;
        end
        else
        begin
          Timer1.Enabled:=false;
          ShowMessage('Zellennummer nicht erkannt!');
        end;
      end;
    end;
  end;
  CloseFile(f);
  inc(Excel_Zeile);
  Cursor:=crDefault;
  Application.ProcessMessages;
end;

procedure TForm1.Datei_Waehlen_ButtonClick(Sender: TObject);
begin
  Form1.SaveDialog1.InitialDir:='C:\';
  Form1.SaveDialog1.Filter := 'txt-Dateien (*.txt)';
  FileName:='';
  if Form1.SaveDialog1.Execute then
  begin
    FileName:=Form1.SaveDialog1.FileName;
    if pos ('.',FileName) = 0 then
    FileName:=FileName + '.txt';
    Form1.Label1.Caption:=FileName;
    txt:=pos ('.TXT',AnsiUpperCase (FileName));
  end;
end;

procedure TForm1.Speichern_ButtonClick(Sender: TObject);
var
  FilenameDated:string;
  jetzt:TDateTime;
  Year,Month,Day,Hour,Minute:integer;
  Bat,Code:integer;
begin
  Val (COM_Port_Edit.Text,COM_Port,Code);
  if Code > 0 then
  begin
    ShowMessage('COM_Port nicht Integer');
    Exit;
  end;
  COM_ID:=OpenCom (COM_Port);
  if COM_ID < 0 then
  begin
    MessageDLG ('COM-Port nicht verfügbar!',mtError,[mbOK],0);
    Exit;
  end;
  if Label1.Caption ='No Name' then
  begin
    Datei_Waehlen_ButtonClick(Sender);
  end;
  Val (Anzahl_Bat_Edit.Text,Anzahl_Bat,Code);
  if Code > 0 then Anzahl_Bat:=1;
  if Anzahl_Bat > 16 then Anzahl_Bat:=16;
  Cursor:=crHourglass;
  Start_Button.Enabled:=false;
  Application.ProcessMessages;
  FilenameDated:=copy(Filename,1,txt-1);
  jetzt:=Now;
  Year:=YearOf(jetzt);
  Month:=MonthOf(jetzt);
  Day:=DayOf(jetzt);
  Hour:=HourOf(jetzt);
  Minute:=MinuteOf(jetzt);
  FilenameDated:=Format('%s_%d%.2d%.2d_%.2d%.2d.txt',
                        [FilenameDated,Year,Month,Day,Hour,Minute]);
  AssignFile(f,FilenameDated);
  Rewrite(f);
  Writeln(f,DateTimeToStr(jetzt));
  if Anzahl_Bat = 1 then
    SendCommand ('bat' + chr(13),1)
  else
    for Bat:=1 to Anzahl_Bat do
    begin
      SendCommand (Format('bat %d',[Bat]) + chr(13),1);
    end;
  CloseFile(f);
  Cursor:=crDefault;
  Start_Button.Enabled:=true;
  Application.ProcessMessages;
  if COM_ID >= 0 then FileClose (COM_ID);
end;

end.
