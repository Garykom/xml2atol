program xml2atol;

{$mode objfpc}{$H+}

uses {$IFDEF UNIX} {$IFDEF UseCThreads}
  cthreads, {$ENDIF} {$ENDIF}
  Classes,
  SysUtils,
  CustApp,
  ComObj,
  DOM,
  XMLRead,
  xmliconv_windows { you can add units after this };

type
  TFiscalString = record
    Name: widestring;
    Quantity: double;
    Price: double;
    Amount: double;
    Department: byte;
    Tax: string;
  end;

  TCheck = record
    PaymentType: byte;
    TaxVariant: byte;
    CustomerEmail, CustomerPhone: widestring;
    Positions: array of TFiscalString;
    Cash, CashLessType1, CashLessType2, CashLessType3: double;
  end;

  TFR = class
    ECR: olevariant;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Init;
    procedure Abort;
    procedure Start;
    function SetMode(Mode: integer; Password: string): integer;
    procedure CancelCheck;
    procedure Stop;
    function OpenCheck(CheckMode: integer; CheckType: integer): integer;
    procedure Registration(Name: widestring; Price: double; Quantity: double;
      TaxTypeNumber: integer; Department: integer; DiscountType: integer;
      DiscountValue: double);
    procedure PrintString(str: widestring);
    procedure Payment(Summ: double; TypeClose: integer; var Remainder: double;
      var Change: double);
    procedure WriteAttribute(AttrNumber: integer; AttrValue: string);
    function CloseCheck: integer;
    procedure Feed(numLines: integer);
    procedure FullCut;
    procedure PartialCut;
    procedure PrintXReport;
    procedure PrintZReport;
    function ResultCode: integer;
    function ResultDescription: string;
    procedure GetLastError(var Error: integer; var ErrorDescription: integer);
    procedure Print(Lines: array of string);

    function AtolGetTaxByString(TaxString: string): integer;
    function AtolGetTaxVariantByLong(TaxVariant: integer): byte;
  end;

  { TMyApplication }
  TMyApplication = class(TCustomApplication)
  protected
    procedure DoRun; override;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
    procedure WriteHelp; virtual;
    function XMLtoCheck(Doc: TXMLDocument): TCheck;
    procedure ProcessCheck(Check: TCheck);
  end;

  { TFR }
  constructor TFR.Create;
  begin

  end;

  destructor TFR.Destroy;
  begin
    inherited Destroy;
  end;

  procedure TFR.Init;
  begin
    // создаем объект общего драйвера ККМ
    // если объект создать не удается генерируется исключение, по которому завершается работа приложения
    try
      ECR := CreateOleObject('AddIn.FprnM45');
    except
      Writeln('Не удалось создать объект общего драйвера ККМ!');
    end;
  end;

  procedure TFR.Abort;
  begin
    CancelCheck();
    Stop();
  end;

  procedure TFR.Start;
  begin
    // занимаем порт
    ECR.DeviceEnabled := True;
    if ECR.ResultCode <> 0 then
      exit;
    // получаем состояние ККМ
    if ECR.GetStatus() <> 0 then
      exit;
    CancelCheck();
  end;

  function TFR.SetMode(Mode: integer; Password: string): integer;
  begin
    ECR.Password := Password;
    ECR.Mode := Mode;
    Result := ECR.SetMode();
  end;

  procedure TFR.CancelCheck;
  begin
    // если есть открытый чек, то отменяем его
    if ECR.CheckState <> 0 then
      if ECR.CancelCheck() <> 0 then
        exit;
  end;

  procedure TFR.Stop;
  begin
    // выходим в режим выбора, чтобы кто-то под введенными паролями не сделал что нибуть нехорошее
    if ECR.ResetMode() <> 0 then
      exit;

    // освобождаем порт
    ECR.DeviceEnabled := False;
    if ECR.ResultCode <> 0 then
      exit;
  end;

  function TFR.OpenCheck(CheckMode: integer; CheckType: integer): integer;
  begin
    ECR.CheckMode := CheckMode;
    ECR.CheckType := CheckType;

    Result := ECR.OpenCheck();
  end;

  procedure TFR.Registration(Name: widestring; Price: double; Quantity: double;
    TaxTypeNumber: integer; Department: integer; DiscountType: integer;
    DiscountValue: double);
  begin
    ECR.Name := Name;
    ECR.Price := Price;
    ECR.Quantity := Quantity;

    ECR.Destination := 1;
    ECR.TaxTypeNumber := TaxTypeNumber;

    ECR.Department := Department;
    ECR.DiscountType := DiscountType;
    ECR.DiscountValue := DiscountValue;

    if ECR.Registration() <> 0 then
      exit;
  end;

  procedure TFR.PrintString(str: widestring);
  begin
    ECR.Caption := str;
    ECR.TextWrap := 1;
    ECR.PrintString();
  end;

  procedure TFR.Payment(Summ: double; TypeClose: integer; var Remainder: double;
  var Change: double);
  begin
    ECR.Summ := Summ;
    ECR.TypeClose := TypeClose;
    ECR.Payment();

    Remainder := ECR.Remainder;
    Change := ECR.Change;
  end;

  procedure TFR.WriteAttribute(AttrNumber: integer; AttrValue: string);
  begin
    ECR.AttrNumber := AttrNumber;
    ECR.AttrValue := AttrValue;
    ECR.WriteAttribute();
  end;

  function TFR.CloseCheck: integer;
  begin
    Result := ECR.CloseCheck();
  end;

  procedure TFR.Feed(numLines: integer);
  var
    i: integer;
  begin
    for i := 0 to (numLines - 1) do
    begin
      ECR.Caption := '  ';
      ECR.PrintString();
    end;
  end;

  procedure TFR.FullCut;
  begin
    ECR.FullCut();
  end;

  procedure TFR.PartialCut;
  begin
    ECR.PartialCut();
  end;

  procedure TFR.PrintXReport;
  begin
    ECR.Password := '30';
    ECR.Mode := 2;
    if ECR.SetMode() <> 0 then
      exit;

    ECR.ReportType := 2;
    ECR.Report();
  end;

  procedure TFR.PrintZReport;
  begin
    ECR.Password := '30';
    ECR.Mode := 3;
    if ECR.SetMode() <> 0 then
      exit;

    ECR.ReportType := 1;
    ECR.Report();
  end;

  function TFR.ResultCode: integer;
  begin
    Result := ECR.ResultCode;
  end;

  function TFR.ResultDescription: string;
  begin
    Result := ECR.ResultDescription;
  end;

  procedure TFR.GetLastError(var Error: integer; var ErrorDescription: integer);
  begin
    ECR.GetLastError();
    Error := ECR.ECRError;
    ErrorDescription := ECR.ECRErrorDescription;
  end;

  procedure TFR.Print(Lines: array of string);
  var
    i: integer;
  begin
    ECR.Caption := 'Начало проверки';
    ECR.TextWrap := 1;
    ECR.PrintString();

    for i := 0 to (Length(Lines) - 1) do
    begin
      //ECR.Caption := Lines[i];
      //ECR.TextWrap := 1;
      //ECR.PrintString();
    end;
  end;

  function TFR.AtolGetTaxByString(TaxString: string): integer;
  begin
    Result := -1;
    if TaxString = '18' then
      Result := 1;
    if TaxString = '10' then
      Result := 2;
    if TaxString = '0' then
      Result := 5;
    if TaxString = 'none' then
      Result := 6;
  end;

  function TFR.AtolGetTaxVariantByLong(TaxVariant: integer): byte;
  begin
    // Применяемая система налогооблажения в чеке:
    // ОСН - 1
    // УСН доход - 2
    // УСН доход-расход - 4
    // ЕНВД - 8
    // ЕСН - 16
    // ПСН - 32
    Result := 0;
    if TaxVariant = 0 then
      Result := 1;
    if TaxVariant = 1 then
      Result := 2;
    if TaxVariant = 2 then
      Result := 4;
    if TaxVariant = 3 then
      Result := 8;
    if TaxVariant = 4 then
      Result := 16;
    if TaxVariant = 5 then
      Result := 32;
  end;


  { TMyApplication }

  procedure TMyApplication.DoRun;
  var
    ErrorMsg: string;
    Doc: TXMLDocument;

    XMLPath: string;
    Check: TCheck;
    FR: TFR;
    F: Text;
    Line: string;

  begin
    // quick check parameters
    ErrorMsg := CheckOptions('h', 'help');
    if ErrorMsg <> '' then
    begin
      ShowException(Exception.Create(ErrorMsg));
      Terminate;
      Exit;
    end;

    // parse parameters
    if HasOption('h', 'help') then
    begin
      WriteHelp;
      Terminate;
      Exit;
    end;

    { add your program here }

    if GetParamCount >= 1 then
    begin
      XMLPath := GetParams(1);
      if ExtractFilePath(XMLPath) = '' then
        XMLPath := ExtractFilePath(GetParams(0)) + XMLPath;
      Writeln('' + XMLPath);
      try
        ReadXMLFile(Doc, XMLPath);
        Check := XMLtoCheck(Doc);
        ProcessCheck(Check);
      except
        AssignFile(F, XMLPath);
        Reset(F);
        Readln(F, Line);

        if line = '#PrintXReport' then
        begin
          FR := TFR.Create;
          FR.Init;
          FR.Start;
          FR.PrintXReport();
          FR.Stop();
        end;
        if line = '#PrintZReport' then
        begin
          FR := TFR.Create;
          FR.Init;
          FR.Start;
          FR.PrintZReport();
          FR.Stop();
        end;
      end;
    end;

    // stop program loop
    Terminate;
  end;

  constructor TMyApplication.Create(TheOwner: TComponent);
  begin
    inherited Create(TheOwner);
    StopOnException := True;
  end;

  destructor TMyApplication.Destroy;
  begin
    inherited Destroy;
  end;

  procedure TMyApplication.WriteHelp;
  begin
    { add your help code here }
    writeln('Usage: ', ExeName, ' -h');
  end;

  function TMyApplication.XMLtoCheck(Doc: TXMLDocument): TCheck;
  var
    Check: TCheck;
    Child: TDOMNode;
    i: integer;
    fs: TFormatSettings;
  begin
    //GetLocaleFormatSettings(1, fs);
    fs := FormatSettings;
    fs.DecimalSeparator := '.';

    Child := Doc.DocumentElement.FirstChild;
    while Assigned(Child) do
    begin
      Writeln(Child.NodeName);
      if Child.NodeName = 'Parameters' then
      begin
        Check.PaymentType :=
          StrToInt(Child.Attributes.GetNamedItem('PaymentType').NodeValue);
        Check.TaxVariant := StrToInt(Child.Attributes.GetNamedItem(
          'TaxVariant').NodeValue);
        Check.CustomerEmail :=
          Child.Attributes.GetNamedItem('CustomerEmail').NodeValue;
        Check.CustomerPhone :=
          Child.Attributes.GetNamedItem('CustomerPhone').NodeValue;

        Writeln('PaymentType=' + IntToStr(Check.PaymentType));
        Writeln('TaxVariant=' + IntToStr(Check.TaxVariant));
      end;
      if Child.NodeName = 'Positions' then
      begin
        SetLength(Check.Positions, Child.ChildNodes.Count);
        for i := 0 to (Child.ChildNodes.Count - 1) do
        begin
          Check.Positions[i].Name :=
            Child.ChildNodes.Item[i].Attributes.GetNamedItem('Name').NodeValue;
          Check.Positions[i].Quantity :=
            StrToFloat(Child.ChildNodes.Item[i].Attributes.GetNamedItem(
            'Quantity').NodeValue, fs);
          Check.Positions[i].Price :=
            StrToFloat(Child.ChildNodes.Item[i].Attributes.GetNamedItem(
            'Price').NodeValue, fs);
          Check.Positions[i].Amount :=
            StrToFloat(Child.ChildNodes.Item[i].Attributes.GetNamedItem(
            'Amount').NodeValue, fs);
          Check.Positions[i].Department :=
            StrToInt(Child.ChildNodes.Item[i].Attributes.GetNamedItem(
            'Department').NodeValue);
          Check.Positions[i].Tax :=
            Child.ChildNodes.Item[i].Attributes.GetNamedItem('Tax').NodeValue;

          Writeln('Name=' + Check.Positions[i].Name);
          Writeln('Quantity=' + FloatToStr(Check.Positions[i].Quantity));
          Writeln('Price=' + FloatToStr(Check.Positions[i].Price));
          Writeln('Amount=' + FloatToStr(Check.Positions[i].Amount));
          Writeln('Department=' + IntToStr(Check.Positions[i].Department));
          Writeln('Tax=' + Check.Positions[i].Tax);
        end;
      end;
      if Child.NodeName = 'Payments' then
      begin
        Check.Cash := StrToFloat(Child.Attributes.GetNamedItem(
          'Cash').NodeValue, fs);
        Check.CashLessType1 :=
          StrToFloat(Child.Attributes.GetNamedItem('CashLessType1').NodeValue, fs);
        Check.CashLessType2 :=
          StrToFloat(Child.Attributes.GetNamedItem('CashLessType2').NodeValue, fs);
        Check.CashLessType3 :=
          StrToFloat(Child.Attributes.GetNamedItem('CashLessType3').NodeValue, fs);

        Writeln('Cash=' + FloatToStr(Check.Cash));
        Writeln('CashLessType1=' + FloatToStr(Check.CashLessType1));
        Writeln('CashLessType2=' + FloatToStr(Check.CashLessType2));
        Writeln('CashLessType3=' + FloatToStr(Check.CashLessType3));
      end;

      Child := Child.NextSibling;
    end;
    Result := Check;
  end;

  procedure TMyApplication.ProcessCheck(Check: TCheck);
  var
    FR: TFR;
    Result: integer;
    i: integer;

    Remainder, Change: double;
    ResultDescription: widestring;
  begin
    FR := TFR.Create;
    FR.Init;
    FR.Start;

    Result := FR.SetMode(1, '30');
    if Result = -3822 then
    begin
      FR.PrintString(UTF8ToAnsi('Смена превысила 24 часа!'));
      FR.PrintString(UTF8ToAnsi('Для продолжения снимите Z-Отчет.'));
      FR.Feed(6);
      FR.PartialCut();
      FR.Stop();
      exit;
    end;

    Result := FR.OpenCheck(1, Check.PaymentType);
    // Применяемая система налогооблажения в чеке:
    FR.WriteAttribute(1055, IntToStr(FR.AtolGetTaxVariantByLong(check.TaxVariant)));

    // Позиции чека
    for i := 0 to (Length(Check.Positions) - 1) do
    begin
      FR.Registration(Check.Positions[i].Name, Check.Positions[i].Price,
        Check.Positions[i].Quantity, FR.AtolGetTaxByString(Check.Positions[i].Tax),
        Check.Positions[i].Department,
        0, 0);
    end;

    // Оплата
    Remainder := 0;
    Change := 0;
    FR.Payment(Check.Cash, 0, Remainder, Change);
    FR.Payment(Check.CashLessType1, 1, Remainder, Change);
    FR.Payment(Check.CashLessType2, 2, Remainder, Change);
    FR.Payment(Check.CashLessType3, 3, Remainder, Change);

    Result := FR.CloseCheck();

    ResultDescription := FR.ResultDescription();
    if Result <> 0 then
    begin
      Writeln('Error: ' + IntToStr(Result) + ' - ' + ResultDescription);
      FR.PrintString('Error: ' + IntToStr(Result) + ' - ' + ResultDescription);
      FR.Abort();
    end;

    FR.Stop();
  end;

var
  Application: TMyApplication;
begin
  Application := TMyApplication.Create(nil);
  Application.Title := 'My Application';
  Application.Run;
  Application.Free;
end.
