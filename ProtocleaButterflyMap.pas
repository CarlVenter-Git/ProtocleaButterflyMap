unit ProtocleaButterflyMap;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, ComObj, FMX.Objects, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdHTTP, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, FMX.ListBox, DateUtils,
  Generics.Collections, RegularExpressions;

type
  TForm1 = class(TForm)
    btnLoad: TButton;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    lblPath: TLabel;
    btnPlotPoints: TButton;
    Image1: TImage;
    IdHTTP1: TIdHTTP;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    cmbYear: TComboBox;
    cmbMonth: TComboBox;
    cmbProvince: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    btnVerify: TButton;
    btnBest: TButton;
    btnExit: TButton;
    procedure btnLoadClick(Sender: TObject);
    procedure btnPlotPointsClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnBestClick(Sender: TObject);
    procedure btnVerifyClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

type
  sightingRecord = record
    rec_nr: string;
    genus: string;
    species: string;
    subspecies: string;
    rec_date: TDateTime;
    province: string;
    latitude: string;
    longitude: string;
end;

var
  Form1: TForm1;
  sightings: TList<sightingRecord>;

implementation
{$R *.fmx}

procedure TForm1.btnLoadClick(Sender: TObject);
var
  excel, book, sheet, range: OLEVariant;
  recordToAdd: sightingRecord;
  selectedFilePath, province: string;
  provinces: TList<string>;
  years, months: TList<integer>;
  x, numRows, year, month(*, numColumns*): integer;

begin
  sightings := TList<sightingRecord>.Create;
  provinces := TList<string>.Create;
  years := TList<integer>.Create;
  months := TList<integer>.Create;

  try
    if OpenDialog1.Execute then
      selectedFilePath := OpenDialog1.FileName;
  finally
    //OpenDialog1.Free; Causes issues when attempting to open the dialog more than once
  end;

  if selectedFilePath <> '' then

  begin

    try
      excel := CreateOleObject('Excel.Application');

      excel.Visible := False;
      excel.DisplayAlerts := False;
      book := excel.Workbooks.Open(selectedFilePath);
      sheet := book.Worksheets.Item[1];
      range := sheet.UsedRange;

      numRows := range.Rows.Count;
      //numColumns := range.Columns.Count; unused in this case

      for x := 2 to numRows do //Start at index 2 to skip the heading
      begin
        lblPath.Text := 'Loading ' + IntToStr(x - 1) + '/' + IntToStr(numRows);

        //genus,species and subspecies is rather redundant in this case but still useful to have in the record for use at a later stage if the data ever includes more
        recordToAdd.rec_nr := sheet.Cells.Item[x, 1];
        recordToAdd.genus := sheet.Cells.Item[x, 2];
        recordToAdd.species := sheet.Cells.Item[x, 3];
        recordToAdd.subspecies := sheet.Cells.Item[x, 4];
        recordToAdd.rec_date := sheet.Cells.Item[x, 5];
        recordToAdd.province := sheet.Cells.Item[x, 6];
        recordToAdd.latitude := sheet.Cells.Item[x, 7];
        recordToAdd.longitude := sheet.Cells.Item[x, 8];

        if not provinces.Contains(recordToAdd.province) then
          provinces.Add(recordToAdd.province);

        if not months.Contains(StrToInt(DateUtils.MonthOf(recordToAdd.rec_date).ToString)) then
          months.Add(StrToInt(DateUtils.MonthOf(recordToAdd.rec_date).ToString));

        if not years.Contains(StrToInt(DateUtils.YearOf(recordToAdd.rec_date).ToString)) then
          years.Add(StrToInt(DateUtils.YearOf(recordToAdd.rec_date).ToString));

        sightings.Add(recordToAdd);
      end;

      years.Sort;
      months.Sort;
      provinces.Sort;

      lblPath.Text := selectedFilePath + ' Successfully Loaded';

      for year in years do
        cmbYear.Items.Add(IntToStr(year));

      for month in months do
        cmbMonth.Items.Add(IntToStr(month));

      for province in provinces do
        cmbProvince.Items.Add(province);

    finally
      provinces.Free;
      years.Free;
      months.Free;

      excel.Quit;
      excel := Unassigned;
    end;
  end;

end;

procedure TForm1.btnPlotPointsClick(Sender: TObject);
const
  key = '';
  prefix = 'https://maps.googleapis.com/maps/api/staticmap?';
  markerColour = 'markers=color:red';
var
  ms: TMemoryStream;
  url, markerString, center: string;
  x, y: single;
  i: integer;
begin
  //Exit the method if data has not been loaded
  if sightings = nil then
    exit;

  x := Panel1.Size.Width;
  y := Panel1.Size.Height;

  markerString := '';

  for i := 0 to sightings.Count - 1 do
  begin

    if (sightings[i].province = cmbProvince.Selected.Text) and
    (DateUtils.YearOf(sightings[i].rec_date).ToString = cmbYear.Selected.Text) and
    (DateUtils.MonthOf(sightings[i].rec_date).ToString = cmbMonth.Selected.Text) then
    begin
      center := sightings[i].province;
      center := center.Replace(' ', '+');

      markerString := markerString + markerColour + '|label:' + sightings[i].rec_nr +
      '|' + sightings[i].latitude + ',' + sightings[i].longitude + '&';
    end;

  end;

  //URL cannot be more than 2048 characters, this needs to be taken into account
  url := prefix + 'center=' + center + 'South+Africa&zoom=5&size=' + //I need to specify South Africa or else the map tends to center on 'Murica
         x.ToString + 'x' + y.ToString + '&maptype=roadmap&'+
         markerString + 'key=' + key;

  ms := TMemoryStream.Create;

  try
    idHTTP1.Get(url, ms);

    ms.Position := 0;

    Image1.Bitmap.LoadFromStream(ms);
  except
    ShowMessage('Map Unavailable');
    ms.Free;
    exit;
  end;
end;

procedure TForm1.btnVerifyClick(Sender: TObject);
var
i: integer;
errorID: string;
errorFound: boolean;
begin
  if sightings = nil then
    exit;

  errorFound := false;

  for i := 0 to sightings.Count - 1 do
  begin
    if not TRegEx.IsMatch(sightings[i].rec_nr,'[0-9 ]+') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    //I am matching the following 3 expressions exactly, my assumption is this data set is meant to focus on them, anything else must be flagged
    if not TRegEx.IsMatch(sightings[i].genus,'Belenois') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;


    if not TRegEx.IsMatch(sightings[i].species,'aurota') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    if not TRegEx.IsMatch(sightings[i].subspecies,'aurota') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    if not TRegEx.IsMatch(DateToStr(sightings[i].rec_date),'\d{1,2}\s\d?...?\s\d{4}') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    if not TRegEx.IsMatch(sightings[i].province,'[a-zA-Z -]{7,20}') then //note the literal space in RegEx
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    if not TRegEx.IsMatch(sightings[i].latitude,'^-?[0-9]{0,3},\d{1,13}$') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    if not TRegEx.IsMatch(sightings[i].longitude,'^-?[0-9]{0,3},\d{1,13}$') then
    begin
      errorID := sightings[i].rec_nr;
      errorFound := true;
      break;
    end;

    if errorFound then
      ShowMessage('Possible error found in record ' + errorID);

  end;
end;

procedure TForm1.btnBestClick(Sender: TObject);
var
  i, value, provinceTotal, monthTotal, yearTotal: integer;
  key, bestProvince, bestYear, bestMonth: string;
  provinceDict, yearDict, MonthDict: TDictionary<string, Integer>;
begin
  //Exit the method if data has not been loaded
  if sightings = nil then
    exit;

  provinceDict := TDictionary<string, Integer>.Create;
  yearDict := TDictionary<string, Integer>.Create;
  monthDict := TDictionary<string, Integer>.Create;

  for i := 0 to sightings.Count - 1 do
  begin

    if not provinceDict.ContainsKey(sightings[i].province) then
      provinceDict.Add(sightings[i].province, 1)
    else
    begin
      key := sightings[i].province;
      provinceDict.TryGetValue(key, value);

      provinceDict.AddOrSetValue(key, value + 1);
    end;

    if not monthDict.ContainsKey(DateUtils.MonthOf(sightings[i].rec_date).ToString) then
      monthDict.Add(DateUtils.MonthOf(sightings[i].rec_date).ToString, 1)
    else
    begin
      key := DateUtils.MonthOf(sightings[i].rec_date).ToString;
      monthDict.TryGetValue(key, value);

      monthDict.AddOrSetValue(key, value + 1);
    end;

    if not yearDict.ContainsKey(DateUtils.YearOf(sightings[i].rec_date).ToString) then
      yearDict.Add(DateUtils.yearOf(sightings[i].rec_date).ToString, 1)
    else
    begin
      key := DateUtils.YearOf(sightings[i].rec_date).ToString;
      yearDict.TryGetValue(key, value);

      yearDict.AddOrSetValue(key, value + 1);
    end;
  end;

  for key in provinceDict.Keys do
  begin
    if value < provinceDict.Items[key] then
      provinceTotal := provinceDict.Items[key];
      bestProvince := key;
  end;

  for key in monthDict.Keys do
  begin
    if value < monthDict.Items[key] then
      monthTotal := monthDict.Items[key];
      bestMonth := key;
  end;

  for key in yearDict.Keys do
  begin
    if value < yearDict.Items[key] then
      yearTotal := yearDict.Items[key];
      bestYear := key;
  end;

  Showmessage('Province with highest sightings: ' + bestProvince + ' with ' +
              IntToStr(provinceTotal) + ' total.' + sLineBreak +
              'Month with the most sightings: ' + bestMonth + ' with ' +
              IntToStr(monthTotal) + ' total.' + sLineBreak +
              'year with the most sightings: ' + bestYear + ' with ' +
              IntToStr(yearTotal) + ' total.');
end;

procedure TForm1.btnExitClick(Sender: TObject);
begin
 Form1.Close;
end;

end.
