unit ProtocleaButterflyMap;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, ComObj, FMX.Objects, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdHTTP, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, FMX.ListBox, DateUtils;

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
    Button1: TButton;
    Button2: TButton;
    procedure btnLoadClick(Sender: TObject);
    procedure btnPlotPointsClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

//I am unsure if this is the correct place for this record type to be declared
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
  sightings: array of sightingRecord;//Not a fan of globals, but I'm not sure how to abstract effectively yet

implementation
{$R *.fmx}

procedure TForm1.btnLoadClick(Sender: TObject);
var
  excel, book, sheet, range: OLEVariant;
  selectedFilePath: string;
  provinces, yearStrings: array of string;
  years: array of TDateTime;
  match: boolean;
  x, y, i, j, yearIndex, numRows, numColumns: integer;

begin

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
      numColumns := range.Columns.Count;
      yearIndex := 0;

      //Ignore the first line to make array correct size
      SetLength(sightings, numRows - 1);
      Setlength(years, 100); //I am making an assumption here, this would cause issues if we are looking at more than 50 years of data
      Setlength(yearStrings, 100);
      SetLength(provinces, 9); //I am more comfortable making this assumption

      for x := 2 to numRows do //Start at index 2 to skip the heading
      begin
        lblPath.Text := 'Loading ' + IntToStr(x - 1) + '/' + IntToStr(numRows);
        y := x - 2;//So that the array index starts at 0

        //genus,species and subspecies is rather redundant in this case but still useful to have in the record for use at a later stage if the data ever includes more
        sightings[y].rec_nr := sheet.Cells.Item[x, 1];
        sightings[y].genus := sheet.Cells.Item[x, 2];
        sightings[y].species := sheet.Cells.Item[x, 3];
        sightings[y].subspecies := sheet.Cells.Item[x, 4];
        sightings[y].rec_date := sheet.Cells.Item[x, 5];
        sightings[y].province := sheet.Cells.Item[x, 6];
        sightings[y].latitude := sheet.Cells.Item[x, 7];
        sightings[y].longitude := sheet.Cells.Item[x, 8];

        match := false;

        //This needs to be extracted, why is there not a Contains method for arrays??
        for i := Low(years) to High(years) do
        begin
          if DateUtils.YearOf(years[i]) = DateUtils.YearOf(sightings[y].rec_date) then
          begin
            match := true;
            break;
          end;
        end;

        if match = false then
        begin
          yearStrings[yearIndex] :=  DateUtils.YearOf(sightings[y].rec_date).ToString;
          yearIndex := yearIndex + 1;
        end;

      end;

      lblPath.Text := selectedFilePath + ' Successfully Loaded';

      for j := 0 to High(yearStrings) do
      begin
        if yearStrings[j] <> '' then
          cmbYear.Items.Add(yearStrings[j]);
      end;

    finally
      excel.Quit;
      excel := Unassigned;
    end;
  end;

end;

procedure TForm1.btnPlotPointsClick(Sender: TObject);
const
  prefix = 'https://maps.googleapis.com/maps/api/staticmap?';
var
  ms: TMemoryStream;
  bitmap: TBitmap;
  url: string;
  x, y: single;
begin
  x := Panel1.Size.Width;
  y := Panel1.Size.Height;

  url := prefix + 'center=Cape+Town,Western+Cape&zoom=20&size=' +
                  x.ToString + 'x' + y.ToString + '&maptype=roadmap&'+
                  'markers=color:red|label:C|40.718217,-73.998284%27&'+
                  'key=AIzaSyBo2oV0QLZwhOsLjeV08m04nA4xlRd0PxA';

  ms := TMemoryStream.Create;
  bitmap := TBitmap.Create;

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

procedure TForm1.Button2Click(Sender: TObject);
begin
 Form1.Close;
end;

end.
