unit ProtocleaButterflyMap;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, ComObj, FMX.Objects, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdHTTP, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, JPEG;

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
    procedure btnLoadClick(Sender: TObject);
    procedure btnPlotPointsClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

//I am unsure if this is the correct place for this type to be declared
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
  x, y, numRows, numColumns: integer;

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

      //Ignore the first line to make array correct size
      SetLength(sightings, numRows - 1);

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
      end;

      lblPath.Text := selectedFilePath + ' Successfully Loaded';

    finally
      excel.Quit;
      excel := Unassigned;
    end;
  end;

end;

procedure TForm1.btnPlotPointsClick(Sender: TObject);
var
    ms: TMemoryStream;

begin
    ms := TMemoryStream.Create;

    try
      idHTTP1.Get('https://maps.googleapis.com/maps/api/staticmap?center=40.714%2c%20-73.998&'+
      'zoom=12&size=400x400&key=AIzaSyC4BpjpllKKkFkhW-L89ij8u6IadYocaZM');

      ms.Seek(0, soFromBeginning);
      Image1.Bitmap.LoadFromStream(ms);

    finally
      FreeAndNil(ms);
    end;
end;

end.
