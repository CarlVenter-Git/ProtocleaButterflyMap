unit ProtocleaButterflyMap;

interface

uses
  System.SysUtils, System.Types, System.UITypes, System.Classes, System.Variants,
  FMX.Types, FMX.Controls, FMX.Forms, FMX.Graphics, FMX.Dialogs,
  FMX.Controls.Presentation, FMX.StdCtrls, ComObj;

type
  TForm1 = class(TForm)
    btnLoad: TButton;
    OpenDialog1: TOpenDialog;
    Panel1: TPanel;
    lblPath: TLabel;
    btnPlotPoints: TButton;
    procedure btnLoadClick(Sender: TObject);
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
  sightings: array of sightingRecord;

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

      SetLength(sightings, numRows - 1);

      for x := 2 to numRows do //Start at index 2 to skip the heading
      begin
        lblPath.Text := 'Loading ' + IntToStr(x - 1) + '/' + IntToStr(numRows);
        y := x - 2;//So that the array index starts at 0

        //genus,species and subspecies is rather redundant in this case but still useful to have in the record for use at a later stage if the data includes more
        sightings[y].rec_nr := sheet.Cells.Item[x, 1].Value;
        sightings[y].genus := sheet.Cells.Item[x, 2].Value;
        sightings[y].species := sheet.Cells.Item[x, 3].Value;
        sightings[y].subspecies := sheet.Cells.Item[x, 4].Value;
        sightings[y].rec_date := sheet.Cells.Item[x, 5].Value;
        sightings[y].province := sheet.Cells.Item[x, 6].Value;
        sightings[y].latitude := sheet.Cells.Item[x, 7].Value;
        sightings[y].longitude := sheet.Cells.Item[x, 8].Value;
      end;

      lblPath.Text := selectedFilePath + ' Successfully Loaded';

    finally
      excel.Quit;
      excel := Unassigned;
    end;
  end;
end;

end.
