unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs,VBIDE_TLB,Word_TLB,Office_TLB,
  Vcl.StdCtrls, Vcl.ComCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    ComboBox1: TComboBox;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Edit9: TEdit;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Edit10: TEdit;
    Edit11: TEdit;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Edit12: TEdit;
    Label19: TLabel;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  wold: WordApplication;
  Doc: WordDocument;
implementation

{$R *.dfm}



procedure TForm1.Button1Click(Sender: TObject);
var
  WordApp: WordApplication;
  Docs: Documents;
  Doc: WordDocument;
  Pars: Paragraphs;
  Par: Paragraph;
  D: OleVariant;
begin
  WordApp := CoWordApplication.Create;
  WordApp.Visible := True;

  Docs := WordApp.Documents;
  Doc := Docs.Add('Normal', False, EmptyParam, True);

  Doc := (WordApp.Documents.Item(1) as WordDocument);
  Doc.Paragraphs.Item(1).Format.LeftIndent:=WordApp.CentimetersToPoints(8) ;
  Doc.Paragraphs.Item(1).Format.SpaceAfter:=12;
  Doc.Paragraphs.Item(1).Range.Text :=
    'Ректору ФГАОУ ВПО СФУ '
    +#13+'Ваганову Е.А.'+ComboBox1.Text
    +#13+'Абитуриента '+Edit1.Text
    +#13+Edit12.Text
    +#13
    +#13
    +#13+'Заявление'
    +#13+'Прошу предоставить мне место в общежитии, так как являюсь иногородним студентом. Обучаюсь на бюджетной / платной основе (нужное подчеркнуть).'
    +#13+'Дата рождения  '
    +#13+DateToStr(DateTimePicker1.DateTime)
    +#13+'Паспортные данные: '
    +#13+Edit2.Text+Edit3.Text
    +#13+'Родители проживают: '
    +#13+Edit4.Text+Edit5.Text
    +#13+'Состав семьи:   '
    +#13+'Отец:  ' +Edit6.Text
    +#13+Edit7.Text
    +#13+'Мать:  '+Edit8.Text
    +#13+Edit9.Text
    +#13+'Брат, сестра:  '+ Edit10.Text
    +#13+ Edit11.Text
    +#13+'ОБЯЗУЮСЬ: '
    +#13+'1. Выполнять правила внутреннего распорядка в общежитии. '
    +#13+'2. Выполнять Федеральный закон № 87-ФЗ «Об ограничении курения», № 15-ФЗ «Об охране здоровья граждан от воздействия окружающего табачного дыма и последствий потреблений табака». '
    +#13+'3. Выполнять требования органов студенческого самоуправления.'
    +#13
    +#13+'Дата: '+DateToStr(DateTimePicker1.DateTime)+#09+#09+#09+#09+#09+#09+#09+'Подпись __________'
    ;

  Doc.Range(192,212).Font.Bold:=1;
  Doc.Paragraphs.Item(1).Range.Font.Name := 'Times New Roman';
  Doc.Paragraphs.Item(1).Range.Font.Size := 12;
  Doc.Paragraphs.Item(1).Range.Font.Color := clBlack;
  Doc.Paragraphs.Item(1).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(1).Range.Font.Italic := 0;
  Doc.Paragraphs.Item(1).Alignment:= wdAlignParagraphLeft;

  WordApp.Selection.WholeStory;
  WordApp.Selection.ParagraphFormat.LineSpacing := WordApp.LinesToPoints(0.9);
  WordApp.Selection.Font.Name:= 'Times New Roman';
  WordApp.Selection.Font.Size:= 12;
  WordApp.Selection.Font.Color:=clBlack;

  Doc.Paragraphs.Item(7).Format.LeftIndent:=WordApp.CentimetersToPoints(0) ;
  Doc.Paragraphs.Item(7).Range.Font.Size := 14 ;
  Doc.Paragraphs.Item(7).Format.Alignment:=wdAlignParagraphCenter;
  Doc.Paragraphs.Item(7).Range.Font.bold:=1;


  Doc.Paragraphs.Item(22).Range.Font.Size := 14 ;
  Doc.Paragraphs.Item(22).Range.Font.bold:=1;
  Doc.Paragraphs.Item(22).Range.Font.Underline:=1;
  Doc.Paragraphs.Item(22).Range.Font.UnderlineColor:=clBlack;

  doc.Paragraphs.Item(8).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(8).Format.FirstLineIndent := WordApp.CentimetersToPoints(0.5);
  Doc.Paragraphs.Item(9).Range.Font.bold:=1;
  Doc.Paragraphs.Item(11).Range.Font.bold:=1;
  Doc.Paragraphs.Item(13).Range.Font.bold:=1;
  doc.Paragraphs.Item(9).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(10).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(11).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(12).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(13).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(14).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(15).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(16).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(17).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(18).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(19).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  Doc.Paragraphs.Item(20).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(21).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(22).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  doc.Paragraphs.Item(24).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  Doc.Paragraphs.Item(23).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  Doc.Paragraphs.Item(25).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  Doc.Paragraphs.Item(26).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  Doc.Paragraphs.Item(27).Format.LeftIndent := WordApp.CentimetersToPoints(0);
  Doc.Paragraphs.Item(28).Format.LeftIndent := WordApp.CentimetersToPoints(0);
 WordApp.Selection.Paragraphs.Item(22).SpaceAfter:=0;
 WordApp.Selection.Paragraphs.Item(23).SpaceAfter:=0;
 WordApp.Selection.Paragraphs.Item(24).SpaceAfter:=0;

end;




end.
