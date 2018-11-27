unit SplashScreen;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Imaging.jpeg,
  Vcl.StdCtrls;

type
  TSplashScn = class(TForm)
    Image1: TImage;
    Timer1: TTimer;
    StaticText1: TStaticText;
    StaticText2: TStaticText;
    StaticText3: TStaticText;
    stReconnect: TStaticText;
    procedure FormShow(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Completed: Boolean;
  end;

var
  SplashScn: TSplashScn;

implementation

{$R *.dfm}

procedure TSplashScn.FormShow(Sender: TObject);
begin
   Timer1.Enabled := True;
   Completed := False
end;

procedure TSplashScn.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Completed := True;
end;

end.
