unit Splash;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls;

type
  TfrmSplashScreen = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    procedure Timer1Timer(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    Completed: Boolean;
  end;

var
  frmSplashScreen: TfrmSplashScreen;

implementation

{$R *.dfm}

procedure TfrmSplashScreen.FormShow(Sender: TObject);
begin
    OnShow := nil;
    Completed := False;
    Timer1.Interval := 1000; // 3s minimum time to show splash screen
    Timer1.Enabled := True;
end;

procedure TfrmSplashScreen.Timer1Timer(Sender: TObject);
begin
  Timer1.Enabled := False;
  Completed := True;
end;

end.
