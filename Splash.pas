unit Splash;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.jpeg,
  Vcl.ExtCtrls;

type
  TFormSplash = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    StartupTimer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure FormPaint(Sender: TObject);
    procedure StartupTimerTimer(Sender: TObject);
  private
    { Private declarations }
    FInitialized: Boolean;
    procedure LoadMainForm;
  public
    { Public declarations }
  end;

var
  FormSplash: TFormSplash;

implementation

{$R *.dfm}

uses About, Main;

procedure TFormSplash.FormCreate(Sender: TObject);
begin
  StartupTimer.Enabled := False;
  StartupTimer.Interval := 1000;
  FInitialized := True;
  Show;
end;

procedure TFormSplash.FormPaint(Sender: TObject);
begin
   StartupTimer.Enabled := not FInitialized;
end;

procedure TFormSplash.LoadMainForm;
begin
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TAboutBox, AboutBox);
  Application.Run;
  FormSplash.Close;
end;

procedure TFormSplash.StartupTimerTimer(Sender: TObject);
begin
  StartupTimer.Enabled := False;
  if not FInitialized then begin
    FInitialized := True;
    LoadMainForm;
  end
end;

end.
