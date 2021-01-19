program EgissoConverter;

uses
  Vcl.Forms,
  U_main in 'U_main.pas' {F_main};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TF_main, F_main);
  Application.Run;
end.
