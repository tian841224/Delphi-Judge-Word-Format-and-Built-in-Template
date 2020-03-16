program Project2;



{$R *.dres}

uses
  Forms,
  Main in 'Main.pas' {CheckTest};

{$R *.res}
{$R Word.RES}
{$R Font.RES}
begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TCheckTest, CheckTest);
  Application.Run;
end.
