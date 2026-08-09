program PPL_Sim;

uses
  Forms,
  IWMain,
  UntEntrada in 'UntEntrada.pas' {IWForm1: TIWAppForm},
  ServerController in 'ServerController.pas' {IWServerController: TIWServerControllerBase},
  UserSessionUnit in 'UserSessionUnit.pas' {IWUserSession: TIWUserSessionBase},
  lpsolve51 in 'lpsolve51.pas',
  uAPI_SolveLP in 'uAPI_SolveLP.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TformIWMain, formIWMain);
  Application.Run;
end.
