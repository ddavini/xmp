unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Bass;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Button1: TButton;
    Button2: TButton;
    OpenDialog1: TOpenDialog;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    outdev: array[0..1] of DWORD;
    chan: array[0..1] of HSTREAM;
  end;

var
  Form1: TForm1;

implementation

uses Unit2;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  if (BASS_GetVersion <> DWORD(MAKELONG(2,2))) then
  begin
		MessageBox(0, 'BASS version 2.2 was not loaded', 'Incorrect BASS.DLL', 0);
    Halt;
  end;

  Application.CreateForm(TForm2, Form2);
  Form2.Caption := 'select output device #1';
  outdev[0] := Form2.ShowModal;
  Form2.Free;
  Application.CreateForm(TForm2, Form2);
  Form2.Caption := 'select output device #2';
  outdev[1] := Form2.ShowModal;
  Form2.Free;

  if (not BASS_Init(outdev[0], 44100, 0, Handle, nil)) then
  begin
    MessageBox(0, PChar('Cant''t initialize device 1' + #13#10 + 'Error #' + IntToStr(BASS_ErrorGetCode)), 'Error', MB_ICONERROR);
    Halt;
  end;
  if (not BASS_Init(outdev[1], 44100, 0, Handle, nil)) then
  begin
    MessageBox(0, PChar('Cant''t initialize device 2' + #13#10 + 'Error #' + IntToStr(BASS_ErrorGetCode)), 'Error', MB_ICONERROR);
    Halt;
  end;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  BASS_SetDevice(outdev[0]);
  BASS_Free;
  BASS_SetDevice(outdev[1]);
  BASS_Free;

  Action := caFree;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  curdev: Integer;
begin
  if (not OpenDialog1.Execute) then
    Exit;

  curdev := TButton(Sender).Tag;

  BASS_StreamFree(chan[curdev]);
  BASS_SetDevice(outdev[curdev]);

  chan[curdev] := BASS_StreamCreateFile(False, PChar(OpenDialog1.FileName), 0, 0, BASS_SAMPLE_LOOP);
  if (chan[curdev] = 0) then
  begin
    TButton(Sender).Caption := 'click here to open a file...';
    MessageBox(0, 'Can''t play the file', 'Error', MB_ICONERROR);
    Exit;
  end;

  TButton(Sender).Caption := OpenDialog1.FileName;
  BASS_ChannelPlay(chan[curdev], False);
end;

end.
