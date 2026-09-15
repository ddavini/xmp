unit frmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, ExtCtrls;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    ListBox1: TListBox;
    Button1: TButton;
    Label1: TLabel;
    TrackBar1: TTrackBar;
    Timer1: TTimer;
    OpenDialog1: TOpenDialog;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure TrackBar1Change(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses Bass;

var
  chan: HSTREAM;

{$R *.dfm}

// display error messages
procedure Error(es: String);
begin
  MessageBox(Application.Handle, PChar(es + #13#10 + '(error code: ' + IntToStr(BASS_ErrorGetCode()) + ')'), PChar('Error'), MB_OK);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  sRec: TSearchRec;
  i: Integer;
begin
	// Check that BASS 2.2 was loaded
	if (BASS_GetVersion <> DWORD(MAKELONG(2,2))) then
  begin
		MessageBox(0,'BASS version 2.2 was not loaded','Incorrect BASS.DLL',0);
		Halt;
	end;

  // initialize default output device
  if (not BASS_Init(-1,44100,0,Handle,nil)) then
  begin
    Error('Can''t initialize device');
    Halt;
  end;
  // look for plugins (in the current directory)
  i := FindFirst('bass*.dll', faAnyFile, sRec);
  while i = 0 do
  begin
    if (BASS_PluginLoad(pchar(sRec.Name)) <> 0) then // plugin loaded...
      ListBox1.Items.Add(sRec.Name);
    i := FindNext(sRec);
  end;
  FindClose(sRec);

  if (ListBox1.Items.Count = 0) then // no plugins...
  begin
    ListBox1.Items.Add('no plugins - ');
    ListBox1.Items.Add('visit the BASS');
    ListBox1.Items.Add('webpage to get some');
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  BASS_Free();
  BASS_PluginFree(0);
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  TrackBar1.Position := Trunc(BASS_ChannelBytes2Seconds(chan,BASS_ChannelGetPosition(chan))); // update position
end;

procedure TForm1.TrackBar1Change(Sender: TObject);
begin
  // set the position
  BASS_ChannelSetPosition(chan,BASS_ChannelSeconds2Bytes(chan,TrackBar1.Position));
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  info: BASS_CHANNELINFO;
begin
  if (not OpenDialog1.Execute) then
    Exit;

  BASS_StreamFree(chan);
  chan := BASS_StreamCreateFile(False, pchar(OpenDialog1.FileName), 0, 0, BASS_SAMPLE_LOOP);
  if (chan = 0) then
  begin // it ain't playable
    Button1.Caption := 'Click here to open a file...';
    Label1.Caption := '';
    Error('Can''t play the file');
    Exit;
  end;

  BASS_ChannelGetInfo(chan, info);
  Button1.Caption := OpenDialog1.FileName;
  Label1.Caption := 'channel type = ' + IntToHex(info.ctype, 5);
  TrackBar1.Max := Trunc(BASS_ChannelBytes2Seconds(chan, BASS_ChannelGetLength(chan)));
  BASS_ChannelPlay(chan, False);
end;

end.
