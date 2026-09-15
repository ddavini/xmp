unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Bass;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    GroupBox2: TGroupBox;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
    cthread: DWORD;
    chan: HSTREAM;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

const
  urls: array[0..9] of String = (
	'http://64.236.34.196/stream/1048','http://205.188.234.129:8024',
	'http://64.236.34.97/stream/1006','http://206.98.167.99:8406',
	'http://160.79.1.141:8000','http://streams.riverhosting.net:8012',
	'http://205.188.234.4:8016','http://205.188.234.4:8014',
	'http://207.200.96.225:8010','http://64.202.98.91:8082'
  );

{$R *.dfm}

{ display error messages }
procedure Error(es: String);
begin
  MessageBox(Form1.Handle, PChar(es + #13#10 + '(error code: ' + IntToStr(BASS_ErrorGetCode) + ')'), 'Error', 0);
end;

{ update stream title from metadata }
procedure DoMeta(meta: PChar);
var
  p: Integer;
begin
  if (meta <> nil) then
  begin
    p := Pos('StreamTitle=', meta);
    if (p = 0) then
      Exit;
    p := p + 13;
    Form1.Label3.Caption := Copy(meta, p, Pos(';', meta) - p - 1);
  end;
end;

procedure MetaSync(handle: HSYNC; channel, data, user: DWORD); stdcall;
begin
  DoMeta(PChar(data));
end;

procedure StatusProc(buffer: Pointer; len, user: DWORD); stdcall;
begin
  if (buffer <> nil) and (len = 0) then
    Form1.Label5.Caption := PChar(buffer); // display connection status
end;

function OpenURL(url: PChar): Integer;
var
  icy: PChar;
begin
  Result := 0;
  BASS_StreamFree(Form1.chan); // close old stream
  Form1.Label4.Caption := 'connecting...';
  Form1.Label3.Caption := '';
  Form1.Label5.Caption := '';
  Form1.chan := BASS_StreamCreateURL(url, 0, BASS_STREAM_STATUS, @StatusProc, 0);
  if (Form1.chan = 0) then
  begin
    Form1.Label4.Caption := 'not playing';
    Error('Can''t play the stream');
  end
  else
  begin
    // get the broadcast name and bitrate
    icy := BASS_StreamGetTags(Form1.chan, BASS_TAG_ICY);
    if (icy <> nil) then
      while (icy^ <> #0) do
      begin
        if (Copy(icy, 1, 9) = 'icy-name:') then
          Form1.Label4.Caption := Copy(icy, 11, Length(icy) - 10)
        else if (Copy(icy, 1, 7) = 'icy-br:') then
          Form1.Label5.Caption := 'bitrate: ' + Copy(icy, 9, Length(icy) - 8);
        icy := icy + Length(icy) + 1;
      end;
    // get the stream title and set sync for subsequent titles
    DoMeta(BASS_StreamGetTags(Form1.chan,BASS_TAG_META));
    BASS_ChannelSetSync(Form1.chan,BASS_SYNC_META,0,@MetaSync,0);
    // play it!
    BASS_ChannelPlay(Form1.chan,FALSE);
  end;
  Form1.cthread := 0;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  // check that BASS 2.2 was loaded
  if (BASS_GetVersion <> DWORD(MAKELONG(2,2))) then
  begin
    MessageBox(0,'BASS version 2.2 was not loaded','Incorrect BASS.DLL',0);
    Halt;
  end;
  if (not BASS_Init(-1,44100,0,Form1.Handle,NIL)) then
  begin
    Error('Can''t initialize device');
    Halt;
  end;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  BASS_Free;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  ThreadId: Cardinal;
begin
  if (cthread <> 0) then
    MessageBeep(0)
  else
    cthread := BeginThread(nil, 0, @OpenURL, PChar(urls[TButton(Sender).Tag]), 0, ThreadId);
end;

end.
