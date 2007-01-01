{*
 * PathParser: Replaces path variables in the format $(VARIABLE) and retrieves the path of Windows' special folders (desktop, application data, templates, programs, personal files, favorites, startup, recent, send to, start menu, fonts, history, cookies, internet cache, common favorites, common desktop, common startup, common programs, common start menu, program files, temporary folder, windows, system)
 * Jonas Raoni Soares da Silva <http://raoni.org>
 * https://github.com/jonasraoni/path-parser
 *}

unit PathParser;

interface

uses
  Classes, SysUtils, TypInfo, ShlObj, ShellApi, Registry, Windows;

type
  TSpecialFolder = ( sfDesktop, sfAppData, sfTemplates, sfPrograms,
    sfPersonal, sfFavorites, sfStartup, sfRecent, sfSendTo, sfStartMenu,
    sfFonts, sfHistory, sfCookies, sfInternetCache, sfCommonFavorites,
    sfCommonDesktop, sfCommonStartup, sfCommonPrograms, sfCommonStartMenu,
    sfProgramFiles, sfTemporary, sfWindows, sfSystem );

  TSpecialFolderSet = set of TSpecialFolder;

  TPathParser = class( TStringList )
  public
    constructor Create( const UseDefaultMap: Boolean = True );
    class function GetSpecialFolder( const Name: TSpecialFolder ): string;
    function Parse( Path: string ): string;
  end;


implementation

{ TPathParser }

uses Dialogs;

function RemoveSlash( const Path: string ): string;
begin
	Result := Path;
	if Result[Length( Result )] = PathDelim then
		Delete( Result, Length( Result ), 1 );
	Result := Result;
end;

function AddSlash( const Path: string ): string;
begin
	Result := Path;
	if not IsPathDelimiter(Result, Length(Result)) then
		Result := Result + PathDelim;
end;

constructor TPathParser.Create(const UseDefaultMap: Boolean);
var
  I: TSpecialFolder;
begin
  CaseSensitive := False;
  if UseDefaultMap then begin
    for I := Low( TSpecialFolder ) to High( TSpecialFolder ) do
      Add( RemoveSlash( LowerCase( Copy( GetEnumName( TypeInfo( TSpecialFolder ),
        Ord( I ) ), 3, MAX_PATH ) ) + '=' + GetSpecialFolder( I ) ) );
    Add( RemoveSlash( Format( 'windowsvolume=%s', [ GetSpecialFolder( sfWindows )[1] ] ) ) );
  end;
end;

class function TPathParser.GetSpecialFolder(
  const Name: TSpecialFolder): string;
const
  FoldersMap: array[TSpecialFolder] of Cardinal = ( CSIDL_DESKTOP,
    CSIDL_APPDATA, CSIDL_TEMPLATES, CSIDL_PROGRAMS, CSIDL_PERSONAL,
    CSIDL_FAVORITES, CSIDL_STARTUP, CSIDL_RECENT, CSIDL_SENDTO, CSIDL_STARTMENU,
    CSIDL_FONTS, CSIDL_HISTORY, CSIDL_COOKIES, CSIDL_INTERNET_CACHE,
    CSIDL_COMMON_FAVORITES, CSIDL_COMMON_DESKTOPDIRECTORY, CSIDL_COMMON_STARTUP,
    CSIDL_COMMON_PROGRAMS, CSIDL_COMMON_STARTMENU, 0, 0, 0, 0 );
var
  Res: Bool;
  Path: array[0..MAX_PATH-1] of Char;
  Reg: TRegistry;
begin
  Result := '';
  case Name of
    sfWindows: GetWindowsDirectory( Path, MAX_PATH );
    sfTemporary: GetTempPath( MAX_PATH, Path );
    sfSystem: GetSystemDirectory( Path, MAX_PATH );
    sfProgramFiles:
    begin
      Reg := TRegistry.Create( KEY_READ );
      try
        Reg.RootKey := HKEY_LOCAL_MACHINE;
        Reg.OpenKey( 'SOFTWARE\Microsoft\Windows\CurrentVersion', False );
        Result := AddSlash( Reg.ReadString( 'ProgramFilesDir' ) );
      finally
        Reg.Free;
      end;
      Exit;
    end;
  else
    Res := ShGetSpecialFolderPath( 0, Path, FoldersMap[ Name ], False );
    if not Res then
      raise Exception.Create( ClassName + '.GetSpecialFolder: Error on ShGetSpecialFolderPath' );
  end;
  Result := AddSlash( Path );
end;

function TPathParser.Parse(Path: string): string;
var
  S: string;
  I, I2, Pos: Integer;
begin
  I := 1;
  while I <= Length( Path )-3 do
  begin
    if ( Path[I] = '$' ) and ( Path[I+1] = '(' ) then
    begin
      I2 := I + 2;
      while ( I2 <= Length( Path ) ) and ( Path[I2] <> ')' ) do
        Inc( I2 );
      if I2 > Length( Path ) then
        Break;
      S := Copy( Path, I + 2, I2 - ( I + 2 ) );
      System.Delete( Path, I, I2 - I + 1 );
      Pos := IndexOfName( S );
      if Pos > -1 then
      begin
        System.Insert( ValueFromIndex[Pos], Path, I );
        Inc( I, Length( ValueFromIndex[Pos] ) );
	  end;
    end
    else
      Inc( I );
  end;
  Result := Path;
end;

end.