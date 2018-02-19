unit VersionInfo;

interface

uses
  Windows, SysUtils;

type
  TLangAndCP = record
    wLanguage: word;
    wCodePage: word;
  end;
  PLangAndCP = ^TLangAndCP;

  RAppInfo = record
    InfoStr: String;
    Value: String;
  end;
  rRAppInfo = array of RAppInfo;
  
  {$M+}
  TVersionInfo = class
  private
    { Private declarations }
    FDefaultValue: String;
    FLang: PLangAndCP;
    FBuf: PChar;
    FAppInfo: rRAppInfo;
    function QueryValue(pInfo: Integer): String;
    procedure ClearAll;
  public
    { Public declarations }
    constructor Create;
    destructor Destroy; override;
  published
    { Published declarations }
    property DefaultValue: String read FDefaultValue write FDefaultValue stored False;
    procedure GetInfo(FName: String);
    function GetInfoString(pInfo: Integer): String;
    function GetInfoValue(pInfo: Integer): String;
  end;

const
  cVICompanyName = Integer(0);
  cVIFileDescription = Integer(1);
  cVIFileVersion = Integer(2);
  cVIInternalName = Integer(3);
  cVILegalCopyright = Integer(4);
  cVILegalTradeMarks = Integer(5);
  cVIOriginalFilename = Integer(6);
  cVIProductName = Integer(7);
  cVIProductVersion = Integer(8);
  cVIComments = Integer(9);
  cVIPrivateBuild = Integer(10);
  cVISpecialBuild = Integer(11);
  //
  cVIMajorVersion = Integer(12);
  cVIMinorVersion = Integer(13);
  cVIRelease = Integer(14);
  cVIBuild = Integer(15);

implementation

//--------------------------------------------------------------------------
constructor TVersionInfo.Create;
begin
  FDefaultValue := '<info not available>';

  SetLength(FAppInfo, 16);

  FAppInfo[cVICompanyName].InfoStr := 'CompanyName';
  FAppInfo[cVIFileDescription].InfoStr := 'FileDescription';
  FAppInfo[cVIFileVersion].InfoStr := 'FileVersion';
  FAppInfo[cVIInternalName].InfoStr := 'InternalName';
  FAppInfo[cVILegalCopyright].InfoStr := 'LegalCopyright';
  FAppInfo[cVILegalTradeMarks].InfoStr := 'LegalTradeMarks';
  FAppInfo[cVIOriginalFilename].InfoStr := 'OriginalFilename';
  FAppInfo[cVIProductName].InfoStr := 'ProductName';
  FAppInfo[cVIProductVersion].InfoStr := 'ProductVersion';
  FAppInfo[cVIComments].InfoStr := 'Comments';
  FAppInfo[cVIPrivateBuild].InfoStr := 'PrivateBuild';
  FAppInfo[cVISpecialBuild].InfoStr := 'SpecialBuild';
  FAppInfo[cVIMajorVersion].InfoStr := 'MajorVersion';
  FAppInfo[cVIMinorVersion].InfoStr := 'MinorVersion';
  FAppInfo[cVIRelease].InfoStr := 'Release';
  FAppInfo[cVIBuild].InfoStr := 'Build';

  ClearAll;
end;

//--------------------------------------------------------------------------
destructor TVersionInfo.Destroy;
begin
  inherited Destroy;
end;

//--------------------------------------------------------------------------
function TVersionInfo.GetInfoString(pInfo: Integer): String;
begin
  try
    Result := FAppInfo[pInfo].InfoStr;
  except
    Result := '';
  end;
end;

//--------------------------------------------------------------------------
function TVersionInfo.GetInfoValue(pInfo: Integer): String;
begin
  try
    Result := FAppInfo[pInfo].Value;
  except
    Result := '';
  end;
end;

//--------------------------------------------------------------------------
function TVersionInfo.QueryValue(pInfo: Integer): String;
var
  Value: PChar;
  SubBlock: String;
  Len: Cardinal;
begin
  SubBlock := Format('\\StringFileInfo\\%.4x%.4x\\%s',[
    FLang^.wLanguage,FLang^.wCodePage,
    FAppInfo[pInfo].InfoStr
  ]);
  VerQueryValue(FBuf,PChar(SubBlock),Pointer(Value),Len);
  if Len > 0 then
    Result := Trim(String(Value))
  else
    Result := '';
end;

//--------------------------------------------------------------------------
procedure TVersionInfo.ClearAll;
var
  i: Integer;
begin
  for i := 0 to Length(FAppInfo)-1 do
    FAppInfo[i].Value := FDefaultValue;
end;

//--------------------------------------------------------------------------
procedure TVersionInfo.GetInfo(FName: String);
var
  pp, i, p: Integer;
  ZValue, LangLen: Cardinal;
  sSep, s: String;
begin
  ClearAll;
  try
    i := GetFileVersionInfoSize(PChar(FName),ZValue);
    if i > 0 then
    begin
      FBuf := AllocMem(i);
      try
        GetFileVersionInfo(PChar(FName),0,i,FBuf);
        VerQueryValue(FBuf,PChar('\\VarFileInfo\\Translation'),Pointer(FLang),LangLen);
        for p := 0 to 11 do
        begin
          try
            FAppInfo[p].Value := QueryValue(p);
          except
            FAppInfo[p].Value := FDefaultValue;
          end;
        end;

        // Seperate FileVersion values
        sSep := '';
        s := FAppInfo[cVIFileVersion].Value;
        if s <> '' then
        begin
          if Pos('.',s) <> 0 then
            sSep := '.'   // Separator as '.'
          else if Pos(',',s) <> 0 then
            sSep := ',';  // Separator as ','
        end;
    	if sSep <> '' then
        begin
          try
            // Version as major.minor.release.build or major,minor,release,build
            for p := cVIMajorVersion to cVIRelease do
            begin
              pp := Pos(sSep,s);
              FAppInfo[p].Value := Copy(s,1,pp-1);
	      Delete(s,1,pp);
            end;
            FAppInfo[cVIBuild].Value := s;
          except
	    FAppInfo[cVIMajorVersion].Value := FDefaultValue;
	    FAppInfo[cVIMinorVersion].Value := FDefaultValue;
	    FAppInfo[cVIRelease].Value := FDefaultValue;
	    FAppInfo[cVIBuild].Value := FDefaultValue;
          end;
	end;
      except
        ClearAll;
      end;
      FreeMem(FBuf,i);
    end;
  except
    ClearAll;
  end;
end;

end.
