# Delphi-VersionInfo
Delphi class to query version info from given application.



# How to use

Include to uses
```
    uses VersionInfo;
```

Define variable
```
    var
      vi: TVersionInfo;
```

Use it
```
    vi := TVersionInfo.Create;
```

What to use as default value, if any query of info fails  (default \<info not available\>)
```    
    vi.DefaultValue := '';
```
    
This trigger the quering of file for version information.
```
    vi.GetInfo(filename);
```

Values can be any of : 
  * cVICompanyName
  * cVIFileDescription
  * cVIFileVersion
  * cVIInternalName
  * cVILegalCopyright
  * cVILegalTradeMarks
  * cVIOriginalFilename
  * cVIProductName
  * cVIProductVersion
  * cVIComments
  * cVIMajorVersion
  * cVIMinorVersion
  * cVIRelease
  * cVIBuild
 
Read a explanation for the value (returns string)
```
    desc := vi.GetInfoString(cVICompanyName);
```    
    
Get the value (returns string)
```
    val := vi.GetInfoValue(cVICompanyName);
```

You can also loop through all values
```
    for i := cVICompanyName to cVIBuild do
    begin
      desc := vi.GetInfoString(i);
      val := vi.GetInfoValue(i);
    end;
```

Make sure you free the class when done
```
    vi.Free;
```
