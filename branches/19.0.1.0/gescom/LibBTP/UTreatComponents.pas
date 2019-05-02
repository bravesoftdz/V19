unit UTreatComponents;

interface
uses classes,TypInfo,variants,SysUtils;

procedure ListComponentProperties(Component: TComponent; var Strings: TStringlist);

implementation

procedure ListComponentProperties(Component: TComponent; var Strings: TStringList);
var
  Count, Size, I: Integer;
  List: PPropList;
  PropInfo: PPropInfo;
  PropOrEvent, PropValue: string;
  ss : string;
begin
  Count := GetPropList(Component.ClassInfo, tkAny, nil);
  Size  := Count * SizeOf(Pointer);
  GetMem(List, Size);
  try
    Count := GetPropList(Component.ClassInfo, tkAny, List);
    for I := 0 to Count - 1 do
    begin
      PropInfo := List^[I];
      if PropInfo^.PropType^.Kind in tkMethods then
        PropOrEvent := 'Event'
      else
        PropOrEvent := 'Property';
      PropValue := VarToStr(GetPropValue(Component, PropInfo^.Name));
      ss := Format('[%s] %s: %s = %s', [PropOrEvent, PropInfo^.Name,PropInfo^.PropType^.Name, PropValue]);
      if PropInfo^.PropType^.Kind in tkMethods then
      begin
        Strings.Add(SS);
      end else Strings.Add(SS);
    end;

  finally
    FreeMem(List);
  end;
end;

end.
