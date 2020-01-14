unit uWmi;

interface

uses
  System.SysUtils, System.Variants, ActiveX, ComObj;

type
  { The root namespaces we query from }
  TWmiNamespace = (nsWmi, nsCmi, nsDefault);

  { A procedure to call on each item returned by the WMI Query }
  TVisitor = procedure(AItem: OleVariant);

  { Simple wrapper around WMI, provides a select method for demonstration }
  TWmi = class
  private
    FLocator: OLEVariant;
    FWmiService: OLEVariant;
    FCmiService: OLEVariant;
    FDefaultService: OLEVariant;

    procedure Run(AVisitor: TVisitor; ANamespace: TWmiNamespace; AQuery: string);

    class var
      FInstance: TWmi;

  public
    constructor Create;

    class constructor Create;
    class destructor Destroy;

    { Executes the specified Query and calls the visitor on each record }
    class procedure Select(AVisitor: TVisitor; ANamespace: TWmiNamespace; AQuery: string);
  end;

implementation

{ TWMIEnumerator }

{------------------------------------------------------------------------------------------------------------}
class constructor TWmi.Create;
begin
  CoInitialize(nil);
  FInstance := TWmi.Create;
end;

{------------------------------------------------------------------------------------------------------------}
constructor TWmi.Create;
begin
  FLocator := CreateOleObject('WbemScripting.SWbemLocator');
  FWmiService := FLocator.ConnectServer('localhost','root\wmi');
  FCmiService := FLocator.ConnectServer('localhost', 'root\cimv2');
  FDefaultService := FLocator.ConnectServer('localhost', 'root\default');
end;

{------------------------------------------------------------------------------------------------------------}
class destructor TWmi.Destroy;
begin
  if Assigned(FInstance) then
    FreeAndNil(FInstance);

  CoUninitialize;
end;

{------------------------------------------------------------------------------------------------------------}
procedure TWmi.Run(AVisitor: TVisitor; ANamespace: TWmiNamespace; AQuery:string);
var
  Items: OLEVariant;
  Item: OLEVariant;
  Enum: IEnumVariant;
  Value: LongWord;
begin
  case ANamespace of
    nsWmi:      Items := FWmiService.ExecQuery(AQuery);
    nsCmi:      Items := FCmiService.ExecQuery(AQuery);
    nsDefault:  Items := FDefaultService.ExecQuery(AQuery);
  end;

  Enum := IUnknown(Items._NewEnum) as IEnumVariant;

  while Enum.Next(1, Item, Value) = 0 do begin
    AVisitor(Item);
    Item := Unassigned;
  end;
end;

{------------------------------------------------------------------------------------------------------------}
class procedure TWmi.Select(AVisitor: TVisitor; ANamespace: TWmiNamespace; AQuery: string);
begin
  FInstance.Run(AVisitor, ANamespace, AQuery);
end;

end.
