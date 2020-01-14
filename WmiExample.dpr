program WmiExample;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  System.Variants,
  uWmi in 'uWmi.pas';

{------------------------------------------------------------------------------------------------------------}
procedure ShowMonitorConnParam(AItem: OLEVariant);
begin
  WriteLn(#9 + Format('%s: %-15s', ['Name', AItem.InstanceName]));
end;

{------------------------------------------------------------------------------------------------------------}
procedure ShowDesktopMonitor(AItem: OLEVariant);
begin
  Write(#9 + Format('%s: %-15s' + #9, ['PNPDeviceID', AItem.PNPDeviceID]));
  WriteLn(Format('%s: %-15s', ['DeviceID', AItem.DeviceID]));
end;

{------------------------------------------------------------------------------------------------------------}
procedure ShowVideoController(AItem: OLEVariant);
begin
  WriteLn(#9 + Format('%s: %-15s', ['Caption', AItem.Caption]));
end;

{------------------------------------------------------------------------------------------------------------}
procedure ShowMonitorDisplayParam(AItem: OLEVariant);
begin
  Write(#9 + Format('%s: %-15s' + #9, ['PNPDeviceID', AItem.InstanceName]));
  WriteLn(Format('%s: %-15s', ['Active', AItem.Active]));
end;

{------------------------------------------------------------------------------------------------------------}
procedure ShowMonitorId(AItem: OLEVariant);
begin
  WriteLn(#9 + Format('%s: %-15s', ['Name', AItem.InstanceName]));
end;

var
  S: string;
begin
  try
    WriteLn('Monitor Connection Params:');
    TWmi.Select(ShowMonitorConnParam, nsWmi, 'Select * from WmiMonitorConnectionParams');

    WriteLn(sLineBreak + 'Desktop Monitors:');
    TWmi.Select(ShowDesktopMonitor, nsCmi, 'Select * from CIM_DesktopMonitor');

    WriteLn(sLineBreak + 'Video Controllers:');
    TWmi.Select(ShowVideoController, nsCmi,'Select * from CIM_VideoController');

    WriteLn(sLinebreak + 'Monitor Basic Display Params:');
    TWmi.Select(ShowMonitorDisplayParam, nsWmi, 'Select * from WmiMonitorBasicDisplayParams');

    WriteLn(sLinebreak + 'Monitor Ids:');
    TWmi.Select(ShowMonitorId, nsWmi, 'Select * from WMIMonitorID');

    WriteLn(sLineBreak + 'Press any key to continuie...');
    Readln(S);
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
