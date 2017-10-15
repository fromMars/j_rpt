/* ES61 Mitre */

Product     := Param.Product;
Profile     := Param.Profile;
Position    := Param.Workoffset;
Direction   := Param.B;             /* 0 = inside, 1 = outside */
ProfileType := Param.C;             /* 0 = vent, 1 = frame, 2 = T-mullion */

/*7.5mm*/
Tool.Name := 'STD_HOLE';

if position=0 then
 Tool.X := position+65.5;           /* 59.5+6 */
else
 Tool.X:=position-65.5;
if Product="0011441" then
 Tool.Y:=11;
else if Product="0011442" then
 Tool.Y:=20;
else
 msgbox("0011441 or 0011442 not found!"); 
Tool.Z:=2.3;
Tool.P1X:=0;
Tool.P1Y:=0;
Tool.P1Z:=70;
Tool.P2X:=100;
Tool.P2Y:=0;
Tool.P2Z:=70;
Tool.P3X:=100;
Tool.P3Y:=0;
Tool.P3Z:=0;
if Product="0011441" then
 Tool.Param['PAR1']:=4;
else if Product="0011442" then
 Tool.Param['PAR1']:=21.5;
Tool.Param['PAR2']:=5;
Tool.Tool:=2;
Tool.Apply();
