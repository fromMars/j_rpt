/* ES61 Mitre */

Product     := Param.Product;
Profile     := Param.Profile;
Position    := Param.Workoffset;
Direction   := Param.B;             /* 0 = inside, 1 = outside */
ProfileType := Param.C;             /* 0 = vent, 1 = frame, 2 = T-mullion */

dx:=0;
p_height:=0;
if Product="0011441" then
{
 dx:=4;
 p_height:=61;
}
else if Product="0011442" then
{
 dx:=22;
 p_height:=70;
} 
/*7.5mm*/
Tool.Name := 'STD_HOLE';
Tool.Profilav:=False;
Tool.X:=Position+dx;
if Product="0011441" then
 Tool.Y:=8.5;
else if Product="0011442" then
 Tool.Y:=50;
else
 msgbox("0011441 or 0011442 not found!"); 
Tool.Z:=2.5;
Tool.P1X:=0;
Tool.P1Y:=0;
Tool.P1Z:=61;
Tool.P2X:=100;
Tool.P2Y:=0;
Tool.P2Z:=61;
Tool.P3X:=100;
Tool.P3Y:=0;
Tool.P3Z:=0;
Tool.Param['PAR1']:=dx-0.5;
Tool.Param['PAR2']:=7.6;
Tool.Tool:=2;
Tool.Apply();
